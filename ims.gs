/*******************************************************************************
 * INVENTORY ERP — Backend (InventoryCode.gs)
 * ---------------------------------------------------------------------------
 * Full inventory ERP running on Google Apps Script + Google Sheets:
 *   • Login / role-based access (Admin / Manager / Purchase / Sales /
 *     Storekeeper / Employee)
 *   • Product & category master, suppliers, customers, warehouses
 *   • Purchase Orders (PO) with line items + approval workflow
 *   • Sales Orders (SO) with line items + dispatch tracking
 *   • Indents (internal material request → approval → fulfillment)
 *   • Stock In (GRN) / Stock Out (Issue) / Transfer / Adjustment
 *   • Live stock ledger + running stock balance per warehouse
 *   • Centralised Approvals inbox
 *   • initializeSheets()  – Creates every sheet AND seeds rich demo data
 *   • CacheService reads, LockService writes, JSON-aware fields
 *
 * DEPLOY:
 *   1. In Apps Script → paste this as `InventoryCode.gs`
 *   2. File → New → HTML → name it **`inventory`** → paste inventory.html
 *   3. Run `initializeSheets` once (grants perms + seeds demo data)
 *   4. Deploy → New deployment → Web app → Execute as *Me*, Access *Anyone*
 ******************************************************************************/

const INV_CFG = {
  CACHE_TTL_SEC: 60,
  LOCK_TIMEOUT_MS: 15000,
  SHEETS: {
    USERS:'InvUsers', WAREHOUSES:'Warehouses', CATEGORIES:'Categories', UOM:'Units',
    PRODUCTS:'Products', SUPPLIERS:'Suppliers', CUSTOMERS:'Customers',
    PURCHASE_ORDERS:'PurchaseOrders', SALES_ORDERS:'SalesOrders', INDENTS:'Indents',
    STOCK_MOVEMENTS:'StockMovements', APPROVALS:'Approvals', ACTIVITY_LOG:'InvActivityLog'
  }
};

const INV_SCHEMAS = {
  InvUsers:['id','email','password','name','role','warehouseId','active','createdDate'],
  Warehouses:['id','code','name','address','manager','type','active','createdDate'],
  Categories:['id','code','name','parent','description','createdDate'],
  Units:['id','name','symbol','type','createdDate'],
  Products:['id','sku','barcode','name','description','categoryId','unit','hsn','gstRate',
    'costPrice','sellingPrice','reorderLevel','minStock','maxStock','supplierId','status',
    'batchTracked','serialTracked','image','createdDate','updatedDate'],
  Suppliers:['id','code','name','contactPerson','email','phone','address','gstin','pan',
    'paymentTerms','status','createdDate','updatedDate'],
  Customers:['id','code','name','contactPerson','email','phone','address','gstin',
    'creditLimit','paymentTerms','status','createdDate','updatedDate'],
  PurchaseOrders:['id','poNumber','supplierId','warehouseId','poDate','expectedDate',
    'status','totalAmount','taxAmount','grandTotal','notes','lines','approvalStatus',
    'approvedBy','approvedDate','receivedDate','createdBy','createdDate','updatedDate'],
  SalesOrders:['id','soNumber','customerId','warehouseId','soDate','deliveryDate',
    'status','totalAmount','taxAmount','grandTotal','notes','lines','dispatchedDate',
    'createdBy','createdDate','updatedDate'],
  Indents:['id','indentNumber','requesterId','warehouseId','indentDate','requiredDate',
    'purpose','status','approvalStatus','approvedBy','approvedDate','fulfilledDate',
    'lines','notes','createdDate','updatedDate'],
  StockMovements:['id','moveNumber','type','productId','warehouseId','fromWarehouseId',
    'toWarehouseId','quantity','unit','unitPrice','totalValue','referenceType',
    'referenceId','batchNumber','expiryDate','notes','performedBy','movementDate','createdDate'],
  Approvals:['id','entityType','entityId','requestedBy','requestedDate','approver',
    'status','remarks','decidedDate','createdDate'],
  InvActivityLog:['ts','user','action','entity','entityId','payload']
};

const INV_JSON_FIELDS = {
  PurchaseOrders:['lines'], SalesOrders:['lines'], Indents:['lines']
};

// ─── WEB ENTRY ──────────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('inventory')
    .setTitle('Inventory · ERP').addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ─── INITIALISATION + DUMMY DATA SEEDING ────────────────────────────────────
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock(); lock.tryLock(INV_CFG.LOCK_TIMEOUT_MS);
  try {
    Object.entries(INV_SCHEMAS).forEach(([name, headers]) => {
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      inv_ensureHeaders_(sh, headers);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#111827').setFontColor('#FFFFFF');
    });
    const s1 = ss.getSheetByName('Sheet1');
    if (s1 && ss.getSheets().length > 1 && s1.getLastRow() <= 0) ss.deleteSheet(s1);

    seedInventoryDemoData_();
    inv_clearCache_();
    return { success:true, message:'Inventory sheets initialised & demo data seeded', sheets:Object.keys(INV_SCHEMAS) };
  } finally { try { lock.releaseLock(); } catch(e){} }
}

function seedInventoryDemoData_() {
  // Skip if already seeded (users exist)
  if (inv_readAll_('InvUsers').length > 0) return;

  const now = new Date();
  const iso = d => new Date(d).toISOString().split('T')[0];
  const daysAgo = n => { const d = new Date(now); d.setDate(d.getDate()-n); return iso(d); };
  const daysAhead = n => { const d = new Date(now); d.setDate(d.getDate()+n); return iso(d); };
  const uuid = () => Utilities.getUuid().slice(0,8).toUpperCase();
  const nowIso = new Date().toISOString();

  // Users
  [
    { email:'admin@inv.com',   password:'admin123', name:'Admin',           role:'Admin',       warehouseId:'' },
    { email:'manager@inv.com', password:'mgr123',   name:'Ops Manager',     role:'Manager',     warehouseId:'WH-001' },
    { email:'purchase@inv.com',password:'pur123',   name:'Purchase Officer',role:'Purchase',    warehouseId:'WH-001' },
    { email:'sales@inv.com',   password:'sal123',   name:'Sales Executive', role:'Sales',       warehouseId:'WH-001' },
    { email:'store@inv.com',   password:'str123',   name:'Store Keeper',    role:'Storekeeper', warehouseId:'WH-001' },
    { email:'emp@inv.com',     password:'emp123',   name:'John Employee',   role:'Employee',    warehouseId:'WH-001' }
  ].forEach(u => inv_insert_('InvUsers', { id:'INVU-'+uuid(), active:true, createdDate:nowIso, ...u }));

  // Warehouses
  [
    ['WH-001','MAIN','Main Warehouse','Plot 42, Andheri West, Mumbai','Ops Manager','Warehouse'],
    ['WH-002','RETAIL','Retail Store','Linking Road, Bandra, Mumbai','Store Manager','Store'],
    ['WH-003','ONLINE','Online Fulfilment Center','Bhiwandi Logistics Park','FC Head','FC']
  ].forEach(w => inv_insert_('Warehouses',{ id:w[0], code:w[1], name:w[2], address:w[3], manager:w[4], type:w[5], active:true, createdDate:daysAgo(120) }));

  // Categories
  [
    ['CAT-01','ELEC','Electronics','','Consumer electronics & accessories'],
    ['CAT-02','APPR','Apparel','','Clothing & accessories'],
    ['CAT-03','FOOD','Food & Beverages','','Consumables'],
    ['CAT-04','STAT','Stationery','','Office supplies'],
    ['CAT-05','FURN','Furniture','','Office & home furniture']
  ].forEach(c => inv_insert_('Categories',{ id:c[0], code:c[1], name:c[2], parent:c[3], description:c[4], createdDate:daysAgo(90) }));

  // Units
  [['UOM-01','Piece','pcs','count'],['UOM-02','Box','box','count'],['UOM-03','Kilogram','kg','weight'],['UOM-04','Litre','L','volume'],['UOM-05','Metre','m','length']]
    .forEach(u => inv_insert_('Units',{ id:u[0], name:u[1], symbol:u[2], type:u[3], createdDate:daysAgo(90) }));

  // Suppliers
  const suppliers = [
    ['SUP-001','SUP001','TechnoWorld Distributors','Rakesh Kumar','rakesh@technoworld.in','+919820011001','Andheri East, Mumbai','27AABCT1234H1Z5','AABCT1234H','Net 30'],
    ['SUP-002','SUP002','FashionHub Wholesale',    'Priya Shah',  'priya@fashionhub.in',  '+919820011002','Kalbadevi, Mumbai','27AAECF5678K1Z6','AAECF5678K','Net 15'],
    ['SUP-003','SUP003','FoodGrain Suppliers',     'Amit Patel',  'amit@foodgrain.in',    '+919820011003','Vashi APMC, Navi Mumbai','27AABCF9012L1Z8','AABCF9012L','Net 7'],
    ['SUP-004','SUP004','OfficeMax Enterprises',   'Sanjay Roy',  'sanjay@officemax.in',  '+919820011004','Powai, Mumbai','27AAACO3456M1Z9','AAACO3456M','Net 30'],
    ['SUP-005','SUP005','WoodCraft Manufacturers', 'Meera Iyer',  'meera@woodcraft.in',   '+919820011005','Bhiwandi Industrial Estate','27AAECW7890P1Z2','AAECW7890P','Net 45']
  ];
  suppliers.forEach(s => inv_insert_('Suppliers',{ id:s[0], code:s[1], name:s[2], contactPerson:s[3], email:s[4], phone:s[5], address:s[6], gstin:s[7], pan:s[8], paymentTerms:s[9], status:'Active', createdDate:daysAgo(60) }));

  // Customers
  const customers = [
    ['CUS-001','CUS001','ABC Retailers Pvt Ltd',    'Vinay Sharma', 'vinay@abcretail.com',    '+919812200001','Colaba, Mumbai','27AABCA1111Z1Z1', 500000,'Net 30'],
    ['CUS-002','CUS002','XYZ Enterprises',          'Neha Kapoor',  'neha@xyzent.com',        '+919812200002','Malad, Mumbai','27AABCX2222Z1Z2', 300000,'Net 15'],
    ['CUS-003','CUS003','Sunrise Traders',          'Karan Mehta',  'karan@sunrise.com',      '+919812200003','Thane West',    '27AABCS3333Z1Z3', 200000,'Net 30'],
    ['CUS-004','CUS004','FreshMart Chain',          'Divya Reddy',  'divya@freshmart.com',    '+919812200004','Powai, Mumbai', '27AABCF4444Z1Z4', 750000,'Net 45'],
    ['CUS-005','CUS005','QuickBuy Online',          'Rohit Verma',  'rohit@quickbuy.com',     '+919812200005','Vashi, Navi Mumbai','27AABCQ5555Z1Z5',1000000,'Net 30']
  ];
  customers.forEach(c => inv_insert_('Customers',{ id:c[0], code:c[1], name:c[2], contactPerson:c[3], email:c[4], phone:c[5], address:c[6], gstin:c[7], creditLimit:c[8], paymentTerms:c[9], status:'Active', createdDate:daysAgo(60) }));

  // Products
  const products = [
    ['PROD-001','SKU-001','8901234500011','Wireless Bluetooth Headphones','Noise-cancelling over-ear headphones','CAT-01','UOM-01','8518',18,1200,2499,50,20,300,'SUP-001'],
    ['PROD-002','SKU-002','8901234500028','USB-C Fast Charger 65W',       '65W GaN charger with USB-C',           'CAT-01','UOM-01','8504',18,650, 1299,80,30,400,'SUP-001'],
    ['PROD-003','SKU-003','8901234500035','Cotton T-Shirt (Round Neck)',   '100% cotton, unisex, S–XXL',           'CAT-02','UOM-01','6109',5, 180, 499, 100,50,600,'SUP-002'],
    ['PROD-004','SKU-004','8901234500042','Denim Jeans (Slim Fit)',        'Straight leg, mid-rise',                'CAT-02','UOM-01','6203',12,650, 1499,60,25,300,'SUP-002'],
    ['PROD-005','SKU-005','8901234500059','Basmati Rice 5kg Pack',         'Premium aged basmati',                  'CAT-03','UOM-02','1006',5, 450, 799, 40,15,200,'SUP-003'],
    ['PROD-006','SKU-006','8901234500066','Refined Sunflower Oil 1L',      'Refined edible oil',                     'CAT-03','UOM-04','1512',5, 120, 165, 100,40,500,'SUP-003'],
    ['PROD-007','SKU-007','8901234500073','A4 Printing Paper (500 sheets)','80 GSM, 500 sheets pack',                'CAT-04','UOM-02','4802',12,220, 349, 70, 30,400,'SUP-004'],
    ['PROD-008','SKU-008','8901234500080','Gel Pen Pack (10 pcs)',         'Blue ink, 0.7mm',                        'CAT-04','UOM-02','9608',18,80,  149, 150,50,700,'SUP-004'],
    ['PROD-009','SKU-009','8901234500097','Office Executive Chair',        'Ergonomic mesh back',                    'CAT-05','UOM-01','9401',18,3200,6499,15, 5, 60, 'SUP-005'],
    ['PROD-010','SKU-010','8901234500103','Wooden Study Table',            '4ft × 2ft solid wood',                    'CAT-05','UOM-01','9403',18,4500,8999,10, 3, 40, 'SUP-005']
  ];
  products.forEach(p => inv_insert_('Products',{
    id:p[0], sku:p[1], barcode:p[2], name:p[3], description:p[4], categoryId:p[5], unit:p[6],
    hsn:p[7], gstRate:p[8], costPrice:p[9], sellingPrice:p[10], reorderLevel:p[11],
    minStock:p[12], maxStock:p[13], supplierId:p[14],
    status:'Active', batchTracked:false, serialTracked:false, image:'',
    createdDate:daysAgo(60)
  }));

  // Purchase Orders (3 records — Draft, Approved, Received)
  const poData = [
    { supplier:'SUP-001', wh:'WH-001', date:daysAgo(20), exp:daysAgo(10), status:'Received', appr:'Approved',
      lines:[
        { productId:'PROD-001', qty:50, price:1200 },
        { productId:'PROD-002', qty:80, price:650 }
      ]},
    { supplier:'SUP-003', wh:'WH-001', date:daysAgo(5), exp:daysAhead(5), status:'Approved', appr:'Approved',
      lines:[
        { productId:'PROD-005', qty:40, price:450 },
        { productId:'PROD-006', qty:100, price:120 }
      ]},
    { supplier:'SUP-002', wh:'WH-002', date:daysAgo(2), exp:daysAhead(10), status:'Draft', appr:'Pending',
      lines:[
        { productId:'PROD-003', qty:100, price:180 },
        { productId:'PROD-004', qty:60, price:650 }
      ]}
  ];
  poData.forEach((p, i) => {
    const sub = p.lines.reduce((s,l) => s + l.qty*l.price, 0);
    const tax = Math.round(sub * 0.18);
    inv_insert_('PurchaseOrders',{
      id:'PO-'+uuid(), poNumber:'PO/'+now.getFullYear()+'/'+String(i+1).padStart(4,'0'),
      supplierId:p.supplier, warehouseId:p.wh, poDate:p.date, expectedDate:p.exp,
      status:p.status, totalAmount:sub, taxAmount:tax, grandTotal:sub+tax,
      notes:'System-seeded order', lines:JSON.stringify(p.lines),
      approvalStatus:p.appr,
      approvedBy: p.appr === 'Approved' ? 'Ops Manager' : '',
      approvedDate: p.appr === 'Approved' ? daysAgo(i+3) : '',
      receivedDate: p.status === 'Received' ? daysAgo(i+1) : '',
      createdBy:'Purchase Officer', createdDate:p.date
    });
  });

  // Sales Orders (3 records)
  const soData = [
    { customer:'CUS-001', wh:'WH-001', date:daysAgo(15), del:daysAgo(12), status:'Dispatched', lines:[{ productId:'PROD-001', qty:10, price:2499 },{ productId:'PROD-002', qty:15, price:1299 }] },
    { customer:'CUS-004', wh:'WH-001', date:daysAgo(4),  del:daysAhead(2), status:'Confirmed', lines:[{ productId:'PROD-005', qty:20, price:799  },{ productId:'PROD-006', qty:30, price:165 }] },
    { customer:'CUS-002', wh:'WH-002', date:daysAgo(1),  del:daysAhead(3), status:'Draft',      lines:[{ productId:'PROD-003', qty:50, price:499  },{ productId:'PROD-004', qty:25, price:1499 }] }
  ];
  soData.forEach((s, i) => {
    const sub = s.lines.reduce((a,l)=>a+l.qty*l.price,0);
    const tax = Math.round(sub * 0.18);
    inv_insert_('SalesOrders',{
      id:'SO-'+uuid(), soNumber:'SO/'+now.getFullYear()+'/'+String(i+1).padStart(4,'0'),
      customerId:s.customer, warehouseId:s.wh, soDate:s.date, deliveryDate:s.del,
      status:s.status, totalAmount:sub, taxAmount:tax, grandTotal:sub+tax,
      notes:'', lines:JSON.stringify(s.lines),
      dispatchedDate: s.status === 'Dispatched' ? daysAgo(i+2) : '',
      createdBy:'Sales Executive', createdDate:s.date
    });
  });

  // Indents (3 records)
  const indents = [
    { req:'INVU-EMP', wh:'WH-001', date:daysAgo(3), need:daysAhead(2), purpose:'Office supplies replenishment', status:'Approved',    appr:'Approved', lines:[{ productId:'PROD-007', qty:5 },{ productId:'PROD-008', qty:10 }] },
    { req:'INVU-EMP', wh:'WH-001', date:daysAgo(1), need:daysAhead(3), purpose:'New joiner desk setup',         status:'Pending',     appr:'Pending',  lines:[{ productId:'PROD-009', qty:1 },{ productId:'PROD-010', qty:1 }] },
    { req:'INVU-EMP', wh:'WH-002', date:daysAgo(10),need:daysAgo(5),   purpose:'Store branding materials',       status:'Fulfilled',  appr:'Approved', lines:[{ productId:'PROD-007', qty:2 }] }
  ];
  indents.forEach((n,i) => inv_insert_('Indents',{
    id:'IND-'+uuid(), indentNumber:'IND/'+now.getFullYear()+'/'+String(i+1).padStart(4,'0'),
    requesterId:n.req, warehouseId:n.wh, indentDate:n.date, requiredDate:n.need,
    purpose:n.purpose, status:n.status, approvalStatus:n.appr,
    approvedBy: n.appr === 'Approved' ? 'Ops Manager' : '',
    approvedDate: n.appr === 'Approved' ? daysAgo(i+1) : '',
    fulfilledDate: n.status === 'Fulfilled' ? daysAgo(i+1) : '',
    lines:JSON.stringify(n.lines), notes:'', createdDate:n.date
  }));

  // Stock Movements (seed initial stock via received PO + a few sales)
  const seedMoves = [
    { type:'Stock In', productId:'PROD-001', wh:'WH-001', qty:50, price:1200, ref:'PO' },
    { type:'Stock In', productId:'PROD-002', wh:'WH-001', qty:80, price:650,  ref:'PO' },
    { type:'Stock In', productId:'PROD-005', wh:'WH-001', qty:40, price:450,  ref:'PO' },
    { type:'Stock In', productId:'PROD-006', wh:'WH-001', qty:100,price:120,  ref:'PO' },
    { type:'Stock In', productId:'PROD-003', wh:'WH-002', qty:100,price:180,  ref:'Opening' },
    { type:'Stock In', productId:'PROD-004', wh:'WH-002', qty:60, price:650,  ref:'Opening' },
    { type:'Stock In', productId:'PROD-007', wh:'WH-001', qty:70, price:220,  ref:'Opening' },
    { type:'Stock In', productId:'PROD-008', wh:'WH-001', qty:150,price:80,   ref:'Opening' },
    { type:'Stock In', productId:'PROD-009', wh:'WH-001', qty:15, price:3200, ref:'Opening' },
    { type:'Stock In', productId:'PROD-010', wh:'WH-001', qty:10, price:4500, ref:'Opening' },
    { type:'Stock Out',productId:'PROD-001', wh:'WH-001', qty:10, price:2499, ref:'SO' },
    { type:'Stock Out',productId:'PROD-002', wh:'WH-001', qty:15, price:1299, ref:'SO' }
  ];
  seedMoves.forEach((m,i) => inv_insert_('StockMovements',{
    id:'MV-'+uuid(), moveNumber:'MV/'+now.getFullYear()+'/'+String(i+1).padStart(5,'0'),
    type:m.type, productId:m.productId, warehouseId:m.wh, fromWarehouseId:'', toWarehouseId:'',
    quantity:m.qty, unit:'pcs', unitPrice:m.price, totalValue:m.qty*m.price,
    referenceType:m.ref, referenceId:'', batchNumber:'', expiryDate:'',
    notes:'Seeded movement', performedBy:'Store Keeper',
    movementDate:daysAgo(20-i), createdDate:daysAgo(20-i)
  }));

  // Approvals
  const draftPo = inv_readAll_('PurchaseOrders').find(p => p.status === 'Draft');
  const pendInd = inv_readAll_('Indents').find(i => i.approvalStatus === 'Pending');
  if (draftPo) inv_insert_('Approvals',{ id:'APR-'+uuid(), entityType:'PurchaseOrder', entityId:draftPo.id, requestedBy:'Purchase Officer', requestedDate:daysAgo(1), approver:'Ops Manager', status:'Pending', remarks:'', createdDate:daysAgo(1) });
  if (pendInd) inv_insert_('Approvals',{ id:'APR-'+uuid(), entityType:'Indent', entityId:pendInd.id, requestedBy:'John Employee', requestedDate:daysAgo(1), approver:'Ops Manager', status:'Pending', remarks:'', createdDate:daysAgo(1) });

  inv_clearCache_();
}

// ─── AUTH ───────────────────────────────────────────────────────────────────
function authenticateUser(email, password, warehouseId) {
  try {
    if (!email || !password) return { success:false, message:'Missing credentials' };
    const u = inv_readAll_('InvUsers').find(x =>
      String(x.email||'').toLowerCase() === String(email).toLowerCase() &&
      String(x.password) === String(password) &&
      (x.active === true || String(x.active).toLowerCase() === 'true' || x.active === '' || x.active == null));
    if (!u) return { success:false, message:'Invalid credentials' };
    const user = { id:u.id, email:u.email, name:u.name, role:u.role, warehouseId:u.warehouseId||null };
    PropertiesService.getUserProperties().setProperty('_invUser', JSON.stringify(user));
    inv_log_(user.email, 'login', 'InvUsers', u.id, { warehouseId });
    return { success:true, user };
  } catch (err) { return { success:false, message:err.message }; }
}
function getCurrentUser() {
  try { const r = PropertiesService.getUserProperties().getProperty('_invUser'); return r ? JSON.parse(r) : null; }
  catch(e){ return null; }
}
function signOut(){ PropertiesService.getUserProperties().deleteProperty('_invUser'); return { success:true }; }

// ─── GENERIC CRUD ENDPOINTS ────────────────────────────────────────────────
function getAllWarehouses(){ return inv_ok_(inv_readAll_('Warehouses')); }
function addWarehouse(d){ return inv_insertWithId_('Warehouses','WH',d); }
function updateWarehouse(id,u){ return inv_updateById_('Warehouses',id,u); }
function deleteWarehouse(id){ return inv_deleteById_('Warehouses',id); }

function getAllCategories(){ return inv_ok_(inv_readAll_('Categories')); }
function addCategory(d){ return inv_insertWithId_('Categories','CAT',d); }
function updateCategory(id,u){ return inv_updateById_('Categories',id,u); }
function deleteCategory(id){ return inv_deleteById_('Categories',id); }

function getAllUnits(){ return inv_ok_(inv_readAll_('Units')); }
function addUnit(d){ return inv_insertWithId_('Units','UOM',d); }

function getAllProducts(){ return inv_ok_(inv_readAll_('Products')); }
function addProduct(d){ return inv_insertWithId_('Products','PROD',d); }
function updateProduct(id,u){ return inv_updateById_('Products',id,u); }
function deleteProduct(id){ return inv_deleteById_('Products',id); }

function getAllSuppliers(){ return inv_ok_(inv_readAll_('Suppliers')); }
function addSupplier(d){ return inv_insertWithId_('Suppliers','SUP',d); }
function updateSupplier(id,u){ return inv_updateById_('Suppliers',id,u); }
function deleteSupplier(id){ return inv_deleteById_('Suppliers',id); }

function getAllCustomers(){ return inv_ok_(inv_readAll_('Customers')); }
function addCustomer(d){ return inv_insertWithId_('Customers','CUS',d); }
function updateCustomer(id,u){ return inv_updateById_('Customers',id,u); }
function deleteCustomer(id){ return inv_deleteById_('Customers',id); }

// ─── PURCHASE ORDERS ───────────────────────────────────────────────────────
function getAllPurchaseOrders(){ return inv_ok_(inv_readAll_('PurchaseOrders')); }
function addPurchaseOrder(data) {
  const u = getCurrentUser();
  data.poNumber = data.poNumber || generateSeqNumber_('PO', 'PurchaseOrders', 'poNumber');
  data.createdBy = data.createdBy || (u && u.name);
  data.status = data.status || 'Draft';
  data.approvalStatus = data.approvalStatus || 'Pending';
  data = inv_computeTotals_(data);
  const res = inv_insertWithId_('PurchaseOrders','PO',data);
  if (res.success) inv_insert_('Approvals',{ id:'APR-'+Utilities.getUuid().slice(0,8).toUpperCase(), entityType:'PurchaseOrder', entityId:res.id, requestedBy:(u&&u.name)||'', requestedDate:new Date().toISOString(), approver:'Manager', status:'Pending', remarks:'', createdDate:new Date().toISOString() });
  return res;
}
function updatePurchaseOrder(id, updates) {
  if (updates && updates.lines) updates = inv_computeTotals_(updates);
  return inv_updateById_('PurchaseOrders', id, updates);
}
function approvePurchaseOrder(id, remarks) {
  const u = getCurrentUser();
  if (!u || (u.role !== 'Admin' && u.role !== 'Manager')) return { success:false, message:'Not authorised' };
  inv_updateById_('PurchaseOrders', id, { approvalStatus:'Approved', approvedBy:u.name, approvedDate:new Date().toISOString(), status:'Approved' });
  inv_updateApprovalByEntity_('PurchaseOrder', id, { status:'Approved', remarks:remarks||'', decidedDate:new Date().toISOString() });
  return { success:true };
}
function rejectPurchaseOrder(id, remarks) {
  const u = getCurrentUser();
  if (!u || (u.role !== 'Admin' && u.role !== 'Manager')) return { success:false, message:'Not authorised' };
  inv_updateById_('PurchaseOrders', id, { approvalStatus:'Rejected', approvedBy:u.name, approvedDate:new Date().toISOString(), status:'Rejected' });
  inv_updateApprovalByEntity_('PurchaseOrder', id, { status:'Rejected', remarks:remarks||'', decidedDate:new Date().toISOString() });
  return { success:true };
}
function receivePurchaseOrder(id) {
  const u = getCurrentUser();
  const po = inv_readAll_('PurchaseOrders').find(p => p.id === id);
  if (!po) return { success:false, message:'PO not found' };
  if (po.approvalStatus !== 'Approved') return { success:false, message:'PO must be approved first' };
  const lines = Array.isArray(po.lines) ? po.lines : (po.lines ? JSON.parse(po.lines) : []);
  lines.forEach(l => {
    inv_insert_('StockMovements',{
      id:'MV-'+Utilities.getUuid().slice(0,8).toUpperCase(),
      moveNumber:generateSeqNumber_('MV','StockMovements','moveNumber'),
      type:'Stock In', productId:l.productId, warehouseId:po.warehouseId,
      fromWarehouseId:'', toWarehouseId:'',
      quantity:l.qty, unit:l.unit||'pcs', unitPrice:l.price, totalValue:l.qty*l.price,
      referenceType:'PO', referenceId:po.id, batchNumber:l.batchNumber||'',
      expiryDate:l.expiryDate||'', notes:'Received via PO', performedBy:(u&&u.name)||'',
      movementDate:new Date().toISOString(), createdDate:new Date().toISOString()
    });
  });
  inv_updateById_('PurchaseOrders', id, { status:'Received', receivedDate:new Date().toISOString() });
  return { success:true };
}
function deletePurchaseOrder(id){ return inv_deleteById_('PurchaseOrders',id); }

// ─── SALES ORDERS ──────────────────────────────────────────────────────────
function getAllSalesOrders(){ return inv_ok_(inv_readAll_('SalesOrders')); }
function addSalesOrder(data) {
  const u = getCurrentUser();
  data.soNumber = data.soNumber || generateSeqNumber_('SO','SalesOrders','soNumber');
  data.createdBy = data.createdBy || (u && u.name);
  data.status = data.status || 'Draft';
  data = inv_computeTotals_(data);
  return inv_insertWithId_('SalesOrders','SO',data);
}
function updateSalesOrder(id, updates) {
  if (updates && updates.lines) updates = inv_computeTotals_(updates);
  return inv_updateById_('SalesOrders', id, updates);
}
function dispatchSalesOrder(id) {
  const u = getCurrentUser();
  const so = inv_readAll_('SalesOrders').find(s => s.id === id);
  if (!so) return { success:false, message:'SO not found' };
  const lines = Array.isArray(so.lines) ? so.lines : (so.lines ? JSON.parse(so.lines) : []);
  // Check stock
  for (const l of lines) {
    const bal = computeStockBalance(l.productId, so.warehouseId);
    if (bal < Number(l.qty || 0)) return { success:false, message:'Insufficient stock for '+l.productId+' (have '+bal+', need '+l.qty+')' };
  }
  lines.forEach(l => inv_insert_('StockMovements',{
    id:'MV-'+Utilities.getUuid().slice(0,8).toUpperCase(),
    moveNumber:generateSeqNumber_('MV','StockMovements','moveNumber'),
    type:'Stock Out', productId:l.productId, warehouseId:so.warehouseId,
    fromWarehouseId:'', toWarehouseId:'',
    quantity:l.qty, unit:l.unit||'pcs', unitPrice:l.price, totalValue:l.qty*l.price,
    referenceType:'SO', referenceId:so.id, batchNumber:'', expiryDate:'',
    notes:'Dispatched via SO', performedBy:(u&&u.name)||'',
    movementDate:new Date().toISOString(), createdDate:new Date().toISOString()
  }));
  inv_updateById_('SalesOrders', id, { status:'Dispatched', dispatchedDate:new Date().toISOString() });
  return { success:true };
}
function deleteSalesOrder(id){ return inv_deleteById_('SalesOrders',id); }

// ─── INDENTS ───────────────────────────────────────────────────────────────
function getAllIndents(){ return inv_ok_(inv_readAll_('Indents')); }
function addIndent(data) {
  const u = getCurrentUser();
  data.indentNumber = data.indentNumber || generateSeqNumber_('IND','Indents','indentNumber');
  data.requesterId = data.requesterId || (u && u.id);
  data.status = data.status || 'Pending';
  data.approvalStatus = data.approvalStatus || 'Pending';
  const res = inv_insertWithId_('Indents','IND',data);
  if (res.success) inv_insert_('Approvals',{ id:'APR-'+Utilities.getUuid().slice(0,8).toUpperCase(), entityType:'Indent', entityId:res.id, requestedBy:(u&&u.name)||'', requestedDate:new Date().toISOString(), approver:'Manager', status:'Pending', remarks:'', createdDate:new Date().toISOString() });
  return res;
}
function approveIndent(id, remarks) {
  const u = getCurrentUser();
  if (!u || (u.role !== 'Admin' && u.role !== 'Manager')) return { success:false, message:'Not authorised' };
  inv_updateById_('Indents', id, { approvalStatus:'Approved', approvedBy:u.name, approvedDate:new Date().toISOString(), status:'Approved' });
  inv_updateApprovalByEntity_('Indent', id, { status:'Approved', remarks:remarks||'', decidedDate:new Date().toISOString() });
  return { success:true };
}
function rejectIndent(id, remarks) {
  const u = getCurrentUser();
  if (!u || (u.role !== 'Admin' && u.role !== 'Manager')) return { success:false, message:'Not authorised' };
  inv_updateById_('Indents', id, { approvalStatus:'Rejected', approvedBy:u.name, approvedDate:new Date().toISOString(), status:'Rejected' });
  inv_updateApprovalByEntity_('Indent', id, { status:'Rejected', remarks:remarks||'', decidedDate:new Date().toISOString() });
  return { success:true };
}
function fulfillIndent(id) {
  const u = getCurrentUser();
  const ind = inv_readAll_('Indents').find(i => i.id === id);
  if (!ind) return { success:false, message:'Indent not found' };
  if (ind.approvalStatus !== 'Approved') return { success:false, message:'Indent must be approved first' };
  const lines = Array.isArray(ind.lines) ? ind.lines : (ind.lines ? JSON.parse(ind.lines) : []);
  lines.forEach(l => inv_insert_('StockMovements',{
    id:'MV-'+Utilities.getUuid().slice(0,8).toUpperCase(),
    moveNumber:generateSeqNumber_('MV','StockMovements','moveNumber'),
    type:'Stock Out', productId:l.productId, warehouseId:ind.warehouseId,
    fromWarehouseId:'', toWarehouseId:'', quantity:l.qty, unit:l.unit||'pcs',
    unitPrice:0, totalValue:0, referenceType:'Indent', referenceId:ind.id,
    batchNumber:'', expiryDate:'', notes:'Issued via indent', performedBy:(u&&u.name)||'',
    movementDate:new Date().toISOString(), createdDate:new Date().toISOString()
  }));
  inv_updateById_('Indents', id, { status:'Fulfilled', fulfilledDate:new Date().toISOString() });
  return { success:true };
}
function deleteIndent(id){ return inv_deleteById_('Indents',id); }

// ─── STOCK MOVEMENTS ───────────────────────────────────────────────────────
function getAllStockMovements(){ return inv_ok_(inv_readAll_('StockMovements')); }
function addStockMovement(data) {
  data.moveNumber = data.moveNumber || generateSeqNumber_('MV','StockMovements','moveNumber');
  const u = getCurrentUser();
  data.performedBy = data.performedBy || (u && u.name);
  data.totalValue = Number(data.quantity||0) * Number(data.unitPrice||0);
  data.movementDate = data.movementDate || new Date().toISOString();
  // Handle transfer as two movements
  if (data.type === 'Transfer' && data.fromWarehouseId && data.toWarehouseId) {
    inv_insert_('StockMovements',{
      id:'MV-'+Utilities.getUuid().slice(0,8).toUpperCase(),
      moveNumber:generateSeqNumber_('MV','StockMovements','moveNumber'),
      type:'Stock Out', productId:data.productId, warehouseId:data.fromWarehouseId,
      fromWarehouseId:data.fromWarehouseId, toWarehouseId:data.toWarehouseId,
      quantity:data.quantity, unit:data.unit||'pcs', unitPrice:data.unitPrice||0,
      totalValue:data.totalValue, referenceType:'Transfer', referenceId:'',
      batchNumber:data.batchNumber||'', expiryDate:data.expiryDate||'',
      notes:'Transfer out: '+data.notes, performedBy:data.performedBy,
      movementDate:data.movementDate, createdDate:new Date().toISOString()
    });
    data.type = 'Stock In';
    data.warehouseId = data.toWarehouseId;
    data.notes = 'Transfer in: ' + data.notes;
  }
  return inv_insertWithId_('StockMovements','MV',data);
}

// ─── STOCK LEVELS / DERIVED  ───────────────────────────────────────────────
function computeStockBalance(productId, warehouseId) {
  const moves = inv_readAll_('StockMovements');
  let bal = 0;
  moves.forEach(m => {
    if (m.productId !== productId) return;
    if (warehouseId && m.warehouseId !== warehouseId) return;
    const q = Number(m.quantity || 0);
    if (m.type === 'Stock In') bal += q;
    else if (m.type === 'Stock Out') bal -= q;
    else if (m.type === 'Adjustment') bal += q; // adjustment quantity can be negative
  });
  return bal;
}
function getStockLevels() {
  const products = inv_readAll_('Products');
  const warehouses = inv_readAll_('Warehouses');
  const moves = inv_readAll_('StockMovements');
  // Aggregate: productId × warehouseId → balance
  const map = {};
  products.forEach(p => warehouses.forEach(w => { map[p.id+'::'+w.id] = { productId:p.id, warehouseId:w.id, product:p, warehouse:w, qty:0, value:0 }; }));
  moves.forEach(m => {
    const k = m.productId+'::'+m.warehouseId;
    if (!map[k]) return;
    const q = Number(m.quantity||0);
    if (m.type === 'Stock In') { map[k].qty += q; map[k].value += q * Number(m.unitPrice||0); }
    else if (m.type === 'Stock Out') { map[k].qty -= q; }
    else if (m.type === 'Adjustment') { map[k].qty += q; }
  });
  const rows = Object.values(map).filter(r => r.qty !== 0);
  return { success:true, data: rows };
}
function getLowStockAlerts() {
  const products = inv_readAll_('Products');
  const stock = getStockLevels().data;
  // Aggregate qty across all warehouses per product
  const totalPerProduct = {};
  stock.forEach(s => { totalPerProduct[s.productId] = (totalPerProduct[s.productId] || 0) + s.qty; });
  return { success:true, data: products.filter(p => (totalPerProduct[p.id]||0) <= Number(p.reorderLevel||0))
    .map(p => ({ product:p, currentStock: totalPerProduct[p.id]||0 })) };
}
function getDashboardStats() {
  const products = inv_readAll_('Products');
  const pos = inv_readAll_('PurchaseOrders');
  const sos = inv_readAll_('SalesOrders');
  const indents = inv_readAll_('Indents');
  const moves = inv_readAll_('StockMovements');
  const suppliers = inv_readAll_('Suppliers');
  const customers = inv_readAll_('Customers');
  const stockValue = moves.reduce((s,m) => s + (m.type==='Stock In' ? Number(m.totalValue||0) : m.type==='Stock Out' ? -Number(m.totalValue||0) : 0), 0);
  return { success:true, data:{
    products:products.length, activeProducts:products.filter(p=>p.status==='Active').length,
    suppliers:suppliers.length, customers:customers.length,
    openPOs:pos.filter(p=>p.status!=='Received' && p.status!=='Rejected').length,
    pendingPOApprovals:pos.filter(p=>p.approvalStatus==='Pending').length,
    openSOs:sos.filter(s=>s.status!=='Dispatched').length,
    pendingIndents:indents.filter(i=>i.approvalStatus==='Pending').length,
    stockValue: Math.round(stockValue),
    lowStockCount: getLowStockAlerts().data.length,
    recentMovements: moves.slice(-10).reverse()
  }};
}

// ─── APPROVALS ─────────────────────────────────────────────────────────────
function getAllApprovals(){ return inv_ok_(inv_readAll_('Approvals')); }
function getPendingApprovalsForMe() {
  const u = getCurrentUser();
  const list = inv_readAll_('Approvals').filter(a => a.status === 'Pending');
  return { success:true, data:list, user:u };
}

// ─── HELPERS ───────────────────────────────────────────────────────────────
function inv_computeTotals_(data) {
  let lines = data.lines;
  if (typeof lines === 'string') { try { lines = JSON.parse(lines); } catch(e){ lines = []; } }
  if (!Array.isArray(lines)) return data;
  const sub = lines.reduce((s,l)=>s + Number(l.qty||0)*Number(l.price||0), 0);
  const gst = data.gstRate ? Number(data.gstRate) : 18;
  const tax = Math.round(sub * gst / 100);
  data.totalAmount = sub; data.taxAmount = tax; data.grandTotal = sub + tax;
  return data;
}
function generateSeqNumber_(prefix, sheetName, field) {
  const rows = inv_readAll_(sheetName);
  const yr = new Date().getFullYear();
  const seq = rows.filter(r => String(r[field]||'').indexOf(prefix+'/'+yr) === 0).length + 1;
  return prefix + '/' + yr + '/' + String(seq).padStart(4, '0');
}
function inv_updateApprovalByEntity_(entityType, entityId, updates) {
  const rec = inv_readAll_('Approvals').find(a => a.entityType === entityType && a.entityId === entityId && a.status === 'Pending');
  if (rec) inv_updateById_('Approvals', rec.id, updates);
}

function inv_getSheet_(name){ const ss = SpreadsheetApp.getActiveSpreadsheet(); let sh = ss.getSheetByName(name); if(!sh){ sh = ss.insertSheet(name); inv_ensureHeaders_(sh, INV_SCHEMAS[name]||['id']); } return sh; }
function inv_ensureHeaders_(sh, wanted){ const lc = Math.max(sh.getLastColumn(),1); const existing = sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v||'')); const present = existing.filter(Boolean); const missing = wanted.filter(h => present.indexOf(h) === -1); if(present.length===0 && missing.length){ sh.getRange(1,1,1,missing.length).setValues([missing]); return; } if(missing.length) sh.getRange(1,present.length+1,1,missing.length).setValues([missing]); }
function inv_getHeaders_(name){ const sh = inv_getSheet_(name); const w = INV_SCHEMAS[name]||[]; if(w.length) inv_ensureHeaders_(sh,w); const lc = Math.max(sh.getLastColumn(),1); return sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v||'')).filter(Boolean); }
function inv_readAll_(name){
  const cache = CacheService.getScriptCache(); const key = 'INV::'+name;
  const hit = cache.get(key); if(hit){ try { return JSON.parse(hit); } catch(e){} }
  const sh = inv_getSheet_(name); const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1){ cache.put(key,'[]',INV_CFG.CACHE_TTL_SEC); return []; }
  const vals = sh.getRange(1,1,lr,lc).getValues();
  const hdrs = vals[0].map(v=>String(v||'')); const jf = INV_JSON_FIELDS[name]||[]; const rows = [];
  for (let i=1;i<vals.length;i++){ const r = vals[i]; if (r.every(v=>v===''||v===null)) continue;
    const o = {}; for (let j=0;j<hdrs.length;j++){ if(!hdrs[j]) continue; let v = r[j]; if (v instanceof Date) v = v.toISOString(); if (jf.indexOf(hdrs[j])!==-1 && typeof v === 'string' && v){ try { v = JSON.parse(v); } catch(e){} } o[hdrs[j]] = v; } rows.push(o); }
  try { cache.put(key, JSON.stringify(rows), INV_CFG.CACHE_TTL_SEC); } catch(e){}
  return rows;
}
function inv_serialize_(name, obj){ const hdrs = inv_getHeaders_(name); const extra = Object.keys(obj).filter(k => hdrs.indexOf(k) === -1); if (extra.length){ const sh = inv_getSheet_(name); sh.getRange(1, hdrs.length+1, 1, extra.length).setValues([extra]); extra.forEach(h => hdrs.push(h)); } const jf = INV_JSON_FIELDS[name]||[]; return hdrs.map(h => { let v = obj[h]; if (v === undefined || v === null) return ''; if (jf.indexOf(h) !== -1 && typeof v === 'object') return JSON.stringify(v); if (v instanceof Date) return v.toISOString(); return v; }); }
function inv_insert_(name, obj){ const sh = inv_getSheet_(name); const row = inv_serialize_(name, obj); sh.appendRow(row); inv_clearKey_(name); return obj; }
function inv_insertWithId_(name, prefix, data){ try { if(!data) return { success:false, message:'No data' }; const lock = LockService.getScriptLock(); lock.tryLock(INV_CFG.LOCK_TIMEOUT_MS); try { const now = new Date().toISOString(); if (!data.id) data.id = prefix+'-'+Utilities.getUuid().slice(0,8).toUpperCase(); if (!data.createdDate) data.createdDate = now; inv_insert_(name, data); inv_log_((getCurrentUser()||{}).email||'system','create',name,data.id,null); return { success:true, id:data.id, data }; } finally { try { lock.releaseLock(); } catch(e){} } } catch (e){ return { success:false, message:e.message }; } }
function inv_findRow_(name, field, val){ const sh = inv_getSheet_(name); const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return { sheet:sh, rowIndex:-1, headers:[], row:null }; const vals = sh.getRange(1,1,lr,lc).getValues(); const hdrs = vals[0].map(v=>String(v||'')); const idx = hdrs.indexOf(field); if (idx === -1) return { sheet:sh, rowIndex:-1, headers:hdrs, row:null }; for (let i=1;i<vals.length;i++){ if (String(vals[i][idx]) === String(val)) return { sheet:sh, rowIndex:i+1, headers:hdrs, row:vals[i] }; } return { sheet:sh, rowIndex:-1, headers:hdrs, row:null }; }
function inv_updateById_(name, id, updates){ try { if (!id) return { success:false, message:'Missing id' }; if (!updates || typeof updates !== 'object') return { success:false, message:'Missing updates' }; const lock = LockService.getScriptLock(); lock.tryLock(INV_CFG.LOCK_TIMEOUT_MS); try { const loc = inv_findRow_(name,'id',id); if (loc.rowIndex === -1) return { success:false, message:'Not found' }; const cur = {}; loc.headers.forEach((h,j)=>{ if(h) cur[h] = loc.row[j]; }); Object.assign(cur, updates); if (!cur.updatedDate) cur.updatedDate = new Date().toISOString(); const nr = inv_serialize_(name, cur); loc.sheet.getRange(loc.rowIndex, 1, 1, nr.length).setValues([nr]); inv_clearKey_(name); inv_log_((getCurrentUser()||{}).email||'system','update',name,id,Object.keys(updates)); return { success:true, id }; } finally { try { lock.releaseLock(); } catch(e){} } } catch(e){ return { success:false, message:e.message }; } }
function inv_deleteById_(name, id){ try { if (!id) return { success:false, message:'Missing id' }; const lock = LockService.getScriptLock(); lock.tryLock(INV_CFG.LOCK_TIMEOUT_MS); try { const loc = inv_findRow_(name,'id',id); if (loc.rowIndex === -1) return { success:false, message:'Not found' }; loc.sheet.deleteRow(loc.rowIndex); inv_clearKey_(name); inv_log_((getCurrentUser()||{}).email||'system','delete',name,id,null); return { success:true }; } finally { try { lock.releaseLock(); } catch(e){} } } catch(e){ return { success:false, message:e.message }; } }
function inv_ok_(rows){ return { success:true, data:rows }; }
function inv_clearKey_(name){ try { CacheService.getScriptCache().remove('INV::'+name); } catch(e){} }
function inv_clearCache_(){ try { CacheService.getScriptCache().removeAll(Object.keys(INV_SCHEMAS).map(s=>'INV::'+s)); } catch(e){} }
function inv_log_(user, action, entity, entityId, payload){ try { const sh = inv_getSheet_('InvActivityLog'); sh.appendRow([new Date().toISOString(), user||'anon', action||'', entity||'', entityId||'', payload ? JSON.stringify(payload).slice(0,4000) : '']); } catch(e){} }

// ─── MAINTENANCE ───────────────────────────────────────────────────────────
function resetCache(){ inv_clearCache_(); return { success:true }; }
function healthCheck(){ const s = {}; Object.keys(INV_SCHEMAS).forEach(n => { const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); s[n] = sh ? { rows:Math.max(0, sh.getLastRow()-1), cols:sh.getLastColumn() } : 'missing'; }); return { success:true, sheets:s, timestamp:new Date().toISOString() }; }

// ─── BULK CSV IMPORT ───────────────────────────────────────────────────────
// Accepts a sheet name (e.g. 'Products') and an array of records (objects).
// Returns { success, ok, failed, errors:[{row, message}] }. Whole call runs
// under a single LockService lease to avoid partial writes on concurrency.
function bulkImport(sheetName, rows) {
  try {
    if (!sheetName) return { success:false, message:'Missing sheetName' };
    if (!Array.isArray(rows) || rows.length === 0) return { success:false, message:'No rows to import' };
    if (!INV_SCHEMAS[sheetName]) return { success:false, message:'Unknown sheet: '+sheetName };
    const prefix = { Products:'PROD', Suppliers:'SUP', Customers:'CUS', Categories:'CAT', Warehouses:'WH', Units:'UOM' }[sheetName] || 'BLK';
    const lock = LockService.getScriptLock(); lock.tryLock(30000);
    const result = { success:true, ok:0, failed:0, errors:[] };
    const nowIso = new Date().toISOString();
    try {
      rows.forEach((row, i) => {
        try {
          if (!row || typeof row !== 'object') throw new Error('Not an object');
          const rec = Object.assign({}, row);
          if (!rec.id) rec.id = prefix + '-' + Utilities.getUuid().slice(0,8).toUpperCase();
          if (!rec.createdDate) rec.createdDate = nowIso;
          inv_insert_(sheetName, rec);
          result.ok++;
        } catch (e) {
          result.failed++;
          result.errors.push({ row: i + 1, message: String(e.message || e) });
        }
      });
      inv_clearKey_(sheetName);
      inv_log_((getCurrentUser()||{}).email||'system', 'bulkImport', sheetName, '', { ok:result.ok, failed:result.failed });
    } finally { try { lock.releaseLock(); } catch(e){} }
    return result;
  } catch (e) { return { success:false, message: e.message }; }
}
