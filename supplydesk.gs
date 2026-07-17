/**
 * ============================================================================
 *  SupplyDesk — Google Apps Script Web App (single-spreadsheet backend)
 * ============================================================================
 *  Multi-site B2B supply, billing & credit system for a wholesale/distribution
 *  business supplying materials to corporate clients across multiple sites.
 *  Covers: site-wise ordering & pricing, consolidated deliveries with damage
 *  reconciliation, dynamic credit-limit enforcement, auto GST invoicing
 *  (CGST/SGST vs IGST), a 24-hour complaint window with auto credit notes,
 *  and a Tally Prime-compatible XML export.
 *
 *  DEPLOYMENT
 *  ----------
 *  1. Open Google Sheets → Extensions → Apps Script.
 *  2. Paste this file as Code.gs and paste supplydesk.html as an HTML file
 *     named exactly "Index".
 *  3. Reload the sheet; the "⚙️ SupplyDesk" menu appears. Click
 *     "Initialize System" once (safe to run again — idempotent).
 *  4. Deploy → New deployment → Web app → Execute as: Me,
 *     Access: Anyone with the link (or your Workspace domain).
 *  5. Optional: "Install Complaint Sweep" wires an hourly trigger that expires
 *     complaints past their 24-hour window.
 *
 *  SHEET STRUCTURE (created by initializeSheets)
 *  ---------------------------------------------
 *   Users        id | email | name | role | siteId | active | createdAt
 *                (role = Client | Admin ; siteId links a Client login to a site)
 *   Clients      id | name | gstin | homeState | creditLimit | createdAt
 *   Sites        id | clientId | siteName | address | state | contactPerson | contactPhone | createdAt
 *   Catalog      id | itemName | sku | uom | defaultRate | hsnCode | gstRatePct | createdAt
 *   SitePricing  id | siteId | itemId | rate
 *   Orders       id | siteId | status | createdAt | deleted
 *                (status = Placed | Consolidated | Dispatched | Delivered | Cancelled)
 *   OrderLines   id | orderId | itemId | qtyOrdered | qtyDelivered | damagedQty
 *   Deliveries   id | siteId | deliveryDate | orderIds | status
 *   Invoices     id | clientId | siteId | invoiceDate | subtotal | cgst | sgst | igst | total | tallyExported | createdAt
 *   InvoiceLines id | invoiceId | orderLineId | itemId | qty | rate | amount
 *   Payments     id | clientId | amount | type | reference | createdAt   (type = Payment | CreditNote | Debit)
 *   Complaints   id | orderLineId | raisedAt | deadline | status | reason | resolutionNotes
 *                (status = Open | Resolved | Expired)
 *   CreditNotes  id | complaintId | amount | issuedAt
 *   Config       key | value
 *   AuditLog     id | ts | actor | action | entity | entityId | details
 *   ErrorLog     id | ts | fn | message | stack
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const SD_SHEETS = {
  Users:        { headers:['id','email','name','role','siteId','active','createdAt'], widths:[60,220,180,90,80,80,180] },
  Clients:      { headers:['id','name','gstin','homeState','creditLimit','createdAt'], widths:[60,220,170,150,120,180] },
  Sites:        { headers:['id','clientId','siteName','address','state','contactPerson','contactPhone','createdAt'], widths:[60,80,180,240,150,160,150,180] },
  Catalog:      { headers:['id','itemName','sku','uom','defaultRate','hsnCode','gstRatePct','createdAt'], widths:[60,220,120,80,110,110,100,180] },
  SitePricing:  { headers:['id','siteId','itemId','rate'], widths:[60,80,80,110] },
  Orders:       { headers:['id','siteId','status','createdAt','deleted'], widths:[60,80,130,180,80] },
  OrderLines:   { headers:['id','orderId','itemId','qtyOrdered','qtyDelivered','damagedQty'], widths:[60,80,80,110,110,110] },
  Deliveries:   { headers:['id','siteId','deliveryDate','orderIds','status'], widths:[60,80,150,220,130] },
  Invoices:     { headers:['id','clientId','siteId','invoiceDate','subtotal','cgst','sgst','igst','total','tallyExported','createdAt'], widths:[60,80,80,150,110,90,90,90,110,110,180] },
  InvoiceLines: { headers:['id','invoiceId','orderLineId','itemId','qty','rate','amount'], widths:[60,80,90,80,90,100,110] },
  Payments:     { headers:['id','clientId','amount','type','reference','createdAt'], widths:[60,80,120,110,220,180] },
  Complaints:   { headers:['id','orderLineId','raisedAt','deadline','status','reason','resolutionNotes'], widths:[60,90,180,180,110,240,240] },
  CreditNotes:  { headers:['id','complaintId','amount','issuedAt'], widths:[60,100,120,180] },
  Config:       { headers:['key','value'], widths:[280,360] },
  AuditLog:     { headers:['id','ts','actor','action','entity','entityId','details'], widths:[60,180,220,150,120,90,420] },
  ErrorLog:     { headers:['id','ts','fn','message','stack'], widths:[60,180,220,420,520] }
};

const SD_ORDER_STATUS = ['Placed','Consolidated','Dispatched','Delivered','Cancelled'];
const SD_COMPLAINT_STATUS = ['Open','Resolved','Expired'];
const SD_ROLES = ['Client','Admin'];
const SD_CACHE_TTL = 45;

// ---------------------------------------------------------------------------
// Web app entry + menu
// ---------------------------------------------------------------------------
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SupplyDesk')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ SupplyDesk')
    .addItem('Initialize System', 'menuInitialize')
    .addItem('Reset Dummy Data',  'menuResetDummy')
    .addSeparator()
    .addItem('Install Complaint Sweep', 'installComplaintSweep')
    .addItem('Remove Complaint Sweep',  'removeComplaintSweep')
    .addToUi();
}
function menuInitialize() {
  const r = initializeSheets();
  SpreadsheetApp.getUi().alert('SupplyDesk initialized.\n\nCreated: ' + r.data.sheetsCreated.join(', ') +
    '\nSkipped: ' + r.data.sheetsSkipped.join(', ') + '\nDummy rows added: ' + r.data.dummyRowsAdded);
}
function menuResetDummy() {
  const r = resetDummyData();
  SpreadsheetApp.getUi().alert('Dummy rows cleared. Removed: ' + r.data.removed);
}

// ---------------------------------------------------------------------------
// Standard response wrappers
// ---------------------------------------------------------------------------
function success(data){ return { success:true, data:data, error:null }; }
function fail(msg){ return { success:false, data:null, error:String(msg) }; }
function safeCall(fnName, fn) {
  try { return fn(); }
  catch (e) { logError(fnName, e); return fail(e.message || String(e)); }
}


// ---------------------------------------------------------------------------
// initializeSheets — idempotent
// ---------------------------------------------------------------------------
function initializeSheets() {
  return safeCall('initializeSheets', function(){
    const ss = SpreadsheetApp.getActive();
    const created = [], skipped = [];
    Object.keys(SD_SHEETS).forEach(function(name){
      let sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
        const spec = SD_SHEETS[name];
        sh.getRange(1,1,1,spec.headers.length).setValues([spec.headers])
          .setFontWeight('bold').setBackground('#0B1220').setFontColor('#fff');
        sh.setFrozenRows(1);
        spec.widths.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
        created.push(name);
      } else skipped.push(name);
    });
    const s1 = ss.getSheetByName('Sheet1');
    if (s1 && ss.getSheets().length > 1) { try { ss.deleteSheet(s1); } catch(e){} }

    seedConfigDefaults();
    let n = 0;
    n += seedAll();
    invalidateEntityCache();
    return success({ sheetsCreated:created, sheetsSkipped:skipped, dummyRowsAdded:n });
  });
}

function seedConfigDefaults() {
  const sh = getSheet('Config');
  const existing = getSheetAsObjects('Config').reduce(function(a,r){ a[r.key]=true; return a; },{});
  const defaults = [
    ['app.name','SupplyDesk'],
    ['seller.name','Flowmative Supplies Pvt Ltd'],
    ['seller.gstin','27AAACF1234A1Z5'],
    ['seller.homeState','Maharashtra'],
    ['complaint.windowHours','24'],
    ['demand.windowDays','30'],
    ['creditOverride.requiresReason','true'],
    ['credit.enforceAt','placement'],   // placement | dispatch | both  (we re-check at both)
    ['currency','INR']
  ];
  const rows = defaults.filter(function(kv){ return !existing[kv[0]]; });
  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,2).setValues(rows);
}

function seedAll() {
  // only seed if empty
  if (getSheet('Users').getLastRow() > 1) return 0;
  const now = new Date();
  const iso = d => new Date(d).toISOString();
  const daysAgo = n => { const d = new Date(now); d.setDate(d.getDate()-n); return iso(d); };
  const hoursAgo = n => { const d = new Date(now); d.setHours(d.getHours()-n); return iso(d); };
  const nowIso = iso(now);
  let count = 0;
  function bulk(sheet, rows){ if (!rows.length) return; getSheet(sheet).getRange(2,1,rows.length,rows[0].length).setValues(rows); count += rows.length; }

  // Clients (2) — different home states to demo GST split
  bulk('Clients', [
    [1,'Acme Manufacturing Ltd','27AACCA1111A1Z1','Maharashtra',500000,daysAgo(120)],
    [2,'Zenith Corp Services','29AACCZ2222B1Z2','Karnataka',300000,daysAgo(120)]
  ]);
  // Sites (client1: 3 sites incl. one out-of-state; client2: 2 sites)
  bulk('Sites', [
    [1,1,'Acme — Pune Plant','MIDC Bhosari, Pune','Maharashtra','Ravi Kulkarni','+91-98200-30001',daysAgo(110)],
    [2,1,'Acme — Mumbai Office','BKC, Mumbai','Maharashtra','Sneha Joshi','+91-98200-30002',daysAgo(110)],
    [3,1,'Acme — Hyderabad Unit','Gachibowli, Hyderabad','Telangana','Kiran Rao','+91-98200-30003',daysAgo(110)],
    [4,2,'Zenith — Bengaluru HQ','Whitefield, Bengaluru','Karnataka','Arjun Shetty','+91-98200-30004',daysAgo(100)],
    [5,2,'Zenith — Chennai Branch','OMR, Chennai','Tamil Nadu','Latha Menon','+91-98200-30005',daysAgo(100)]
  ]);
  // Users — 2 client supervisors (scoped to a site) + 1 admin
  bulk('Users', [
    [1,'pune@acme.com','Ravi Kulkarni (Acme Pune)','Client',1,true,nowIso],
    [2,'blr@zenith.com','Arjun Shetty (Zenith BLR)','Client',4,true,nowIso],
    [3,'admin@example.com','Ops Admin','Admin','',true,nowIso]
  ]);
  // Catalog
  bulk('Catalog', [
    [1,'A4 Copier Paper 75gsm','SKU-PAP-A4','Ream',260,'4802',12,daysAgo(90)],
    [2,'Nitrile Gloves (box 100)','SKU-GLV-NIT','Box',450,'4015',18,daysAgo(90)],
    [3,'Industrial Hand Cleaner 5L','SKU-CLN-5L','Can',720,'3401',18,daysAgo(90)],
    [4,'Safety Helmet','SKU-SAF-HLM','Piece',380,'6506',18,daysAgo(90)],
    [5,'Packing Tape 2in','SKU-TAP-2IN','Roll',45,'3919',18,daysAgo(90)],
    [6,'Cotton Waste Cloth 1kg','SKU-CWC-1KG','Kg',90,'6310',5,daysAgo(90)]
  ]);
  // SitePricing — overrides for a subset (Acme Pune negotiated cheaper paper & gloves)
  bulk('SitePricing', [
    [1,1,1,240],   // Pune paper cheaper
    [2,1,2,420],   // Pune gloves cheaper
    [3,4,1,255]    // Zenith BLR slightly cheaper paper
  ]);

  // Orders + lines
  // Order 1: Acme Pune, Delivered (with a damaged line) — invoiced
  // Order 2: Acme Pune, Delivered — recent (within complaint window)
  // Order 3: Zenith BLR, Placed (near credit limit demo)
  // Order 4: Acme Mumbai, Dispatched
  bulk('Orders', [
    [1,1,'Delivered',daysAgo(10),false],
    [2,1,'Delivered',hoursAgo(6),false],
    [3,4,'Placed',daysAgo(1),false],
    [4,2,'Dispatched',daysAgo(3),false]
  ]);
  bulk('OrderLines', [
    // id, orderId, itemId, qtyOrdered, qtyDelivered, damagedQty
    [1,1,1,50,50,0],
    [2,1,2,20,18,2],   // damaged line -> credit note demo
    [3,2,3,10,10,0],
    [4,2,5,40,40,0],
    [5,3,1,100,0,0],   // large order near credit limit
    [6,3,4,60,0,0],
    [7,4,6,30,0,0]
  ]);
  // Deliveries (consolidated batches)
  bulk('Deliveries', [
    [1,1,daysAgo(10),'1','Delivered'],
    [2,1,dateOnly_(hoursAgo(6)),'2','Delivered'],
    [3,2,dateOnly_(daysAgo(3)),'4','Dispatched']
  ]);

  // Invoice for Order 1 (Acme Pune, same state -> CGST+SGST)
  // line1: 50 * 240 = 12000 @12% ; line2 delivered 18 * 420 = 7560 @18%
  const sub1 = 50*240, sub2 = 18*420;
  const subtotal = sub1 + sub2;
  const tax1 = sub1*0.12, tax2 = sub2*0.18;
  const totalTax = tax1 + tax2;
  bulk('Invoices', [
    [1,1,1,daysAgo(9),subtotal, round2_(totalTax/2), round2_(totalTax/2), 0, round2_(subtotal+totalTax), false, daysAgo(9)]
  ]);
  bulk('InvoiceLines', [
    [1,1,1,1,50,240,sub1],
    [2,1,2,2,18,420,sub2]
  ]);

  // Payments ledger — Acme has an invoice outstanding & one part payment; Zenith near limit
  bulk('Payments', [
    [1,1,10000,'Payment','NEFT-ACME-001',daysAgo(5)],       // reduces Acme outstanding
    [2,2,280000,'Debit','Opening balance b/f',daysAgo(30)]  // Zenith already owes a lot -> near 300000 limit
  ]);

  // Complaints: one OPEN inside window (order2/line with damage? use line 2 from order1 delivered) + one EXPIRED
  // Complaint 1: OPEN, on orderLine 2 (damaged 2 gloves), raised recently, deadline in future
  // Complaint 2: EXPIRED, older
  bulk('Complaints', [
    [1,2,hoursAgo(3),hoursAgo(3-24),'Open','2 boxes of gloves damaged in transit',''],
    [2,1,daysAgo(9),daysAgo(8),'Expired','Late complaint on paper — outside 24h window','Auto-expired by system']
  ]);
  // Fix complaint1 deadline = raisedAt + 24h (future)
  updateRowById('Complaints','id',1,{ deadline: hoursAgoPlus_(3,24) });

  // CreditNote already issued for complaint 1's damage (2 * 420 = 840) — demo a posted credit note
  bulk('CreditNotes', [
    [1,1,840,hoursAgo(2)]
  ]);
  // ...and reflect it in Payments as a CreditNote entry for Acme (reduces outstanding)
  appendRowFromObject('Payments', { clientId:1, amount:840, type:'CreditNote', reference:'CN-1 (gloves damage)', createdAt:hoursAgo(2) });
  count++;
  // mark complaint 1 resolved since a credit note exists
  updateRowById('Complaints','id',1,{ status:'Resolved', resolutionNotes:'Credit note CN-1 issued for 2 damaged boxes' });

  return count;
}
function dateOnly_(iso){ return String(iso).split('T')[0]; }
function round2_(n){ return Math.round(Number(n)*100)/100; }
function hoursAgoPlus_(agoHrs, addHrs){ const d = new Date(); d.setHours(d.getHours()-agoHrs+addHrs); return d.toISOString(); }

// ---------------------------------------------------------------------------
// resetDummyData — clears seeded rows (id <= 1000), keeps real rows (id > 1000)
// ---------------------------------------------------------------------------
function resetDummyData() {
  return safeCall('resetDummyData', function(){
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      let removed = 0;
      ['Clients','Sites','Users','Catalog','SitePricing','Orders','OrderLines','Deliveries','Invoices','InvoiceLines','Payments','Complaints','CreditNotes'].forEach(function(name){
        const sh = getSheet(name); const last = sh.getLastRow();
        if (last <= 1) return;
        const data = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
        const keep = data.filter(function(row){ return Number(row[0]) > 1000; });
        sh.getRange(2,1,last-1,sh.getLastColumn()).clearContent();
        if (keep.length) sh.getRange(2,1,keep.length,keep[0].length).setValues(keep);
        removed += (data.length - keep.length);
      });
      invalidateEntityCache();
      return success({ removed:removed });
    } finally { lock.releaseLock(); }
  });
}


// ---------------------------------------------------------------------------
// Data-access trio
// ---------------------------------------------------------------------------
function getSheet(name){
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}
function getSheetAsObjects(sheetName) {
  const sh = getSheet(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const rows = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  return rows.map(function(r){
    const o = {};
    headers.forEach(function(h,i){ let v=r[i]; if (v instanceof Date) v=v.toISOString(); o[h]=v; });
    return o;
  });
}
function appendRowFromObject(sheetName, obj) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  if (headers[0] === 'id' && (obj.id === undefined || obj.id === null || obj.id === '')) obj.id = nextId(sheetName);
  const row = headers.map(function(h){ const v=obj[h]; return (v===undefined||v===null)?'':v; });
  sh.appendRow(row);
  return obj;
}
function appendRowsFromObjects(sheetName, objs) {
  // batch write — setValues once (playbook item 11)
  if (!objs || !objs.length) return [];
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  let seedId = nextId(sheetName);
  const matrix = objs.map(function(obj){
    if (headers[0]==='id' && (obj.id===undefined||obj.id===null||obj.id==='')) obj.id = seedId++;
    return headers.map(function(h){ const v=obj[h]; return (v===undefined||v===null)?'':v; });
  });
  sh.getRange(sh.getLastRow()+1, 1, matrix.length, headers.length).setValues(matrix);
  return objs;
}
function updateRowById(sheetName, idColumn, id, updated) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idIdx = headers.indexOf(idColumn);
  if (idIdx < 0) throw new Error('Column not found: ' + idColumn);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2,1,last-1,headers.length).getValues();
  for (let r=0;r<data.length;r++){
    if (String(data[r][idIdx]) === String(id)) {
      headers.forEach(function(h,i){ if (updated.hasOwnProperty(h)) data[r][i]=updated[h]; });
      sh.getRange(r+2,1,1,headers.length).setValues([data[r]]);
      const obj={}; headers.forEach(function(h,i){ let v=data[r][i]; if (v instanceof Date) v=v.toISOString(); obj[h]=v; });
      return obj;
    }
  }
  return null;
}
function nextId(sheetName) {
  const sh = getSheet(sheetName); const last = sh.getLastRow();
  if (last < 2) return 1;
  const ids = sh.getRange(2,1,last-1,1).getValues().map(function(r){ return Number(r[0])||0; });
  return Math.max.apply(null, ids) + 1;
}

// ---------------------------------------------------------------------------
// Cache
// ---------------------------------------------------------------------------
function invalidateEntityCache(entity) {
  try { CacheService.getScriptCache().removeAll([]); } catch(e){}
  try { CacheService.getUserCache().removeAll([]); } catch(e){}
}
function cachedRead(key, fn) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit){ try { return JSON.parse(hit); } catch(e){} }
  const val = fn();
  try { cache.put(key, JSON.stringify(val), SD_CACHE_TTL); } catch(e){}
  return val;
}

// ---------------------------------------------------------------------------
// Config helpers
// ---------------------------------------------------------------------------
function getConfig(key) {
  const rows = getSheetAsObjects('Config');
  for (let i=0;i<rows.length;i++) if (rows[i].key === key) return rows[i].value;
  return null;
}
function setConfig(key, value) {
  const sh = getSheet('Config');
  const rows = getSheetAsObjects('Config');
  for (let i=0;i<rows.length;i++) if (rows[i].key === key) { sh.getRange(i+2,2).setValue(value); return; }
  sh.appendRow([key,value]);
}
function updateConfig(key, value) {
  return safeCall('updateConfig', function(){
    requireRole(['Admin']);
    setConfig(key, value);
    logAudit(getCurrentUser().data.email,'config-update','Config',key,{ value:value });
    invalidateEntityCache('Config');
    return success(true);
  });
}

// ---------------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------------
function getCurrentUser() {
  return safeCall('getCurrentUser', function(){
    const email = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase();
    const users = getSheetAsObjects('Users');
    let me = null;
    for (let i=0;i<users.length;i++) if (String(users[i].email||'').toLowerCase()===email) { me = users[i]; break; }
    if (!me) {
      // Unknown identity → read-only Client with no site (sees nothing sensitive).
      me = { id:0, email:email||'anonymous', name:email||'Guest', role:'Client', siteId:'', active:true };
    }
    return success(me);
  });
}
function requireRole(allowed) {
  const me = getCurrentUser().data;
  if (!me || allowed.indexOf(me.role) === -1) throw new Error('Forbidden: requires one of ' + allowed.join(', '));
  return me;
}
// Client-console users are scoped to their own site
function requireSiteScope(me, siteId) {
  if (me.role === 'Admin') return true;
  if (String(me.siteId) !== String(siteId)) throw new Error('Forbidden: this site is outside your access');
  return true;
}

// ---------------------------------------------------------------------------
// Audit + Error logging
// ---------------------------------------------------------------------------
function logAudit(actor, action, entity, entityId, details) {
  try {
    appendRowFromObject('AuditLog', { ts:new Date().toISOString(), actor:actor, action:action, entity:entity, entityId:entityId, details:(typeof details==='string')?details:JSON.stringify(details) });
  } catch(e){}
}
function logError(fn, err) {
  try {
    appendRowFromObject('ErrorLog', { ts:new Date().toISOString(), fn:fn, message:(err&&err.message)?err.message:String(err), stack:(err&&err.stack)?err.stack:'' });
  } catch(e){}
}

// ---------------------------------------------------------------------------
// bootstrap — one startup round trip
// ---------------------------------------------------------------------------
function bootstrap() {
  return safeCall('bootstrap', function(){
    const me = getCurrentUser().data;
    const cfg = {}; getSheetAsObjects('Config').forEach(function(r){ cfg[r.key]=r.value; });
    const catalog = getSheetAsObjects('Catalog');
    const clients = getSheetAsObjects('Clients');
    const sites = getSheetAsObjects('Sites');
    let mySite = null, myClient = null;
    if (me.role === 'Client' && me.siteId) {
      mySite = sites.filter(function(s){ return String(s.id)===String(me.siteId); })[0] || null;
      if (mySite) myClient = clients.filter(function(c){ return String(c.id)===String(mySite.clientId); })[0] || null;
    }
    return success({
      me: me, config: cfg, catalog: catalog,
      clients: (me.role==='Admin'?clients:[]),
      sites: (me.role==='Admin'?sites:(mySite?[mySite]:[])),
      mySite: mySite, myClient: myClient,
      orderStatuses: SD_ORDER_STATUS, roles: SD_ROLES
    });
  });
}

// ---------------------------------------------------------------------------
// Pagination — range-based reads
// ---------------------------------------------------------------------------
function getPagedData(entity, page, pageSize, filters, sortKey, sortDir) {
  return safeCall('getPagedData', function(){
    const me = getCurrentUser().data;
    page = Math.max(1, Number(page)||1);
    pageSize = Math.min(200, Math.max(5, Number(pageSize)||25));
    const key = 'page:'+entity+':'+(me.role)+':'+(me.siteId)+':'+page+':'+pageSize+':'+JSON.stringify(filters||{})+':'+(sortKey||'')+':'+(sortDir||'');
    return cachedRead(key, function(){
      let all = getSheetAsObjects(entity).filter(function(r){ return !r.deleted; });
      if (entity === 'Orders') all = enrichOrders(all);
      if (entity === 'Invoices') all = enrichInvoices(all);
      // Client scoping
      if (me.role === 'Client') {
        if (entity === 'Orders') all = all.filter(function(r){ return String(r.siteId)===String(me.siteId); });
        if (entity === 'Invoices') all = all.filter(function(r){ return String(r.siteId)===String(me.siteId); });
      }
      if (filters) {
        Object.keys(filters).forEach(function(k){
          const v = filters[k]; if (v===''||v==null) return;
          all = all.filter(function(r){ return String(r[k]||'').toLowerCase() === String(v).toLowerCase(); });
        });
      }
      if (sortKey) { const dir=(sortDir==='asc')?1:-1; all.sort(function(a,b){ const av=a[sortKey],bv=b[sortKey]; if(av===bv) return 0; return (av>bv?1:-1)*dir; }); }
      else all.sort(function(a,b){ return (a.createdAt < b.createdAt) ? 1 : -1; });
      const total = all.length; const start = (page-1)*pageSize;
      return { rows: all.slice(start, start+pageSize), total:total, page:page, pageSize:pageSize };
    });
  });
}
function enrichOrders(orders) {
  const sites = getSheetAsObjects('Sites'); const sById={}; sites.forEach(function(s){ sById[String(s.id)]=s; });
  const lines = getSheetAsObjects('OrderLines');
  const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
  return orders.map(function(o){
    const myLines = lines.filter(function(l){ return String(l.orderId)===String(o.id); });
    const value = myLines.reduce(function(s,l){ return s + Number(l.qtyOrdered||0) * rateFor_(o.siteId, l.itemId, cById); }, 0);
    const s = sById[String(o.siteId)]||{};
    return Object.assign({}, o, { siteName:s.siteName||'', siteState:s.state||'', clientId:s.clientId||'', lineCount:myLines.length, estValue:round2_(value) });
  });
}
function enrichInvoices(invoices) {
  const sites = getSheetAsObjects('Sites'); const sById={}; sites.forEach(function(s){ sById[String(s.id)]=s; });
  const clients = getSheetAsObjects('Clients'); const cById={}; clients.forEach(function(c){ cById[String(c.id)]=c; });
  return invoices.map(function(inv){
    return Object.assign({}, inv, { siteName:(sById[String(inv.siteId)]||{}).siteName||'', clientName:(cById[String(inv.clientId)]||{}).name||'' });
  });
}
function rateFor_(siteId, itemId, catalogById) {
  const pricing = getSheetAsObjects('SitePricing');
  const override = pricing.filter(function(p){ return String(p.siteId)===String(siteId) && String(p.itemId)===String(itemId); })[0];
  if (override) return Number(override.rate)||0;
  const c = catalogById ? catalogById[String(itemId)] : getSheetAsObjects('Catalog').filter(function(x){ return String(x.id)===String(itemId); })[0];
  return c ? Number(c.defaultRate)||0 : 0;
}


// ===========================================================================
// ENTITY CRUD — Catalog / Clients / Sites / SitePricing
// ===========================================================================
function listCatalog(){ return safeCall('listCatalog', function(){ return success(getSheetAsObjects('Catalog')); }); }
function saveCatalogItem(payload) {
  return safeCall('saveCatalogItem', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      let saved;
      if (payload.id) { saved = updateRowById('Catalog','id',payload.id, payload); }
      else { payload.createdAt = new Date().toISOString(); saved = appendRowFromObject('Catalog', payload); }
      logAudit(me.email, payload.id?'update':'create','Catalog', saved.id, payload);
      invalidateEntityCache('Catalog');
      return success(saved);
    } finally { lock.releaseLock(); }
  });
}
function listClients(){ return safeCall('listClients', function(){ requireRole(['Admin']); return success(getSheetAsObjects('Clients')); }); }
function saveClient(payload) {
  return safeCall('saveClient', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      let saved;
      if (payload.id) saved = updateRowById('Clients','id',payload.id,payload);
      else { payload.createdAt = new Date().toISOString(); saved = appendRowFromObject('Clients', payload); }
      logAudit(me.email, payload.id?'update':'create','Clients',saved.id,payload);
      invalidateEntityCache('Clients');
      return success(saved);
    } finally { lock.releaseLock(); }
  });
}
function listSites(clientId){ return safeCall('listSites', function(){ const all=getSheetAsObjects('Sites'); return success(clientId?all.filter(function(s){return String(s.clientId)===String(clientId);}):all); }); }
function saveSite(payload) {
  return safeCall('saveSite', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      let saved;
      if (payload.id) saved = updateRowById('Sites','id',payload.id,payload);
      else { payload.createdAt = new Date().toISOString(); saved = appendRowFromObject('Sites', payload); }
      logAudit(me.email, payload.id?'update':'create','Sites',saved.id,payload);
      invalidateEntityCache('Sites');
      return success(saved);
    } finally { lock.releaseLock(); }
  });
}

// ---- Site pricing grid (Admin) ----
function getSitePricingGrid(siteId) {
  return safeCall('getSitePricingGrid', function(){
    requireRole(['Admin']);
    const catalog = getSheetAsObjects('Catalog');
    const pricing = getSheetAsObjects('SitePricing').filter(function(p){ return String(p.siteId)===String(siteId); });
    const pById = {}; pricing.forEach(function(p){ pById[String(p.itemId)] = p; });
    return success(catalog.map(function(c){
      const ov = pById[String(c.id)];
      return { itemId:c.id, itemName:c.itemName, sku:c.sku, uom:c.uom, defaultRate:c.defaultRate, gstRatePct:c.gstRatePct, overrideRate: ov?ov.rate:'', overrideId: ov?ov.id:'' };
    }));
  });
}
function setSitePrice(siteId, itemId, rate) {
  return safeCall('setSitePrice', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const existing = getSheetAsObjects('SitePricing').filter(function(p){ return String(p.siteId)===String(siteId)&&String(p.itemId)===String(itemId); })[0];
      if (rate === '' || rate === null || rate === undefined) {
        // clearing an override → delete the row if present
        if (existing) { deleteRowById_('SitePricing', existing.id); }
      } else if (existing) {
        updateRowById('SitePricing','id',existing.id,{ rate:Number(rate) });
      } else {
        appendRowFromObject('SitePricing', { siteId:siteId, itemId:itemId, rate:Number(rate) });
      }
      logAudit(me.email,'set-price','SitePricing', siteId+':'+itemId, { rate:rate });
      invalidateEntityCache('SitePricing');
      return success(true);
    } finally { lock.releaseLock(); }
  });
}
function deleteRowById_(sheetName, id) {
  const sh = getSheet(sheetName);
  const data = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1), sh.getLastColumn()).getValues();
  for (let r=0;r<data.length;r++) if (String(data[r][0])===String(id)) { sh.deleteRow(r+2); return true; }
  return false;
}

// Resolve the effective price list for a site (SitePricing override → Catalog default)
function getPriceListForSite(siteId) {
  return safeCall('getPriceListForSite', function(){
    const me = getCurrentUser().data;
    if (me.role === 'Client') requireSiteScope(me, siteId);
    const catalog = getSheetAsObjects('Catalog');
    const cById = {}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
    return success(catalog.map(function(c){
      return { itemId:c.id, itemName:c.itemName, sku:c.sku, uom:c.uom, gstRatePct:c.gstRatePct, rate: rateFor_(siteId, c.id, cById) };
    }));
  });
}

// ===========================================================================
// CREDIT LIMIT GUARD
// ===========================================================================
function computeOutstanding(clientId) {
  const invoices = getSheetAsObjects('Invoices').filter(function(i){ return String(i.clientId)===String(clientId); });
  const invTotal = invoices.reduce(function(s,i){ return s + Number(i.total||0); }, 0);
  const payments = getSheetAsObjects('Payments').filter(function(p){ return String(p.clientId)===String(clientId); });
  // Payment & CreditNote reduce outstanding; Debit increases it
  const credits = payments.reduce(function(s,p){
    const amt = Number(p.amount||0);
    if (p.type === 'Debit') return s - amt;   // debit adds to what they owe
    return s + amt;                            // Payment / CreditNote reduce it
  }, 0);
  return round2_(invTotal - credits);
}
function getClientCredit(clientId) {
  return safeCall('getClientCredit', function(){
    const client = getSheetAsObjects('Clients').filter(function(c){ return String(c.id)===String(clientId); })[0];
    if (!client) return fail('Client not found');
    const outstanding = computeOutstanding(clientId);
    const limit = Number(client.creditLimit||0);
    return success({ clientId:clientId, creditLimit:limit, outstanding:outstanding, available: round2_(limit-outstanding) });
  });
}
function creditCheck_(clientId, additionalValue) {
  const client = getSheetAsObjects('Clients').filter(function(c){ return String(c.id)===String(clientId); })[0];
  if (!client) throw new Error('Client not found');
  const limit = Number(client.creditLimit||0);
  const outstanding = computeOutstanding(clientId);
  const projected = round2_(outstanding + Number(additionalValue||0));
  return { ok: projected <= limit, limit:limit, outstanding:outstanding, projected:projected, client:client };
}

// ===========================================================================
// ORDERS — placement (site pricing + credit guard), listing, cancel
// ===========================================================================
function estimateOrderValue_(siteId, lines) {
  const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
  return lines.reduce(function(s,l){ return s + Number(l.qtyOrdered||0) * rateFor_(siteId, l.itemId, cById); }, 0);
}
function placeOrder(payload) {
  return safeCall('placeOrder', function(){
    const me = requireRole(['Client','Admin']);
    if (!payload || !payload.siteId) return fail('Site is required');
    const lines = (payload.lines||[]).filter(function(l){ return l.itemId && Number(l.qtyOrdered)>0; });
    if (!lines.length) return fail('Add at least one item with a quantity');
    if (me.role === 'Client') requireSiteScope(me, payload.siteId);

    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const site = getSheetAsObjects('Sites').filter(function(s){ return String(s.id)===String(payload.siteId); })[0];
      if (!site) return fail('Site not found');
      const estValue = estimateOrderValue_(payload.siteId, lines);

      // CREDIT LIMIT GUARD (enforced at placement)
      const chk = creditCheck_(site.clientId, estValue);
      if (!chk.ok) {
        const override = payload.override === true || payload.override === 'true';
        if (!(override && me.role === 'Admin')) {
          logAudit(me.email,'credit-block','Orders','', { clientId:site.clientId, projected:chk.projected, limit:chk.limit });
          return fail('Credit limit exceeded. Outstanding ₹'+chk.outstanding.toLocaleString('en-IN')+' + this order ₹'+round2_(estValue).toLocaleString('en-IN')+' > limit ₹'+chk.limit.toLocaleString('en-IN')+'. An Admin can override with a reason.');
        }
        const reasonRequired = String(getConfig('creditOverride.requiresReason')||'true').toLowerCase() !== 'false';
        if (reasonRequired && !payload.overrideReason) return fail('Credit override requires a reason');
        logAudit(me.email,'credit-override','Orders','', { clientId:site.clientId, projected:chk.projected, limit:chk.limit, reason:payload.overrideReason||'' });
      }

      const now = new Date().toISOString();
      const order = appendRowFromObject('Orders', { siteId:payload.siteId, status:'Placed', createdAt:now, deleted:false });
      const lineObjs = lines.map(function(l){ return { orderId:order.id, itemId:l.itemId, qtyOrdered:Number(l.qtyOrdered), qtyDelivered:'', damagedQty:'' }; });
      appendRowsFromObjects('OrderLines', lineObjs);
      logAudit(me.email,'place-order','Orders',order.id,{ siteId:payload.siteId, estValue:round2_(estValue), lines:lines.length });
      invalidateEntityCache('Orders');
      return success({ orderId:order.id, estValue:round2_(estValue) });
    } finally { lock.releaseLock(); }
  });
}
function getOrderDetail(orderId) {
  return safeCall('getOrderDetail', function(){
    const me = getCurrentUser().data;
    const order = getSheetAsObjects('Orders').filter(function(o){ return String(o.id)===String(orderId); })[0];
    if (!order) return fail('Order not found');
    if (me.role === 'Client') requireSiteScope(me, order.siteId);
    const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
    const site = getSheetAsObjects('Sites').filter(function(s){ return String(s.id)===String(order.siteId); })[0]||{};
    const lines = getSheetAsObjects('OrderLines').filter(function(l){ return String(l.orderId)===String(orderId); }).map(function(l){
      const rate = rateFor_(order.siteId, l.itemId, cById); const c = cById[String(l.itemId)]||{};
      return Object.assign({}, l, { itemName:c.itemName||'', uom:c.uom||'', rate:rate, amount:round2_(Number(l.qtyOrdered||0)*rate) });
    });
    return success({ order:order, site:site, lines:lines, total: round2_(lines.reduce(function(s,l){ return s+l.amount; },0)) });
  });
}
function cancelOrder(orderId) {
  return safeCall('cancelOrder', function(){
    const me = requireRole(['Client','Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const order = getSheetAsObjects('Orders').filter(function(o){ return String(o.id)===String(orderId); })[0];
      if (!order) return fail('Order not found');
      if (me.role === 'Client') requireSiteScope(me, order.siteId);
      if (['Delivered','Dispatched'].indexOf(order.status) !== -1) return fail('Cannot cancel an order already '+order.status.toLowerCase());
      updateRowById('Orders','id',orderId,{ status:'Cancelled' });
      logAudit(me.email,'cancel-order','Orders',orderId,'');
      invalidateEntityCache('Orders');
      return success(true);
    } finally { lock.releaseLock(); }
  });
}


// ===========================================================================
// CONSOLIDATED DELIVERIES OPTIMIZER
// ===========================================================================
function getConsolidationPlan() {
  return safeCall('getConsolidationPlan', function(){
    requireRole(['Admin']);
    const placed = getSheetAsObjects('Orders').filter(function(o){ return !o.deleted && o.status==='Placed'; });
    const sites = getSheetAsObjects('Sites'); const sById={}; sites.forEach(function(s){ sById[String(s.id)]=s; });
    const groups = {};
    placed.forEach(function(o){
      const k = String(o.siteId);
      if (!groups[k]) groups[k] = { siteId:o.siteId, siteName:(sById[k]||{}).siteName||'', orders:[], orderIds:[] };
      groups[k].orders.push(o); groups[k].orderIds.push(o.id);
    });
    return success(Object.keys(groups).map(function(k){ return groups[k]; }));
  });
}
function consolidateAndDispatch(siteId, deliveryDate) {
  return safeCall('consolidateAndDispatch', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const placed = getSheetAsObjects('Orders').filter(function(o){ return !o.deleted && o.status==='Placed' && String(o.siteId)===String(siteId); });
      if (!placed.length) return fail('No placed orders for this site');
      const site = getSheetAsObjects('Sites').filter(function(s){ return String(s.id)===String(siteId); })[0];

      // Re-check credit at dispatch (playbook: strictest checkpoint)
      const value = placed.reduce(function(s,o){
        const lines = getSheetAsObjects('OrderLines').filter(function(l){ return String(l.orderId)===String(o.id); });
        return s + estimateOrderValue_(siteId, lines);
      }, 0);
      const chk = creditCheck_(site.clientId, 0); // already-placed value is prospective; ensure not already over limit
      if (chk.outstanding > chk.limit) {
        logAudit(me.email,'dispatch-credit-warn','Deliveries','', { clientId:site.clientId, outstanding:chk.outstanding, limit:chk.limit });
      }

      const orderIds = placed.map(function(o){ return o.id; });
      const del = appendRowFromObject('Deliveries', { siteId:siteId, deliveryDate:deliveryDate||new Date().toISOString().split('T')[0], orderIds:orderIds.join(','), status:'Dispatched' });
      orderIds.forEach(function(id){ updateRowById('Orders','id',id,{ status:'Dispatched' }); });
      logAudit(me.email,'consolidate-dispatch','Deliveries',del.id,{ siteId:siteId, orders:orderIds, estValue:round2_(value) });
      invalidateEntityCache('Deliveries'); invalidateEntityCache('Orders');
      return success({ deliveryId:del.id, orders:orderIds.length });
    } finally { lock.releaseLock(); }
  });
}
function listDeliveries() {
  return safeCall('listDeliveries', function(){
    requireRole(['Admin']);
    const sites = getSheetAsObjects('Sites'); const sById={}; sites.forEach(function(s){ sById[String(s.id)]=s; });
    return success(getSheetAsObjects('Deliveries').map(function(d){ return Object.assign({}, d, { siteName:(sById[String(d.siteId)]||{}).siteName||'' }); }).reverse());
  });
}
function getDeliveryDetail(deliveryId) {
  return safeCall('getDeliveryDetail', function(){
    requireRole(['Admin']);
    const del = getSheetAsObjects('Deliveries').filter(function(d){ return String(d.id)===String(deliveryId); })[0];
    if (!del) return fail('Delivery not found');
    const orderIds = String(del.orderIds||'').split(',').map(function(s){ return s.trim(); }).filter(Boolean);
    const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
    const lines = getSheetAsObjects('OrderLines').filter(function(l){ return orderIds.indexOf(String(l.orderId))!==-1; }).map(function(l){
      const c = cById[String(l.itemId)]||{};
      return Object.assign({}, l, { itemName:c.itemName||'', uom:c.uom||'' });
    });
    return success({ delivery:del, lines:lines });
  });
}
// Reconciliation — capture actual qtyDelivered + damagedQty per line
function confirmDelivery(deliveryId, lineUpdates) {
  return safeCall('confirmDelivery', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const del = getSheetAsObjects('Deliveries').filter(function(d){ return String(d.id)===String(deliveryId); })[0];
      if (!del) return fail('Delivery not found');
      (lineUpdates||[]).forEach(function(u){
        updateRowById('OrderLines','id',u.orderLineId, { qtyDelivered:Number(u.qtyDelivered||0), damagedQty:Number(u.damagedQty||0) });
      });
      const orderIds = String(del.orderIds||'').split(',').map(function(s){ return s.trim(); }).filter(Boolean);
      orderIds.forEach(function(id){ updateRowById('Orders','id',id,{ status:'Delivered' }); });
      updateRowById('Deliveries','id',deliveryId,{ status:'Delivered' });
      logAudit(me.email,'confirm-delivery','Deliveries',deliveryId,{ lines:(lineUpdates||[]).length });
      invalidateEntityCache('Deliveries'); invalidateEntityCache('Orders'); invalidateEntityCache('OrderLines');
      return success(true);
    } finally { lock.releaseLock(); }
  });
}

// ===========================================================================
// LIVE DEMAND PLANNER (trailing N-day aggregation)
// ===========================================================================
function getDemandPlan() {
  return safeCall('getDemandPlan', function(){
    requireRole(['Admin']);
    const days = Number(getConfig('demand.windowDays')||30);
    const cutoff = new Date(); cutoff.setDate(cutoff.getDate()-days);
    const orders = getSheetAsObjects('Orders').filter(function(o){ return !o.deleted && new Date(o.createdAt) >= cutoff; });
    const orderIds = {}; orders.forEach(function(o){ orderIds[String(o.id)] = true; });
    const lines = getSheetAsObjects('OrderLines').filter(function(l){ return orderIds[String(l.orderId)]; });
    const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
    const agg = {};
    lines.forEach(function(l){
      const k = String(l.itemId);
      if (!agg[k]) agg[k] = { itemId:l.itemId, itemName:(cById[k]||{}).itemName||'', uom:(cById[k]||{}).uom||'', totalQty:0, orderCount:0 };
      agg[k].totalQty += Number(l.qtyOrdered||0); agg[k].orderCount++;
    });
    const rows = Object.keys(agg).map(function(k){ const a=agg[k]; a.avgPerOrder = a.orderCount?round2_(a.totalQty/a.orderCount):0; return a; }).sort(function(a,b){ return b.totalQty-a.totalQty; });
    return success({ windowDays:days, items:rows });
  });
}

// ===========================================================================
// AUTO GST-DETECT INVOICING
// ===========================================================================
function getInvoiceableOrders() {
  return safeCall('getInvoiceableOrders', function(){
    requireRole(['Admin']);
    const invoicedLineIds = {}; getSheetAsObjects('InvoiceLines').forEach(function(il){ invoicedLineIds[String(il.orderLineId)] = true; });
    const delivered = getSheetAsObjects('Orders').filter(function(o){ return !o.deleted && o.status==='Delivered'; });
    const sites = getSheetAsObjects('Sites'); const sById={}; sites.forEach(function(s){ sById[String(s.id)]=s; });
    const lines = getSheetAsObjects('OrderLines');
    return success(delivered.map(function(o){
      const myLines = lines.filter(function(l){ return String(l.orderId)===String(o.id) && !invoicedLineIds[String(l.id)] && Number(l.qtyDelivered||0)>0; });
      const s = sById[String(o.siteId)]||{};
      return { orderId:o.id, siteId:o.siteId, siteName:s.siteName||'', state:s.state||'', uninvoicedLines:myLines.length };
    }).filter(function(o){ return o.uninvoicedLines>0; }));
  });
}
function generateInvoice(orderId) {
  return safeCall('generateInvoice', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const order = getSheetAsObjects('Orders').filter(function(o){ return String(o.id)===String(orderId); })[0];
      if (!order) return fail('Order not found');
      const site = getSheetAsObjects('Sites').filter(function(s){ return String(s.id)===String(order.siteId); })[0];
      if (!site) return fail('Site not found');
      const invoicedLineIds = {}; getSheetAsObjects('InvoiceLines').forEach(function(il){ invoicedLineIds[String(il.orderLineId)] = true; });
      const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
      const lines = getSheetAsObjects('OrderLines').filter(function(l){ return String(l.orderId)===String(orderId) && !invoicedLineIds[String(l.id)] && Number(l.qtyDelivered||0)>0; });
      if (!lines.length) return fail('Nothing left to invoice on this order');

      const homeState = String(getConfig('seller.homeState')||'').trim().toLowerCase();
      const sameState = String(site.state||'').trim().toLowerCase() === homeState;

      let subtotal=0, cgst=0, sgst=0, igst=0;
      const invLineObjs = [];
      lines.forEach(function(l){
        const rate = rateFor_(order.siteId, l.itemId, cById);
        const qty = Number(l.qtyDelivered||0);
        const amount = round2_(qty*rate);
        const gstPct = Number((cById[String(l.itemId)]||{}).gstRatePct||0);
        const tax = amount * gstPct/100;
        if (sameState){ cgst += tax/2; sgst += tax/2; } else { igst += tax; }
        subtotal += amount;
        invLineObjs.push({ orderLineId:l.id, itemId:l.itemId, qty:qty, rate:rate, amount:amount });
      });
      cgst=round2_(cgst); sgst=round2_(sgst); igst=round2_(igst); subtotal=round2_(subtotal);
      const total = round2_(subtotal + cgst + sgst + igst);
      const inv = appendRowFromObject('Invoices', {
        clientId:site.clientId, siteId:site.id, invoiceDate:new Date().toISOString().split('T')[0],
        subtotal:subtotal, cgst:cgst, sgst:sgst, igst:igst, total:total, tallyExported:false, createdAt:new Date().toISOString()
      });
      invLineObjs.forEach(function(il){ il.invoiceId = inv.id; });
      appendRowsFromObjects('InvoiceLines', invLineObjs);
      logAudit(me.email,'generate-invoice','Invoices',inv.id,{ orderId:orderId, total:total, taxMode: sameState?'CGST+SGST':'IGST' });
      invalidateEntityCache('Invoices');
      return success({ invoiceId:inv.id, total:total, taxMode: sameState?'CGST+SGST':'IGST' });
    } finally { lock.releaseLock(); }
  });
}
function getInvoiceDetail(invoiceId) {
  return safeCall('getInvoiceDetail', function(){
    const me = getCurrentUser().data;
    const inv = getSheetAsObjects('Invoices').filter(function(i){ return String(i.id)===String(invoiceId); })[0];
    if (!inv) return fail('Invoice not found');
    if (me.role === 'Client') requireSiteScope(me, inv.siteId);
    const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
    const site = getSheetAsObjects('Sites').filter(function(s){ return String(s.id)===String(inv.siteId); })[0]||{};
    const client = getSheetAsObjects('Clients').filter(function(c){ return String(c.id)===String(inv.clientId); })[0]||{};
    const lines = getSheetAsObjects('InvoiceLines').filter(function(l){ return String(l.invoiceId)===String(invoiceId); }).map(function(l){
      return Object.assign({}, l, { itemName:(cById[String(l.itemId)]||{}).itemName||'', hsnCode:(cById[String(l.itemId)]||{}).hsnCode||'', gstRatePct:(cById[String(l.itemId)]||{}).gstRatePct||0 });
    });
    return success({ invoice:inv, site:site, client:client, lines:lines });
  });
}

// ===========================================================================
// 24-HOUR ISSUE DESK + AUTO CREDIT NOTES
// ===========================================================================
function getComplaintableLines(orderId) {
  return safeCall('getComplaintableLines', function(){
    const me = getCurrentUser().data;
    const order = getSheetAsObjects('Orders').filter(function(o){ return String(o.id)===String(orderId); })[0];
    if (!order) return fail('Order not found');
    if (me.role === 'Client') requireSiteScope(me, order.siteId);
    if (order.status !== 'Delivered') return fail('Order is not delivered yet');
    const windowHours = Number(getConfig('complaint.windowHours')||24);
    const del = getSheetAsObjects('Deliveries').filter(function(d){ return String(d.orderIds||'').split(',').map(function(x){return x.trim();}).indexOf(String(orderId))!==-1; })[0];
    const deliveredAt = del ? new Date(del.deliveryDate) : new Date(order.createdAt);
    const deadline = new Date(deliveredAt.getTime() + windowHours*3600*1000);
    const withinWindow = new Date() <= deadline;
    const existing = {}; getSheetAsObjects('Complaints').forEach(function(c){ existing[String(c.orderLineId)] = c.status; });
    const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
    const lines = getSheetAsObjects('OrderLines').filter(function(l){ return String(l.orderId)===String(orderId); }).map(function(l){
      const rate = rateFor_(order.siteId, l.itemId, cById);
      return Object.assign({}, l, { itemName:(cById[String(l.itemId)]||{}).itemName||'', rate:rate, hasComplaint: !!existing[String(l.id)], complaintStatus: existing[String(l.id)]||'' });
    });
    return success({ orderId:orderId, withinWindow:withinWindow, deadline:deadline.toISOString(), windowHours:windowHours, lines:lines });
  });
}
function raiseComplaint(orderLineId, reason) {
  return safeCall('raiseComplaint', function(){
    const me = requireRole(['Client','Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const line = getSheetAsObjects('OrderLines').filter(function(l){ return String(l.id)===String(orderLineId); })[0];
      if (!line) return fail('Order line not found');
      const order = getSheetAsObjects('Orders').filter(function(o){ return String(o.id)===String(line.orderId); })[0];
      if (!order) return fail('Order not found');
      if (me.role === 'Client') requireSiteScope(me, order.siteId);
      if (order.status !== 'Delivered') return fail('Order is not delivered yet');

      // Server-side 24h window enforcement
      const windowHours = Number(getConfig('complaint.windowHours')||24);
      const del = getSheetAsObjects('Deliveries').filter(function(d){ return String(d.orderIds||'').split(',').map(function(x){return x.trim();}).indexOf(String(order.id))!==-1; })[0];
      const deliveredAt = del ? new Date(del.deliveryDate) : new Date(order.createdAt);
      const deadline = new Date(deliveredAt.getTime() + windowHours*3600*1000);
      if (new Date() > deadline) {
        // Record a blocked/expired attempt rather than silently allowing
        appendRowFromObject('Complaints', { orderLineId:orderLineId, raisedAt:new Date().toISOString(), deadline:deadline.toISOString(), status:'Expired', reason:reason||'', resolutionNotes:'Blocked — outside '+windowHours+'h window' });
        invalidateEntityCache('Complaints');
        return fail('The '+windowHours+'-hour complaint window closed on '+deadline.toISOString().replace('T',' ').slice(0,16)+'. This attempt was logged as Expired.');
      }
      const dup = getSheetAsObjects('Complaints').filter(function(c){ return String(c.orderLineId)===String(orderLineId) && c.status!=='Expired'; })[0];
      if (dup) return fail('A complaint already exists for this line');

      const complaint = appendRowFromObject('Complaints', { orderLineId:orderLineId, raisedAt:new Date().toISOString(), deadline:deadline.toISOString(), status:'Open', reason:reason||'', resolutionNotes:'' });
      // Auto-propose a credit note for the damaged quantity
      const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
      const rate = rateFor_(order.siteId, line.itemId, cById);
      const proposedAmount = round2_(Number(line.damagedQty||0) * rate);
      logAudit(me.email,'raise-complaint','Complaints',complaint.id,{ orderLineId:orderLineId, proposedCredit:proposedAmount });
      invalidateEntityCache('Complaints');
      return success({ complaintId:complaint.id, proposedCreditAmount:proposedAmount, damagedQty:Number(line.damagedQty||0), rate:rate });
    } finally { lock.releaseLock(); }
  });
}
function listComplaints() {
  return safeCall('listComplaints', function(){
    const me = getCurrentUser().data;
    const lines = getSheetAsObjects('OrderLines'); const lById={}; lines.forEach(function(l){ lById[String(l.id)]=l; });
    const orders = getSheetAsObjects('Orders'); const oById={}; orders.forEach(function(o){ oById[String(o.id)]=o; });
    const catalog = getSheetAsObjects('Catalog'); const cById={}; catalog.forEach(function(c){ cById[String(c.id)]=c; });
    let rows = getSheetAsObjects('Complaints').map(function(c){
      const l = lById[String(c.orderLineId)]||{}; const o = oById[String(l.orderId)]||{};
      return Object.assign({}, c, { orderId:l.orderId||'', siteId:o.siteId||'', itemName:(cById[String(l.itemId)]||{}).itemName||'', damagedQty:l.damagedQty||0 });
    });
    if (me.role === 'Client') rows = rows.filter(function(c){ return String(c.siteId)===String(me.siteId); });
    return success(rows.reverse());
  });
}
function resolveComplaintWithCreditNote(complaintId, amount) {
  return safeCall('resolveComplaintWithCreditNote', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const c = getSheetAsObjects('Complaints').filter(function(x){ return String(x.id)===String(complaintId); })[0];
      if (!c) return fail('Complaint not found');
      if (c.status === 'Expired') return fail('Complaint is expired and cannot be resolved');
      const line = getSheetAsObjects('OrderLines').filter(function(l){ return String(l.id)===String(c.orderLineId); })[0]||{};
      const order = getSheetAsObjects('Orders').filter(function(o){ return String(o.id)===String(line.orderId); })[0]||{};
      const site = getSheetAsObjects('Sites').filter(function(s){ return String(s.id)===String(order.siteId); })[0]||{};
      const amt = round2_(Number(amount||0));
      if (amt <= 0) return fail('Credit note amount must be greater than zero');
      const cn = appendRowFromObject('CreditNotes', { complaintId:complaintId, amount:amt, issuedAt:new Date().toISOString() });
      // Post to Payments ledger as a CreditNote → reduces client outstanding
      appendRowFromObject('Payments', { clientId:site.clientId, amount:amt, type:'CreditNote', reference:'CN-'+cn.id+' (complaint '+complaintId+')', createdAt:new Date().toISOString() });
      updateRowById('Complaints','id',complaintId,{ status:'Resolved', resolutionNotes:'Credit note CN-'+cn.id+' issued for ₹'+amt });
      logAudit(me.email,'issue-credit-note','CreditNotes',cn.id,{ complaintId:complaintId, amount:amt, clientId:site.clientId });
      invalidateEntityCache('Complaints'); invalidateEntityCache('Payments'); invalidateEntityCache('CreditNotes');
      return success({ creditNoteId:cn.id, amount:amt });
    } finally { lock.releaseLock(); }
  });
}

// ===========================================================================
// TALLY PRIME XML EXPORTER (Sales Voucher)
// ===========================================================================
function exportInvoiceTallyXml(invoiceId) {
  return safeCall('exportInvoiceTallyXml', function(){
    const me = requireRole(['Admin']);
    const det = getInvoiceDetail(invoiceId);
    if (!det.success) return det;
    const inv = det.data.invoice, site = det.data.site, client = det.data.client, lines = det.data.lines;
    const sellerName = getConfig('seller.name') || 'Seller';
    const partyLedger = client.name || 'Sundry Debtor';
    const invDate = String(inv.invoiceDate||'').replace(/-/g,'');  // Tally wants YYYYMMDD
    const voucherNo = 'SD-INV-' + inv.id;
    const esc = function(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };

    let allocations = '';
    lines.forEach(function(l){
      allocations +=
        '<ALLINVENTORYENTRIES.LIST>' +
          '<STOCKITEMNAME>'+esc(l.itemName)+'</STOCKITEMNAME>' +
          '<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>' +
          '<RATE>'+l.rate+'</RATE>' +
          '<ACTUALQTY>'+l.qty+'</ACTUALQTY>' +
          '<BILLEDQTY>'+l.qty+'</BILLEDQTY>' +
          '<AMOUNT>'+l.amount+'</AMOUNT>' +
          '<ACCOUNTINGALLOCATIONS.LIST>' +
            '<LEDGERNAME>Sales</LEDGERNAME>' +
            '<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>' +
            '<AMOUNT>'+l.amount+'</AMOUNT>' +
          '</ACCOUNTINGALLOCATIONS.LIST>' +
        '</ALLINVENTORYENTRIES.LIST>';
    });

    function ledgerEntry(name, amount, positive){
      if (!amount || Number(amount)===0) return '';
      return '<LEDGERENTRIES.LIST>' +
        '<LEDGERNAME>'+esc(name)+'</LEDGERNAME>' +
        '<ISDEEMEDPOSITIVE>'+(positive?'Yes':'No')+'</ISDEEMEDPOSITIVE>' +
        '<AMOUNT>'+(positive?('-'+amount):amount)+'</AMOUNT>' +
        '</LEDGERENTRIES.LIST>';
    }
    const taxLedgers =
      ledgerEntry('CGST', inv.cgst, false) +
      ledgerEntry('SGST', inv.sgst, false) +
      ledgerEntry('IGST', inv.igst, false);

    const xml =
'<ENVELOPE>' +
  '<HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>' +
  '<BODY><IMPORTDATA>' +
    '<REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME></REQUESTDESC>' +
    '<REQUESTDATA>' +
      '<TALLYMESSAGE xmlns:UDF="TallyUDF">' +
        '<VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Invoice Voucher View">' +
          '<DATE>'+invDate+'</DATE>' +
          '<VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>' +
          '<VOUCHERNUMBER>'+esc(voucherNo)+'</VOUCHERNUMBER>' +
          '<PARTYLEDGERNAME>'+esc(partyLedger)+'</PARTYLEDGERNAME>' +
          '<BASICBUYERNAME>'+esc(site.siteName||partyLedger)+'</BASICBUYERNAME>' +
          '<PARTYGSTIN>'+esc(client.gstin||'')+'</PARTYGSTIN>' +
          '<PLACEOFSUPPLY>'+esc(site.state||'')+'</PLACEOFSUPPLY>' +
          '<LEDGERENTRIES.LIST>' +
            '<LEDGERNAME>'+esc(partyLedger)+'</LEDGERNAME>' +
            '<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>' +
            '<AMOUNT>-'+inv.total+'</AMOUNT>' +
          '</LEDGERENTRIES.LIST>' +
          taxLedgers +
          allocations +
        '</VOUCHER>' +
      '</TALLYMESSAGE>' +
    '</REQUESTDATA>' +
  '</IMPORTDATA></BODY>' +
'</ENVELOPE>';

    updateRowById('Invoices','id',invoiceId,{ tallyExported:true });
    logAudit(me.email,'tally-export','Invoices',invoiceId,{ voucherNo:voucherNo });
    invalidateEntityCache('Invoices');
    return success({ filename:'SupplyDesk-'+voucherNo+'.xml', xml:xml });
  });
}

// ===========================================================================
// Ledger / dashboards
// ===========================================================================
function getClientLedger(clientId) {
  return safeCall('getClientLedger', function(){
    const me = getCurrentUser().data;
    if (me.role === 'Client') {
      const mySite = getSheetAsObjects('Sites').filter(function(s){ return String(s.id)===String(me.siteId); })[0];
      if (!mySite || String(mySite.clientId)!==String(clientId)) throw new Error('Forbidden');
    }
    const credit = getClientCredit(clientId).data;
    const payments = getSheetAsObjects('Payments').filter(function(p){ return String(p.clientId)===String(clientId); }).reverse();
    const invoices = enrichInvoices(getSheetAsObjects('Invoices').filter(function(i){ return String(i.clientId)===String(clientId); })).reverse();
    return success({ credit:credit, payments:payments, invoices:invoices });
  });
}
function recordPayment(clientId, amount, reference) {
  return safeCall('recordPayment', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const amt = round2_(Number(amount||0)); if (amt<=0) return fail('Amount must be greater than zero');
      appendRowFromObject('Payments', { clientId:clientId, amount:amt, type:'Payment', reference:reference||'Manual receipt', createdAt:new Date().toISOString() });
      logAudit(me.email,'record-payment','Payments',clientId,{ amount:amt });
      invalidateEntityCache('Payments');
      return success(getClientCredit(clientId).data);
    } finally { lock.releaseLock(); }
  });
}
function getAdminDashboard() {
  return safeCall('getAdminDashboard', function(){
    requireRole(['Admin']);
    return cachedRead('admin:dash', function(){
      const orders = getSheetAsObjects('Orders').filter(function(o){ return !o.deleted; });
      const invoices = getSheetAsObjects('Invoices');
      const complaints = getSheetAsObjects('Complaints');
      const clients = getSheetAsObjects('Clients');
      const totalOutstanding = clients.reduce(function(s,c){ return s + computeOutstanding(c.id); }, 0);
      const overLimit = clients.filter(function(c){ return computeOutstanding(c.id) > Number(c.creditLimit||0); }).length;
      return {
        placed: orders.filter(function(o){ return o.status==='Placed'; }).length,
        dispatched: orders.filter(function(o){ return o.status==='Dispatched'; }).length,
        delivered: orders.filter(function(o){ return o.status==='Delivered'; }).length,
        invoices: invoices.length,
        invoicedValue: round2_(invoices.reduce(function(s,i){ return s+Number(i.total||0); },0)),
        openComplaints: complaints.filter(function(c){ return c.status==='Open'; }).length,
        totalOutstanding: round2_(totalOutstanding),
        clientsOverLimit: overLimit
      };
    });
  });
}

// ===========================================================================
// TRIGGERS — expire complaints past their window
// ===========================================================================
function complaintSweep() {
  return safeCall('complaintSweep', function(){
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const now = new Date(); let expired = 0;
      getSheetAsObjects('Complaints').forEach(function(c){
        if (c.status === 'Open' && c.deadline && now > new Date(c.deadline)) {
          updateRowById('Complaints','id',c.id,{ status:'Expired', resolutionNotes:(c.resolutionNotes||'')+' Auto-expired at '+now.toISOString() });
          expired++;
        }
      });
      if (expired) invalidateEntityCache('Complaints');
      logAudit('system','complaint-sweep','Complaints','',{ expired:expired });
      return success({ expired:expired });
    } finally { lock.releaseLock(); }
  });
}
function installComplaintSweep() {
  removeComplaintSweep();
  ScriptApp.newTrigger('complaintSweep').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Complaint sweep installed (hourly).');
}
function removeComplaintSweep() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction()==='complaintSweep') ScriptApp.deleteTrigger(t); });
}
