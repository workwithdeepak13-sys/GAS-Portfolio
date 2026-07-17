/**
 * ============================================================================
 *  ISP Billing & Customer Management ERP — Backend (isp.gs / Code.gs)
 * ============================================================================
 *
 *  DEPLOYMENT
 *  ----------
 *  1. Open Google Sheets → Extensions → Apps Script.
 *  2. Paste this file as `Code.gs` and paste `isp.html` as an HTML file
 *     named exactly `Index`.
 *  3. Reload the sheet; the "⚙️ ISP ERP" menu appears. Click
 *     "Initialize System" once (safe to run again — idempotent).
 *  4. Deploy → New deployment → Web app → Execute as: Me,
 *     Access: Anyone with the link.
 *  5. Optional: "Install Billing Trigger" wires a daily sweep if needed.
 *
 *  SHEET STRUCTURE (created by initializeSheets)
 *  ---------------------------------------------
 *   Users         id | email | name | role | phone | active | otp | otpExpiry | createdAt
 *   Customers     id | name | phone | email | address | area | packageId | status | joinDate | createdAt
 *   Packages      id | name | speed | price | billingCycle | active | createdAt
 *   Areas         id | name | collectorId | createdAt
 *   Invoices      id | invoiceNumber | customerId | amount | dueDate | status | month | year | createdAt
 *   Payments      id | invoiceId | customerId | amount | method | collector | bankAccount | date | createdAt
 *   BankAccounts  id | name | bankName | accountNumber | createdAt
 *   Leads         id | name | phone | email | source | medium | status | createdAt
 *   Complaints    id | customerId | issueType | description | status | assignedTo | dueDate | createdAt
 *   Collections   id | date | collector | openingBalance | cashCollected | onlineCollected | expenses | bankDeposit | createdAt
 *   Expenses      id | category | payee | amount | date | description | createdAt
 *   Vendors       id | name | phone | email | gstin | createdAt
 *   Settings      key | value
 *   ActivityLog   id | ts | actor | action | entity | entityId | details
 *   ErrorLog      id | ts | fn | message | stack
 *
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const ISP_CFG = {
  CACHE_TTL_SEC: 60,
  LOCK_TIMEOUT_MS: 15000,
  MENU_NAME: '⚙️ ISP ERP',
  APP_NAME: 'ISP ERP'
};

const ISP_SHEETS = {
  Users: { headers:['id','email','name','role','phone','active','otp','otpExpiry','createdAt'], widths:[60,220,180,140,140,80,120,180,180] },
  Customers: { headers:['id','name','phone','email','address','area','packageId','status','joinDate','createdAt'], widths:[60,180,140,220,260,140,120,120,140,180] },
  Packages: { headers:['id','name','speed','price','billingCycle','active','createdAt'], widths:[60,220,140,120,140,80,180] },
  Areas: { headers:['id','name','collectorId','createdAt'], widths:[60,200,160,180] },
  Invoices: { headers:['id','invoiceNumber','customerId','amount','dueDate','status','month','year','createdAt'], widths:[60,180,80,120,120,120,90,90,180] },
  Payments: { headers:['id','invoiceId','customerId','amount','method','collector','bankAccount','date','createdAt'], widths:[60,80,80,120,110,160,160,120,180] },
  BankAccounts: { headers:['id','name','bankName','accountNumber','createdAt'], widths:[60,200,220,220,180] },
  Leads: { headers:['id','name','phone','email','source','medium','status','createdAt'], widths:[60,180,140,220,160,160,120,180] },
  Complaints: { headers:['id','customerId','issueType','description','status','assignedTo','dueDate','createdAt'], widths:[60,80,160,320,120,160,120,180] },
  Collections: { headers:['id','date','collector','openingBalance','cashCollected','onlineCollected','expenses','bankDeposit','createdAt'], widths:[60,120,180,140,140,140,120,140,180] },
  Expenses: { headers:['id','category','payee','amount','date','description','createdAt'], widths:[60,160,180,120,120,320,180] },
  Vendors: { headers:['id','name','phone','email','gstin','createdAt'], widths:[60,220,140,220,160,180] },
  Settings: { headers:['key','value'], widths:[260,520] },
  ActivityLog: { headers:['id','ts','actor','action','entity','entityId','details'], widths:[60,180,220,140,120,80,420] },
  ErrorLog: { headers:['id','ts','fn','message','stack'], widths:[60,180,220,420,520] }
};

const ISP_ROLES = ['Super Admin','Admin','Accountant','Collector'];
const ISP_PERMS = {
  'Super Admin': { dashboard:['view'], customers:['view','add','edit','delete'], billing:['view','run'], payments:['view','add','edit'], leads:['view','add','edit','delete'], complaints:['view','add','edit','delete'], collections:['view','add'], finance:['view','add','edit'], settings:['view','edit'], reports:['view','export'] },
  Admin: { dashboard:['view'], customers:['view','add','edit','delete'], billing:['view','run'], payments:['view','add','edit'], leads:['view','add','edit','delete'], complaints:['view','add','edit','delete'], collections:['view','add'], finance:['view','add','edit'], settings:['view','edit'], reports:['view','export'] },
  Accountant: { dashboard:['view'], customers:['view'], billing:['view','run'], payments:['view','add'], leads:['view'], complaints:['view'], collections:['view'], finance:['view','add','edit'], settings:['view'], reports:['view','export'] },
  Collector: { dashboard:['view'], customers:['view'], billing:['view'], payments:['view'], leads:['view'], complaints:['view'], collections:['view','add'], finance:['view'], settings:['view'], reports:['view'] }
};
const INVOICE_STATUSES = ['Unpaid','Paid','Overdue'];
const INVOICE_MONTHS = ['01','02','03','04','05','06','07','08','09','10','11','12'];
const PAYMENT_METHODS = ['Cash','Online','Bank'];
const LEAD_STATUSES = ['New','Contacted','Won','Lost'];
const COMPLAINT_STATUSES = ['Open','In Progress','Resolved'];

// ---------------------------------------------------------------------------
// Web app entry + menu
// ---------------------------------------------------------------------------
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ISP ERP')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(ISP_CFG.MENU_NAME)
    .addItem('Initialize System', 'menuInitialize')
    .addItem('Reset Dummy Data',  'menuResetDummy')
    .addItem('Export Demo CSV', 'menuExportDemoCsv')
    .addToUi();
}
function menuInitialize() {
  const r = initializeSheets();
  SpreadsheetApp.getUi().alert('ISP ERP initialized.\n\nCreated: ' + r.data.sheetsCreated.join(', ') +
    '\nSkipped: ' + r.data.sheetsSkipped.join(', ') +
    '\nDummy rows added: ' + r.data.dummyRowsAdded);
}
function menuResetDummy() {
  const r = resetDummyData();
  SpreadsheetApp.getUi().alert('Dummy rows cleared. Removed: ' + r.data.removed);
}
function menuExportDemoCsv() {
  const folder = DriveApp.getRootFolder();
  const file = folder.createFile('isp-demo.csv', 'demo export placeholder');
  SpreadsheetApp.getUi().alert('Exported: ' + file.getUrl());
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
    Object.keys(ISP_SHEETS).forEach(function(name){
      let sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
        const spec = ISP_SHEETS[name];
        sh.getRange(1,1,1,spec.headers.length).setValues([spec.headers])
          .setFontWeight('bold').setBackground('#0F172A').setFontColor('#ffffff');
        sh.setFrozenRows(1);
        spec.widths.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
        created.push(name);
      } else {
        skipped.push(name);
      }
    });
    const s1 = ss.getSheetByName('Sheet1');
    if (s1 && ss.getSheets().length > 1) { try { ss.deleteSheet(s1); } catch(e){} }

    seedConfigDefaults();
    let dummyRowsAdded = 0;
    dummyRowsAdded += seedAll();
    return success({ sheetsCreated:created, sheetsSkipped:skipped, dummyRowsAdded:dummyRowsAdded });
  });
}

function seedConfigDefaults() {
  const sh = getSheet('Settings');
  const existing = getSheetAsObjects('Settings').reduce(function(a,r){ a[r.key]=true; return a; },{});
  const defaults = [
    ['app.name','ISP ERP'],
    ['company.name','NetFiber Telecom'],
    ['company.email','support@netfiber.example.com'],
    ['currency.symbol','₹'],
    ['currency.code','INR'],
    ['billing.day','1'],
    ['billing.dueDays','15'],
    ['whatsapp.enabled','true'],
    ['whatsapp.template','Dear {name}, your invoice #{invoice} for {amount} is due on {dueDate}. Please pay to avoid disconnection.'],
    ['otp.ttl','10']
  ];
  const rows = defaults.filter(function(kv){ return !existing[kv[0]]; });
  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,2).setValues(rows);
}

function seedAll() {
  if (getSheet('Users').getLastRow() > 1) return 0;
  const now = new Date();
  const iso = d => new Date(d).toISOString();
  const daysAgo = function(n, baseDate){ const base = baseDate ? new Date(baseDate) : new Date(now); base.setDate(base.getDate()-n); return iso(base); };
  const nowIso = iso(now);
  let count = 0;
  function bulk(sheet, rows){ if (!rows.length) return; getSheet(sheet).getRange(2,1,rows.length,rows[0].length).setValues(rows); count += rows.length; }

  bulk('Users', [
    [1,'super@isp.com','Admin Root','Super Admin','+91-98100-00001',true,'', '', nowIso],
    [2,'admin@isp.com','Ops Admin','Admin','+91-98100-00002',true,'', '', nowIso],
    [3,'acc@isp.com','Finance Lead','Accountant','+91-98100-00003',true,'', '', nowIso],
    [4,'col1@isp.com','Field Collector 1','Collector','+91-98100-00004',true,'', '', daysAgo(5)],
    [5,'col2@isp.com','Field Collector 2','Collector','+91-98100-00005',true,'', '', daysAgo(5)]
  ]);

  bulk('Areas', [
    [1,'Sector 12','4'],
    [2,'Sector 15','5'],
    [3,'DLF Phase 3','4']
  ]);

  bulk('Packages', [
    [1,'Basic 100','100 Mbps','699','Monthly',true,nowIso],
    [2,'Standard 300','300 Mbps','999','Monthly',true,nowIso],
    [3,'Premium 600','600 Mbps','1499','Monthly',true,nowIso],
    [4,'Business 1G','1 Gbps','3499','Monthly',true,nowIso]
  ]);

  bulk('Customers', [
    [1,'Rahul Verma','+91-98100-10001','rahul@example.com','242 DLF Phase 3','DLF Phase 3',2,'Active',daysAgo(90),daysAgo(90)],
    [2,'Sneha Kapoor','+91-98100-10002','sneha@example.com','A-12 Sector 15','Sector 15',1,'Active',daysAgo(60),daysAgo(60)],
    [3,'Amit Singh','+91-98100-10003','amit@example.com','B-77 Sector 12','Sector 12',3,'Active',daysAgo(20),daysAgo(20)],
    [4,'Priya Menon','+91-98100-10004','priya@example.com','18 Green Park','Sector 15',2,'Active',daysAgo(10),daysAgo(10)]
  ]);

  bulk('Invoices', [
    [1,'INV/2026/0001',1,699,daysAgo(5),'Unpaid','07','2026',daysAgo(20)],
    [2,'INV/2026/0002',2,699,daysAgo(8),'Paid','07','2026',daysAgo(20)],
    [3,'INV/2026/0003',3,1499,daysAgo(12),'Overdue','07','2026',daysAgo(20)]
  ]);

  bulk('Payments', [
    [1,2,2,699,'Cash','col1','',daysAgo(6),daysAgo(6)],
    [2,3,3,1499,'Online','col2','HDFC-001',daysAgo(13),daysAgo(13)]
  ]);

  bulk('BankAccounts', [
    [1,'Main Current','HDFC Bank','502000123456',''],
    [2,'Collections Wallet','ICICI Bank','502000987654','']
  ]);

  bulk('Vendors', [
    [1,'FiberLink Tech','+91-98100-20001','ops@fiberlink.example.com','27AABCA1234A1Z5',''],
    [2,'NetPro Supplies','+91-98100-20002','sales@netpro.example.com','27AABCB1234A1Z5','']
  ]);

  bulk('Leads', [
    [1,'Kiran Rao','+91-98100-30001','kiran@example.com','Referral','Website','New',daysAgo(2)],
    [2,'Meera Iyer','+91-98100-30002','meera@example.com','Cold Call','Phone','Contacted',daysAgo(4)],
    [3,'Rakesh Jain','+91-98100-30003','rakesh@example.com','Ads','Instagram','Won',daysAgo(10)]
  ]);

  bulk('Complaints', [
    [1,3,'Internet Slow','Speed drops after 8PM','In Progress','col1',daysAgo(1),daysAgo(1)],
    [2,1,'No Internet','Complete outage in sector','Open','col2',daysAgo(0),daysAgo(0)]
  ]);

  bulk('Collections', [
    [1,daysAgo(1),'col1',12000,18500,3200,450,'HDFC-001',daysAgo(1)],
    [2,daysAgo(2),'col2',9500,11200,1500,200,'ICICI-002',daysAgo(2)]
  ]);

  bulk('Expenses', [
    [1,'Fuel','Bharat Petroleum',1200,daysAgo(1,'2026-07-01'),'Collection run',''],
    [2,'Repair','Laptop Clinic',3400,daysAgo(2,'2026-07-02'),'ONT replacement','']
  ]);

  return count;
}

function resetDummyData() {
  return safeCall('resetDummyData', function(){
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const ss = SpreadsheetApp.getActive();
      let removed = 0;
      Object.keys(ISP_SHEETS).forEach(function(name){
        if (name === 'Settings' || name === 'ActivityLog' || name === 'ErrorLog') return;
        const sh = ss.getSheetByName(name);
        if (!sh) return;
        const last = sh.getLastRow();
        if (last > 1) { sh.deleteRows(2, last-1); removed += last-1; }
      });
      seedAll();
      return success({ removed:removed });
    } finally { try { lock.releaseLock(); } catch(e){} }
  });
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function getSheet(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}
function getSheetAsObjects(name) {
  const sh = getSheet(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  return values.map(function(row){ const o = {}; headers.forEach((h,i)=> o[h]=row[i]); return o; });
}
function nowIso() { return new Date().toISOString(); }
function formatCurrency(n) {
  const sym = (getSheetAsObjects('Settings').find(function(r){ return r.key === 'currency.symbol'; }) || {}).value || '₹';
  return sym + Number(n||0).toLocaleString('en-IN',{maximumFractionDigits:2});
}
function generateId(prefix, sheetName, colIndex) {
  const sh = getSheet(sheetName);
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, colIndex||1, last-1, 1).getValues() : [];
  const next = (rows.filter(String).length) + 1;
  return prefix + '-' + String(next).padStart(6,'0');
}
function hasPermission(role, module, action) {
  const perm = ISP_PERMS[role];
  if (!perm || !perm[module]) return false;
  const a = perm[module];
  return a.indexOf('view') !== -1 || a.indexOf(action) !== -1 || a.indexOf('add') !== -1 || a.indexOf('edit') !== -1 || a.indexOf('delete') !== -1 || a.indexOf('run') !== -1 || a.indexOf('export') !== -1;
}

// ---------------------------------------------------------------------------
// Logging
// ---------------------------------------------------------------------------
function logError(fnName, e) {
  try {
    const sh = getSheet('ErrorLog');
    sh.appendRow([Utilities.getUuid(), nowIso(), fnName, e.message || String(e), (e.stack||'').slice(0,4000)]);
  } catch(err){}
}
function logActivity(user, action, entity, entityId, details) {
  try { getSheet('ActivityLog').appendRow([Utilities.getUuid(), nowIso(), user||'system', action, entity, entityId, details||'']); } catch(e){}
}

// ---------------------------------------------------------------------------
// Option request if external calls hit CORs
// ---------------------------------------------------------------------------
function doOptions(e){ return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT); }

// ---------------------------------------------------------------------------
// Auth bootstrap
// ---------------------------------------------------------------------------
function bootstrap() {
  return safeCall('bootstrap', function(){
    const me = getCurrentUser();
    if (!me) return success({ me:null, config:{}, modules:{}, packages:[], areas:[], bankAccounts:[], vendors:[], customers:[] });
    const config = getConfig();
    const packages = getSheetAsObjects('Packages').filter(function(x){ return x.active !== false; });
    const areas = getSheetAsObjects('Areas');
    const bankAccounts = getSheetAsObjects('BankAccounts');
    const vendors = getSheetAsObjects('Vendors');
    const modules = {};
    ['dashboard','customers','billing','payments','leads','complaints','collections','finance','settings','reports'].forEach(function(m){
      modules[m] = { view: hasPermission(me.role,m,'view'), add: hasPermission(me.role,m,'add'), edit: hasPermission(me.role,m,'edit'), delete: hasPermission(me.role,m,'delete'), run: hasPermission(me.role,m,'run'), export: hasPermission(me.role,m,'export') };
    });
    return success({ me:me, config:config, modules:modules, packages:packages, areas:areas, bankAccounts:bankAccounts, vendors:vendors });
  });
}
function getCurrentUser() {
  try { const r = PropertiesService.getUserProperties().getProperty('_ispUser'); return r ? JSON.parse(r) : null; } catch(e){ return null; }
}
function signOut() { PropertiesService.getUserProperties().deleteProperty('_ispUser'); return success({}); }
function getConfig() {
  const rows = getSheetAsObjects('Settings');
  const out = {};
  rows.forEach(function(r){ out[r.key] = r.value; }); return out;
}

// ---------------------------------------------------------------------------
// Auth / OTP helpers (client sends OTP and verifies server-side)
// ---------------------------------------------------------------------------
function sendOtp(email) {
  return safeCall('sendOtp', function(){
    const users = getSheetAsObjects('Users');
    const u = users.find(function(x){
      const isActive = x.active === true || String(x.active).toLowerCase() === 'true' || x.active === '' || x.active == null;
      return String(x.email||'').toLowerCase() === String(email).toLowerCase() && isActive;
    });
    if (!u) return fail('User not found');
    const code = String(Math.floor(1000 + Math.random()*9000));
    const ttl = (getConfig()['otp.ttl'] || '10');
    const expiry = new Date(new Date().getTime() + Number(ttl)*60000).toISOString();
    const sh = getSheet('Users');
    const row = users.findIndex(function(x){ return x.id === u.id; }) + 2;
    sh.getRange(row, 7, 1, 2).setValues([[code, expiry]]);
    try {
      MailApp.sendEmail({ to: email, subject: 'Your ISP ERP OTP', htmlBody: '<b>OTP:</b> ' + code + '<br/>Valid for ' + ttl + ' minutes.' });
    } catch (e) { return fail('Email failed: ' + e.message); }
    logActivity(u.email, 'otp-sent', 'Users', u.id, 'OTP sent to ' + email);
    return success({ message:'OTP sent', masked:email.replace(/(.{2})(.*)(@.*)/, '$1***$3') });
  });
}
function verifyOtp(email, passwordOrOtp) {
  return safeCall('verifyOtp', function(){
    const users = getSheetAsObjects('Users');
    const u = users.find(function(x){ return String(x.email||'').toLowerCase() === String(email).toLowerCase(); });
    if (!u) return fail('Invalid credentials');
    const usersSh = getSheet('Users');
    const row = users.findIndex(function(x){ return x.id === u.id; }) + 2;
    const otpRow = usersSh.getRange(row, 7, 1, 2).getValues()[0];
    const otp = String(otpRow[0]||'');
    const expiry = otpRow[1];
    let valid = false;
    if (otp && expiry && new Date().getTime() < new Date(expiry).getTime()) {
      if (String(passwordOrOtp) === String(otp)) valid = true;
    }
    if (!valid) return fail('Invalid or expired OTP');
    const user = { id:u.id, email:u.email, name:u.name, role:u.role, phone:u.phone };
    PropertiesService.getUserProperties().setProperty('_ispUser', JSON.stringify(user));
    usersSh.getRange(row, 7, 1, 2).clearContent();
    logActivity(u.email, 'login', 'Users', u.id, 'Logged in');
    return success({ user:user });
  });
}

// ---------------------------------------------------------------------------
// Customers
// ---------------------------------------------------------------------------
function getCustomers() { return success(getSheetAsObjects('Customers')); }
function addCustomer(data) {
  return safeCall('addCustomer', function(){
    const u = getCurrentUser(); if (!u) return fail('Unauthorized');
    if (!hasPermission(u.role,'customers','add')) return fail('Forbidden');
    const id = generateId('CUS','Customers',1);
    data.id = id; data.createdAt = nowIso(); data.status = data.status || 'Active';
    getSheet('Customers').appendRow([data.id,data.name,data.phone,data.email,data.address,data.area,data.packageId,data.status,data.joinDate||nowIso().split('T')[0],data.createdAt]);
    logActivity(u.email,'create','Customers',id,'Created customer'); return success({ id:id });
  });
}
function updateCustomer(id, data) {
  return safeCall('updateCustomer', function(){
    const u = getCurrentUser(); if (!u) return fail('Unauthorized'); if (!hasPermission(u.role,'customers','edit')) return fail('Forbidden');
    const sh = getSheet('Customers'), rows = sh.getDataRange().getValues(), headers = rows[0];
    const idx = headers.indexOf('id'); if (idx < 0) return fail('Bad schema');
    for (let i=1;i<rows.length;i++) if (String(rows[i][idx]) === String(id)) {
      Object.keys(data).forEach(function(k){ const c=headers.indexOf(k); if (c>=0) sh.getRange(i+1,c+1).setValue(data[k]); });
      logActivity(u.email,'update','Customers',id,'Updated customer'); return success({});
    } return fail('Not found');
  });
}
function deleteCustomer(id) {
  return safeCall('deleteCustomer', function(){
    const u = getCurrentUser(); if (!u) return fail('Unauthorized'); if (!hasPermission(u.role,'customers','delete')) return fail('Forbidden');
    const sh = getSheet('Customers'), rows = sh.getDataRange().getValues(), headers = rows[0], idx = headers.indexOf('id');
    for (let i=1;i<rows.length;i++) if (String(rows[i][idx]) === String(id)) { sh.deleteRow(i+1); logActivity(u.email,'delete','Customers',id,'Deleted'); return success({}); }
    return fail('Not found');
  });
}

// ---------------------------------------------------------------------------
// Packages / Areas / Bank / Vendor
// ---------------------------------------------------------------------------
function getPackages() { return success(getSheetAsObjects('Packages').filter(function(r){ return r.active !== false; })); }
function addPackage(data) {
  return safeCall('addPackage', function(){
    const u=getCurrentUser(); if(!u||!hasPermission(u.role,'settings','add')) return fail('Forbidden');
    data.id = generateId('PKG','Packages',1); data.active = true; data.createdAt = nowIso();
    getSheet('Packages').appendRow([data.id,data.name,data.speed,data.price,data.billingCycle,data.active,data.createdAt]);
    logActivity(u.email,'create','Packages',data.id,'Created package'); return success({ id:data.id });
  });
}
function getAreas() { return success(getSheetAsObjects('Areas')); }
function addArea(data) {
  return safeCall('addArea', function(){
    const u=getCurrentUser(); if(!u||!hasPermission(u.role,'settings','add')) return fail('Forbidden');
    data.id = generateId('AREA','Areas',1); data.createdAt = nowIso();
    getSheet('Areas').appendRow([data.id,data.name,data.collectorId||'',data.createdAt]);
    logActivity(u.email,'create','Areas',data.id,'Created area'); return success({ id:data.id });
  });
}
function getBankAccounts() { return success(getSheetAsObjects('BankAccounts')); }
function getVendors() { return success(getSheetAsObjects('Vendors')); }
function addVendor(data) {
  return safeCall('addVendor', function(){
    const u=getCurrentUser(); if(!u||!hasPermission(u.role,'finance','add')) return fail('Forbidden');
    data.id = generateId('VEN','Vendors',1); data.createdAt = nowIso();
    getSheet('Vendors').appendRow([data.id,data.name,data.phone,data.email,data.gstin||'',data.createdAt]);
    logActivity(u.email,'create','Vendors',data.id,'Created vendor'); return success({ id:data.id });
  });
}

// ---------------------------------------------------------------------------
// Invoicing & Billing
// ---------------------------------------------------------------------------
function getInvoices() { return success(getSheetAsObjects('Invoices')); }
function runBilling(month, year) {
  return safeCall('runBilling', function(){
    const u=getCurrentUser(); if(!u||!hasPermission(u.role,'billing','run')) return fail('Forbidden');
    const existing = getSheetAsObjects('Invoices').filter(function(x){ return String(x.month)===String(month) && String(x.year)===String(year); });
    if (existing.length) return fail('Billing already ran for this month/year');
    const customers = getSheetAsObjects('Customers').filter(function(c){ return c.status === 'Active'; });
    if (!customers.length) return fail('No active customers');
    const rows = customers.map(function(c){
      const pkg = (getSheetAsObjects('Packages').find(function(p){ return String(p.id) === String(c.packageId); }) || {});
      const amount = Number(pkg.price || 0);
      const due = new Date(); due.setMonth(due.getMonth()+1); due.setDate(15);
      return [generateId('INV','Invoices',2), 'INV/'+year+'/'+String(Math.floor(Math.random()*89999)+10000), c.id, amount, Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd'), 'Unpaid', month, year, nowIso()];
    });
    getSheet('Invoices').getRange(getSheet('Invoices').getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
    logActivity(u.email,'billing-run','Invoices','MN:'+month+'/'+year,'Generated invoices'); return success({ invoicesGenerated: rows.length });
  });
}
function payInvoice(invoiceId, data) {
  return safeCall('payInvoice', function(){
    const u=getCurrentUser(); if(!u||!hasPermission(u.role,'payments','add')) return fail('Forbidden');
    const sh = getSheet('Invoices'), rows = sh.getDataRange().getValues(), headers = rows[0];
    const idx = headers.indexOf('id');
    for (let i=1;i<rows.length;i++) if (String(rows[i][idx]) === String(invoiceId)) {
      sh.getRange(i+1, headers.indexOf('status')+1).setValue('Paid'); sh.getRange(i+1, headers.indexOf('createdAt')+1).setValue(nowIso());
      getSheet('Payments').appendRow([Utilities.getUuid(), invoiceId, rows[i][headers.indexOf('customerId')], Number(data.amount||rows[i][headers.indexOf('amount')]), data.method||'Cash', data.collector||'', data.bankAccount||'', data.date||nowIso().split('T')[0], nowIso()]);
      logActivity(u.email,'payment','Payments',invoiceId,'Recorded payment for '+invoiceId); return success({});
    } return fail('Invoice not found');
  });
}

// ---------------------------------------------------------------------------
// Leads
// ---------------------------------------------------------------------------
function getLeads() { return success(getSheetAsObjects('Leads')); }
function addLead(data) {
  return safeCall('addLead', function(){
    const u=getCurrentUser(); if (!u) return fail('Unauthorized'); if (!hasPermission(u.role,'leads','add')) return fail('Forbidden');
    data.id = generateId('LED','Leads',1); data.status = data.status || 'New'; data.createdAt = nowIso();
    getSheet('Leads').appendRow([data.id,data.name,data.phone,data.email,data.source,data.medium,data.status,data.createdAt]);
    logActivity(u.email,'create','Leads',data.id,'Created lead'); return success({ id:data.id });
  });
}
function markLeadWon(id, packageId, area) {
  return safeCall('markLeadWon', function(){
    const u=getCurrentUser(); if (!u) return fail('Unauthorized'); if (!hasPermission(u.role,'leads','edit')) return fail('Forbidden');
    const leads = getSheetAsObjects('Leads');
    const lead = leads.find(function(x){ return x.id === id; });
    if (!lead) return fail('Lead not found');
    const cid = addCustomer({ name:lead.name, phone:lead.phone, email:lead.email, address:'', area:area||'', packageId:packageId||'' });
    if (!cid.success) return fail('Customer conversion failed: ' + cid.error);
    const sh=getSheet('Leads'), rows=sh.getDataRange().getValues(), headers=rows[0], idx=headers.indexOf('id');
    for (let i=1;i<rows.length;i++) if (String(rows[i][idx])===String(id)) { sh.getRange(i+1,headers.indexOf('status')+1).setValue('Won'); break; }
    logActivity(u.email,'convert','Leads',id,'Converted to customer '+cid.data.id); return success({ customerId:cid.data.id });
  });
}

// ---------------------------------------------------------------------------
// Complaints
// ---------------------------------------------------------------------------
function getComplaints() { return success(getSheetAsObjects('Complaints')); }
function addComplaint(data) {
  return safeCall('addComplaint', function(){
    const u=getCurrentUser(); if (!u) return fail('Unauthorized'); if (!hasPermission(u.role,'complaints','add')) return fail('Forbidden');
    data.id = generateId('CMP','Complaints',1); data.status = data.status || 'Open'; data.createdAt = nowIso();
    getSheet('Complaints').appendRow([data.id,data.customerId,data.issueType,data.description,data.status,data.assignedTo||'',data.dueDate||'',data.createdAt]);
    logActivity(u.email,'create','Complaints',data.id,'Created complaint'); return success({ id:data.id });
  });
}
function updateComplaintStatus(id, status) {
  return safeCall('updateComplaintStatus', function(){
    const u=getCurrentUser(); if (!u) return fail('Unauthorized'); if (!hasPermission(u.role,'complaints','edit')) return fail('Forbidden');
    const sh = getSheet('Complaints'), rows = sh.getDataRange().getValues(), headers = rows[0], idx = headers.indexOf('id');
    if (String(status).toLowerCase() === 'resolved') { if (!hasPermission(u.role,'complaints','edit')) return fail('Forbidden'); }
    for (let i=1;i<rows.length;i++) if (String(rows[i][idx]) === String(id)) { sh.getRange(i+1,headers.indexOf('status')+1).setValue(status); logActivity(u.email,'update','Complaints',id,'Status -> '+status); return success({}); } return fail('Not found');
  });
}

// ---------------------------------------------------------------------------
// Collections & Finance
// ---------------------------------------------------------------------------
function getCollectionRuns() { return success(getSheetAsObjects('Collections')); }
function addCollectionRun(data) {
  return safeCall('addCollectionRun', function(){
    const u=getCurrentUser(); if (!u||!hasPermission(u.role,'collections','add')) return fail('Forbidden');
    data.id = generateId('COL','Collections',1); data.createdAt = nowIso();
    getSheet('Collections').appendRow([data.id,data.date,data.collector,data.openingBalance,data.cashCollected,data.onlineCollected,data.expenses,data.bankDeposit,data.createdAt]);
    logActivity(u.email,'create','Collections',data.id,'Created collection run'); return success({ id:data.id });
  });
}
function getExpenses() { return success(getSheetAsObjects('Expenses')); }
function addExpense(data) {
  return safeCall('addExpense', function(){
    const u=getCurrentUser(); if (!u||!hasPermission(u.role,'finance','add')) return fail('Forbidden');
    data.id = generateId('EXP','Expenses',1); data.createdAt = nowIso();
    getSheet('Expenses').appendRow([data.id,data.category,data.payee,data.amount,data.date,data.description,data.createdAt]);
    logActivity(u.email,'create','Expenses',data.id,'Created expense'); return success({ id:data.id });
  });
}
function addVendorExpensePair(data) {
  const r = addVendor(data.vendor || { name:data.payee, phone:data.phone||'', email:data.email||'', gstin:data.gstin||'' });
  if (r.success && r.data && r.data.id) {
    addExpense({ category:data.category, payee:data.payee, amount:data.amount, date:data.date, description:data.description||'', qty:data.qty||'' });
    return success({ vendorId:r.data.id });
  }
  return fail('Vendor create failed');
}

// ---------------------------------------------------------------------------
// Dashboard summary
// ---------------------------------------------------------------------------
function getDashboard() {
  return safeCall('getDashboard', function(){
    const customers = getSheetAsObjects('Customers');
    const invoices = getSheetAsObjects('Invoices');
    const payments = getSheetAsObjects('Payments');
    const complaints = getSheetAsObjects('Complaints');
    const leads = getSheetAsObjects('Leads');
    const expenses = getSheetAsObjects('Expenses');
    const paid = invoices.filter(function(x){ return x.status === 'Paid'; });
    const overdue = invoices.filter(function(x){ return x.status === 'Overdue'; });
    const revenue = paid.reduce(function(s,x){ return s + Number(x.amount||0); },0);
    const expenseTotal = expenses.reduce(function(s,x){ return s + Number(x.amount||0); },0);
    const thisMonth = invoices.filter(function(x){ return String(x.month) === String(new Date().getMonth()+1) });
    return success({
      customers:customers.length,
      monthlyRevenue: revenue,
      unpaidInvoices: invoices.filter(function(x){ return x.status === 'Unpaid'; }).length,
      overdueAmount: overdue.reduce(function(s,x){ return s + Number(x.amount||0); },0),
      newLeads: leads.filter(function(x){ return x.status === 'New'; }).length,
      openComplaints: complaints.filter(function(x){ return x.status === 'Open' || x.status === 'In Progress'; }).length,
      expenseSummary: expenseTotal,
      recentInvoices: invoices.slice(-8).reverse()
    });
  });
}

// ---------------------------------------------------------------------------
// CSV exports
// ---------------------------------------------------------------------------
function exportCsv(entity) {
  return safeCall('exportCsv', function(){
    let rows = [];
    if (entity === 'invoices') rows = getSheetAsObjects('Invoices');
    else if (entity === 'payments') rows = getSheetAsObjects('Payments');
    else if (entity === 'customers') rows = getSheetAsObjects('Customers');
    else if (entity === 'complaints') rows = getSheetAsObjects('Complaints');
    else if (entity === 'recovery') rows = getSheetAsObjects('Invoices').filter(function(x){ return x.status === 'Overdue'; });
    else if (entity === 'collections') rows = getSheetAsObjects('Collections');
    else if (entity === 'expenses') rows = getSheetAsObjects('Expenses');
    else return fail('Unknown entity');
    if (!rows.length) return fail('No data');
    const headers = Object.keys(rows[0]);
    const lines = [headers.join(',')];
    rows.forEach(function(r){ lines.push(headers.map(function(h){ return String(r[h]||'').replace(/,/g,' '); }).join(',')); });
    const csv = lines.join('\n');
    return success({ csv:csv });
  });
}

// ---------------------------------------------------------------------------
// Generic paginated data
// ---------------------------------------------------------------------------
function getPagedData(entity, page, pageSize, filters) {
  return safeCall('getPagedData', function(){
    let rows = [];
    if (entity === 'Customers') rows = getSheetAsObjects('Customers');
    else if (entity === 'Invoices') rows = getSheetAsObjects('Invoices');
    else if (entity === 'Payments') rows = getSheetAsObjects('Payments');
    else if (entity === 'Leads') rows = getSheetAsObjects('Leads');
    else if (entity === 'Complaints') rows = getSheetAsObjects('Complaints');
    else if (entity === 'Collections') rows = getSheetAsObjects('Collections');
    else if (entity === 'Expenses') rows = getSheetAsObjects('Expenses');
    else return fail('Unsupported entity');
    if (filters) {
      Object.keys(filters).forEach(function(k){
        const v = String(filters[k]||'').toLowerCase();
        if (!v) return;
        rows = rows.filter(function(r){ return String(r[k]||'').toLowerCase().indexOf(v) !== -1; });
      });
    }
    const total = rows.length;
    const start = (page-1)*pageSize;
    rows = rows.slice(start, start+pageSize);
    return success({ rows:rows, total:total, page:page, pageSize:pageSize });
  });
}
