/**
 * ============================================================================
 *  GateDesk — Google Apps Script Web App (single-spreadsheet backend)
 * ============================================================================
 *  Digital gate register / visitor management. A visitor scans a QR at the
 *  gate → fills a public self-registration form (with photo capture) → the
 *  host is notified and approves/rejects → guard/reception sees a live
 *  "who's inside" board → admin gets history, audit search & CSV export.
 *
 *  DEPLOYMENT
 *  ----------
 *  1. Open Google Sheets → Extensions → Apps Script.
 *  2. Paste this file as Code.gs. Add HTML files:
 *       • gatedesk.html          → HTML file named exactly "Index"
 *       • gatedesk-checkin.html  → HTML file named exactly "CheckinForm"
 *  3. Reload the sheet; the "⚙️ GateDesk" menu appears. Click
 *     "Initialize System" once (safe to run again — idempotent).
 *  4. Deploy → New deployment → Web app → Execute as: Me,
 *     Access: Anyone with the link (so visitors can open the check-in form).
 *  5. Optional: "Install Overstay Trigger" from the same menu wires a daily
 *     sweep that flags anyone still inside past the configured cutoff.
 *
 *  The public self-registration form is reachable at:
 *     <web-app-url>?page=checkin
 *
 *  SHEET STRUCTURE (created by initializeSheets)
 *  ---------------------------------------------
 *   Users       id | email | name | role | active | createdAt
 *               (role = Guard | Host | Admin)
 *   Visitors    id | phone | name | company | photoUrl | idProofLast4 | createdAt
 *   VisitLogs   id | visitorId | hostId | purpose | checkInTime | checkOutTime |
 *               durationMins | status | approvalMethod | source | createdAt | deleted
 *               status = Pending Approval | Approved | Rejected | Inside | Checked Out
 *               source = QR Self-Registration | Guard Manual Entry
 *   Hosts       id | name | phone | department | active
 *   Config      key | value
 *   AuditLog    id | ts | actor | action | entity | entityId | details
 *   ErrorLog    id | ts | fn | message | stack
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const GD_SHEETS = {
  Users: {
    headers: ['id','email','name','role','active','createdAt'],
    widths:  [60,220,180,120,80,180]
  },
  Visitors: {
    headers: ['id','phone','name','company','photoUrl','idProofLast4','createdAt'],
    widths:  [60,140,180,200,260,120,180]
  },
  VisitLogs: {
    headers: ['id','visitorId','hostId','purpose','checkInTime','checkOutTime','durationMins','status','approvalMethod','source','createdAt','deleted'],
    widths:  [60,90,80,240,170,170,110,140,140,170,180,80]
  },
  Hosts: {
    headers: ['id','name','phone','department','active'],
    widths:  [60,190,150,180,80]
  },
  Config: {
    headers: ['key','value'],
    widths:  [280,360]
  },
  AuditLog: {
    headers: ['id','ts','actor','action','entity','entityId','details'],
    widths:  [60,180,220,150,120,90,420]
  },
  ErrorLog: {
    headers: ['id','ts','fn','message','stack'],
    widths:  [60,180,220,420,520]
  }
};

const GD_STATUS  = ['Pending Approval','Approved','Rejected','Inside','Checked Out'];
const GD_SOURCES = ['QR Self-Registration','Guard Manual Entry'];
const GD_ROLES   = ['Guard','Host','Admin'];
const GD_CACHE_TTL = 45; // seconds

// ---------------------------------------------------------------------------
// Web app entry + menu
// ---------------------------------------------------------------------------
function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  if (page === 'checkin') {
    return HtmlService.createTemplateFromFile('CheckinForm')
      .evaluate()
      .setTitle('GateDesk · Visitor Check-In')
      .addMetaTag('viewport','width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('GateDesk')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function getWebAppUrl() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ GateDesk')
    .addItem('Initialize System', 'menuInitialize')
    .addItem('Reset Dummy Data',  'menuResetDummy')
    .addSeparator()
    .addItem('Install Overstay Trigger', 'installOverstayTrigger')
    .addItem('Remove Overstay Trigger',  'removeOverstayTrigger')
    .addToUi();
}
function menuInitialize() {
  const r = initializeSheets();
  SpreadsheetApp.getUi().alert('GateDesk initialized.\n\nCreated: ' + r.data.sheetsCreated.join(', ') +
    '\nSkipped: ' + r.data.sheetsSkipped.join(', ') +
    '\nDummy rows added: ' + r.data.dummyRowsAdded);
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
    Object.keys(GD_SHEETS).forEach(function(name){
      let sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
        const spec = GD_SHEETS[name];
        sh.getRange(1,1,1,spec.headers.length).setValues([spec.headers])
          .setFontWeight('bold').setBackground('#0B1220').setFontColor('#ffffff');
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
    dummyRowsAdded += seedUsers();
    dummyRowsAdded += seedHosts();
    dummyRowsAdded += seedVisitorsAndLogs();

    invalidateEntityCache();
    return success({ sheetsCreated:created, sheetsSkipped:skipped, dummyRowsAdded:dummyRowsAdded });
  });
}

function seedConfigDefaults() {
  const sh = getSheet('Config');
  const existing = getSheetAsObjects('Config').reduce(function(a,r){ a[r.key]=true; return a; },{});
  const defaults = [
    ['app.name','GateDesk'],
    ['site.name','Flowmative Industries — Main Gate'],
    ['notify.channel','A'],                 // A = Email (MailApp), B = WhatsApp BSP scaffold
    ['notify.enabled','true'],
    ['notify.bspEndpoint',''],              // Option B REST endpoint (Script Property holds token)
    ['workday.start','09:00'],
    ['workday.end','18:00'],
    ['overstay.cutoff','20:00'],            // flag visitors still Inside past this time
    ['photo.folderName','GateDesk Visitor Photos'],
    ['checkin.url','']                      // optional explicit public URL for QR (else auto)
  ];
  const rows = defaults.filter(function(kv){ return !existing[kv[0]]; });
  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,2).setValues(rows);
}

function seedUsers() {
  const sh = getSheet('Users');
  if (sh.getLastRow() > 1) return 0;
  const now = new Date().toISOString();
  const rows = [
    [1,'guard@example.com','Ramesh Yadav','Guard',true,now],
    [2,'reception@example.com','Anita Desai','Guard',true,now],
    [3,'host1@example.com','Vikram Nair','Host',true,now],
    [4,'host2@example.com','Priya Menon','Host',true,now],
    [5,'admin@example.com','Root Admin','Admin',true,now]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

function seedHosts() {
  const sh = getSheet('Hosts');
  if (sh.getLastRow() > 1) return 0;
  const rows = [
    [1,'Vikram Nair','+91-98100-00101','Engineering',true],
    [2,'Priya Menon','+91-98100-00102','Engineering',true],
    [3,'Rahul Sharma','+91-98100-00103','Operations',true],
    [4,'Sunita Rao','+91-98100-00104','Human Resources',true],
    [5,'Deepak Shah','+91-98100-00105','Finance',true],
    [6,'Meera Iyer','+91-98100-00106','Purchase',true]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

function seedVisitorsAndLogs() {
  const vsh = getSheet('Visitors');
  const lsh = getSheet('VisitLogs');
  if (vsh.getLastRow() > 1 || lsh.getLastRow() > 1) return 0;
  const now = new Date();
  function ago(mins){ const d = new Date(now); d.setMinutes(d.getMinutes()-mins); return d.toISOString(); }

  // Master visitor directory (note: visitor id=5 is a returning visitor — same phone appears twice in logs)
  const visitors = [
    [1,'9820011001','Sanjay Gupta','Infosys Ltd','', '4521', ago(600)],
    [2,'9820011002','Neha Kapoor','TechnoWorld Distributors','', '7788', ago(400)],
    [3,'9820011003','Imran Sheikh','BlueDart Express','', '1122', ago(300)],
    [4,'9820011004','Divya Reddy','SecureIT Audit Partners','', '3344', ago(200)],
    [5,'9820011005','Karan Malhotra','Individual','', '9955', ago(5000)]
  ];
  vsh.getRange(2,1,visitors.length,visitors[0].length).setValues(visitors);

  // VisitLogs — cover every status + a returning visitor (visitorId 5 twice)
  const logs = [
    // id, visitorId, hostId, purpose, checkIn, checkOut, durationMins, status, approvalMethod, source, createdAt, deleted
    [1, 1, 3, 'Business meeting',          ago(120), '',       '',  'Inside',           'In-App', 'QR Self-Registration', ago(130), false],
    [2, 2, 6, 'Product demo',              ago(90),  ago(15),  75,  'Checked Out',      'In-App', 'QR Self-Registration', ago(95),  false],
    [3, 3, 3, 'Courier delivery',          '',       '',       '',  'Pending Approval', '',       'QR Self-Registration', ago(20),  false],
    [4, 4, 5, 'Compliance audit',          ago(60),  '',       '',  'Inside',           'In-App', 'Guard Manual Entry',   ago(65),  false],
    [5, 5, 4, 'Job interview',             '',       '',       '',  'Rejected',         'In-App', 'QR Self-Registration', ago(45),  false],
    [6, 5, 3, 'Follow-up interview round', ago(4900),ago(4770),130, 'Checked Out',      'In-App', 'QR Self-Registration', ago(4905),false]
  ];
  lsh.getRange(2,1,logs.length,logs[0].length).setValues(logs);
  return visitors.length + logs.length;
}

// ---------------------------------------------------------------------------
// resetDummyData — clears seeded rows (ids <= 1000), keeps real rows (id > 1000)
// ---------------------------------------------------------------------------
function resetDummyData() {
  return safeCall('resetDummyData', function(){
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      let removed = 0;
      ['Visitors','VisitLogs','Hosts','Users'].forEach(function(name){
        const sh = getSheet(name);
        const last = sh.getLastRow();
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
    headers.forEach(function(h,i){ let v = r[i]; if (v instanceof Date) v = v.toISOString(); o[h] = v; });
    return o;
  });
}
function appendRowFromObject(sheetName, obj) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  if (headers[0] === 'id' && (obj.id === undefined || obj.id === null || obj.id === '')) {
    obj.id = nextId(sheetName);
  }
  const row = headers.map(function(h){ const v = obj[h]; return (v === undefined || v === null) ? '' : v; });
  sh.appendRow(row);
  return obj;
}
function updateRowById(sheetName, idColumn, id, updated) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idIdx = headers.indexOf(idColumn);
  if (idIdx < 0) throw new Error('Column not found: ' + idColumn);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2,1,last-1,headers.length).getValues();
  for (let r=0; r<data.length; r++) {
    if (String(data[r][idIdx]) === String(id)) {
      headers.forEach(function(h,i){ if (updated.hasOwnProperty(h)) data[r][i] = updated[h]; });
      sh.getRange(r+2,1,1,headers.length).setValues([data[r]]);
      const obj = {}; headers.forEach(function(h,i){ let v=data[r][i]; if (v instanceof Date) v=v.toISOString(); obj[h]=v; });
      return obj;
    }
  }
  return null;
}
function nextId(sheetName) {
  const sh = getSheet(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const ids = sh.getRange(2,1,last-1,1).getValues().map(function(r){ return Number(r[0])||0; });
  return Math.max.apply(null, ids) + 1;
}

// ---------------------------------------------------------------------------
// Cache helpers
// ---------------------------------------------------------------------------
function invalidateEntityCache(entity) {
  try { CacheService.getScriptCache().removeAll([]); } catch(e){}
  try { CacheService.getUserCache().removeAll([]); } catch(e){}
}
function cachedRead(key, fn) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch(e){} }
  const val = fn();
  try { cache.put(key, JSON.stringify(val), GD_CACHE_TTL); } catch(e){}
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
    logAudit(getCurrentUser().data.email, 'config-update', 'Config', key, { value:value });
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
    for (let i=0;i<users.length;i++) if (String(users[i].email||'').toLowerCase() === email) { me = users[i]; break; }
    if (!me) {
      // Unknown Google identity → treat as Guard-level operator so the gate is never locked out,
      // but with no admin powers. Admins add real accounts to the Users sheet.
      me = { id:0, email:email||'anonymous', name:email||'Gate Operator', role:'Guard', active:true };
    }
    return success(me);
  });
}
function requireRole(allowed) {
  const me = getCurrentUser().data;
  if (!me || allowed.indexOf(me.role) === -1) {
    throw new Error('Forbidden: requires one of ' + allowed.join(', '));
  }
  return me;
}

// ---------------------------------------------------------------------------
// Audit + Error logging
// ---------------------------------------------------------------------------
function logAudit(actor, action, entity, entityId, details) {
  try {
    appendRowFromObject('AuditLog', {
      ts: new Date().toISOString(), actor:actor, action:action, entity:entity, entityId:entityId,
      details: (typeof details === 'string') ? details : JSON.stringify(details)
    });
  } catch(e) {}
}
function logError(fn, err) {
  try {
    appendRowFromObject('ErrorLog', {
      ts: new Date().toISOString(), fn:fn,
      message: (err && err.message) ? err.message : String(err),
      stack: (err && err.stack) ? err.stack : ''
    });
  } catch(e){}
}


// ---------------------------------------------------------------------------
// bootstrap — single startup round trip
// ---------------------------------------------------------------------------
function bootstrap() {
  return safeCall('bootstrap', function(){
    const me = getCurrentUser().data;
    const cfg = {};
    getSheetAsObjects('Config').forEach(function(r){ cfg[r.key] = r.value; });
    const hosts = getSheetAsObjects('Hosts').filter(function(h){ return h.active === true || String(h.active).toLowerCase()==='true'; });
    return success({
      me: me,
      config: cfg,
      hosts: hosts,
      statuses: GD_STATUS,
      sources: GD_SOURCES,
      roles: GD_ROLES,
      checkinUrl: (cfg['checkin.url'] || (getWebAppUrl() ? getWebAppUrl() + '?page=checkin' : ''))
    });
  });
}

// Public bootstrap for the unauthenticated check-in form (no role needed)
function bootstrapPublic() {
  return safeCall('bootstrapPublic', function(){
    const cfg = {};
    getSheetAsObjects('Config').forEach(function(r){ cfg[r.key] = r.value; });
    const hosts = getSheetAsObjects('Hosts')
      .filter(function(h){ return h.active === true || String(h.active).toLowerCase()==='true'; })
      .map(function(h){ return { id:h.id, name:h.name, department:h.department }; });
    return success({ siteName: cfg['site.name'] || 'GateDesk', hosts: hosts });
  });
}

// ---------------------------------------------------------------------------
// Pagination — range-based
// ---------------------------------------------------------------------------
function getPagedData(entity, page, pageSize, filters, sortKey, sortDir) {
  return safeCall('getPagedData', function(){
    page = Math.max(1, Number(page)||1);
    pageSize = Math.min(200, Math.max(5, Number(pageSize)||25));
    const cacheKey = 'page:'+entity+':'+page+':'+pageSize+':'+JSON.stringify(filters||{})+':'+(sortKey||'')+':'+(sortDir||'');
    return cachedRead(cacheKey, function(){
      let all = getSheetAsObjects(entity).filter(function(r){ return !r.deleted; });
      // enrich VisitLogs with visitor + host names for display/search
      if (entity === 'VisitLogs') all = enrichLogs(all);
      if (filters) {
        Object.keys(filters).forEach(function(k){
          const v = filters[k];
          if (v === '' || v === null || v === undefined) return;
          if (k === '_from') { all = all.filter(function(r){ return String(r.createdAt||'') >= v; }); return; }
          if (k === '_to')   { all = all.filter(function(r){ return String(r.createdAt||'') <= (v + 'T23:59:59'); }); return; }
          if (k === '_q') {
            const q = String(v).toLowerCase();
            all = all.filter(function(r){ return ['visitorName','visitorPhone','company','hostName','purpose','status'].some(function(f){ return String(r[f]||'').toLowerCase().indexOf(q)!==-1; }); });
            return;
          }
          all = all.filter(function(r){ return String(r[k]||'').toLowerCase() === String(v).toLowerCase(); });
        });
      }
      if (sortKey) {
        const dir = (sortDir === 'asc') ? 1 : -1;
        all.sort(function(a,b){ const av=a[sortKey], bv=b[sortKey]; if (av===bv) return 0; return (av>bv?1:-1)*dir; });
      } else {
        all.sort(function(a,b){ return (a.createdAt < b.createdAt) ? 1 : -1; });
      }
      const total = all.length;
      const start = (page-1)*pageSize;
      return { rows: all.slice(start, start+pageSize), total:total, page:page, pageSize:pageSize };
    });
  });
}
function enrichLogs(logs) {
  const visitors = getSheetAsObjects('Visitors');
  const hosts = getSheetAsObjects('Hosts');
  const vById = {}; visitors.forEach(function(v){ vById[String(v.id)] = v; });
  const hById = {}; hosts.forEach(function(h){ hById[String(h.id)] = h; });
  return logs.map(function(l){
    const v = vById[String(l.visitorId)] || {};
    const h = hById[String(l.hostId)] || {};
    return Object.assign({}, l, {
      visitorName: v.name || '', visitorPhone: v.phone || '', company: v.company || '', photoUrl: v.photoUrl || '',
      idProofLast4: v.idProofLast4 || '', hostName: h.name || '', hostDepartment: h.department || ''
    });
  });
}

// ---------------------------------------------------------------------------
// Hosts CRUD
// ---------------------------------------------------------------------------
function listHosts() { return safeCall('listHosts', function(){ return success(getSheetAsObjects('Hosts')); }); }
function createHost(payload) {
  return safeCall('createHost', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const host = { name:payload.name, phone:payload.phone||'', department:payload.department||'', active:true };
      const saved = appendRowFromObject('Hosts', host);
      logAudit(me.email,'create','Hosts',saved.id,host);
      invalidateEntityCache('Hosts');
      return success(saved);
    } finally { lock.releaseLock(); }
  });
}
function updateHost(id, patch) {
  return safeCall('updateHost', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const updated = updateRowById('Hosts','id',id,patch);
      logAudit(me.email,'update','Hosts',id,patch);
      invalidateEntityCache('Hosts');
      return success(updated);
    } finally { lock.releaseLock(); }
  });
}
function deactivateHost(id) {
  return safeCall('deactivateHost', function(){
    const me = requireRole(['Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const updated = updateRowById('Hosts','id',id,{ active:false });
      logAudit(me.email,'deactivate','Hosts',id,'');
      invalidateEntityCache('Hosts');
      return success(updated);
    } finally { lock.releaseLock(); }
  });
}

// ---------------------------------------------------------------------------
// Visitors directory — returning-visitor auto-fill by phone
// ---------------------------------------------------------------------------
function lookupVisitorByPhone(phone) {
  return safeCall('lookupVisitorByPhone', function(){
    const p = String(phone||'').replace(/\D/g,'').slice(-10);
    if (p.length < 10) return success(null);
    const visitors = getSheetAsObjects('Visitors');
    for (let i=0;i<visitors.length;i++) {
      if (String(visitors[i].phone||'').replace(/\D/g,'').slice(-10) === p) return success(visitors[i]);
    }
    return success(null);
  });
}
function upsertVisitor_(payload) {
  // internal — assumes caller holds a lock
  const p = String(payload.phone||'').replace(/\D/g,'').slice(-10);
  const visitors = getSheetAsObjects('Visitors');
  let existing = null;
  for (let i=0;i<visitors.length;i++) if (String(visitors[i].phone||'').replace(/\D/g,'').slice(-10) === p) { existing = visitors[i]; break; }
  const fields = {
    phone: p, name: payload.name||'', company: payload.company||'',
    photoUrl: payload.photoUrl||'', idProofLast4: String(payload.idProofLast4||'').slice(-4)
  };
  if (existing) {
    // keep existing photo if none supplied this time
    if (!fields.photoUrl) fields.photoUrl = existing.photoUrl || '';
    updateRowById('Visitors','id',existing.id,fields);
    return existing.id;
  }
  const saved = appendRowFromObject('Visitors', Object.assign({ createdAt:new Date().toISOString() }, fields));
  return saved.id;
}


// ---------------------------------------------------------------------------
// Photo upload → Drive → URL
// ---------------------------------------------------------------------------
function savePhoto_(base64DataUrl, label) {
  if (!base64DataUrl || String(base64DataUrl).indexOf('base64,') === -1) return '';
  try {
    const folderName = getConfig('photo.folderName') || 'GateDesk Visitor Photos';
    let folder;
    const it = DriveApp.getFoldersByName(folderName);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
    const parts = String(base64DataUrl).split('base64,');
    const meta = parts[0]; const data = parts[1];
    const contentType = (meta.match(/data:(.*?);/) || [,'image/jpeg'])[1];
    const bytes = Utilities.base64Decode(data);
    const blob = Utilities.newBlob(bytes, contentType, (label||'visitor') + '_' + Date.now() + '.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch (e) { logError('savePhoto_', e); return ''; }
}

// ---------------------------------------------------------------------------
// Public self-registration (unauthenticated — called from CheckinForm)
// ---------------------------------------------------------------------------
function submitCheckin(payload) {
  return safeCall('submitCheckin', function(){
    if (!payload || !payload.phone || !payload.name) return fail('Name and phone are required');
    if (!payload.hostId) return fail('Please select a host to meet');
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      let photoUrl = payload.photoUrl || '';
      if (payload.photoBase64) photoUrl = savePhoto_(payload.photoBase64, payload.name);
      const visitorId = upsertVisitor_({
        phone: payload.phone, name: payload.name, company: payload.company,
        photoUrl: photoUrl, idProofLast4: payload.idProofLast4
      });
      const now = new Date().toISOString();
      const log = {
        visitorId: visitorId, hostId: payload.hostId, purpose: payload.purpose || '',
        checkInTime: '', checkOutTime: '', durationMins: '',
        status: 'Pending Approval', approvalMethod: '', source: 'QR Self-Registration',
        createdAt: now, deleted: false
      };
      const saved = appendRowFromObject('VisitLogs', log);
      logAudit(payload.phone,'self-register','VisitLogs',saved.id,{ host:payload.hostId });
      invalidateEntityCache('VisitLogs');
      Backend_Notify.notifyHost(saved.id);
      return success({ visitLogId: saved.id, status:'Pending Approval' });
    } finally { lock.releaseLock(); }
  });
}

// ---------------------------------------------------------------------------
// Guard manual check-in (visitor without a smartphone) — role: Guard/Admin
// ---------------------------------------------------------------------------
function guardManualCheckin(payload) {
  return safeCall('guardManualCheckin', function(){
    const me = requireRole(['Guard','Admin']);
    if (!payload || !payload.phone || !payload.name) return fail('Name and phone are required');
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const visitorId = upsertVisitor_({ phone:payload.phone, name:payload.name, company:payload.company, photoUrl:payload.photoUrl||'', idProofLast4:payload.idProofLast4 });
      const now = new Date().toISOString();
      // Guard entry goes straight Inside (physical presence verified at the gate)
      const log = {
        visitorId: visitorId, hostId: payload.hostId||'', purpose: payload.purpose||'',
        checkInTime: now, checkOutTime:'', durationMins:'',
        status:'Inside', approvalMethod:'Guard Override', source:'Guard Manual Entry',
        createdAt: now, deleted:false
      };
      const saved = appendRowFromObject('VisitLogs', log);
      logAudit(me.email,'manual-checkin','VisitLogs',saved.id,{ visitorId:visitorId });
      invalidateEntityCache('VisitLogs');
      return success(saved);
    } finally { lock.releaseLock(); }
  });
}

// ---------------------------------------------------------------------------
// Host approve / reject — role: Host (own visits) / Admin (any)
// ---------------------------------------------------------------------------
function getVisitLog(id) {
  return safeCall('getVisitLog', function(){
    const rows = enrichLogs(getSheetAsObjects('VisitLogs').filter(function(r){ return !r.deleted; }));
    const found = rows.filter(function(r){ return String(r.id)===String(id); })[0];
    return found ? success(found) : fail('Visit not found');
  });
}
function approveVisit(id, method) {
  return safeCall('approveVisit', function(){
    const me = requireRole(['Host','Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const log = getSheetAsObjects('VisitLogs').filter(function(r){ return String(r.id)===String(id); })[0];
      if (!log) return fail('Visit not found');
      if (me.role === 'Host' && !hostOwnsVisit_(me, log)) return fail('This visit is addressed to another host');
      if (log.status !== 'Pending Approval') return fail('Visit is not pending approval');
      const now = new Date().toISOString();
      const updated = updateRowById('VisitLogs','id',id,{ status:'Inside', checkInTime:now, approvalMethod: method||'In-App' });
      logAudit(me.email,'approve','VisitLogs',id,'');
      invalidateEntityCache('VisitLogs');
      return success(updated);
    } finally { lock.releaseLock(); }
  });
}
function rejectVisit(id, reason) {
  return safeCall('rejectVisit', function(){
    const me = requireRole(['Host','Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const log = getSheetAsObjects('VisitLogs').filter(function(r){ return String(r.id)===String(id); })[0];
      if (!log) return fail('Visit not found');
      if (me.role === 'Host' && !hostOwnsVisit_(me, log)) return fail('This visit is addressed to another host');
      const updated = updateRowById('VisitLogs','id',id,{ status:'Rejected', approvalMethod:'In-App', purpose: log.purpose + (reason ? (' — Rejected: '+reason) : '') });
      logAudit(me.email,'reject','VisitLogs',id,{ reason:reason||'' });
      invalidateEntityCache('VisitLogs');
      return success(updated);
    } finally { lock.releaseLock(); }
  });
}
function hostOwnsVisit_(me, log) {
  // Match the logged-in Host user to their Hosts-sheet row by name (Users.name === Hosts.name)
  const hosts = getSheetAsObjects('Hosts');
  const mine = hosts.filter(function(h){ return String(h.name||'').toLowerCase() === String(me.name||'').toLowerCase(); });
  return mine.some(function(h){ return String(h.id) === String(log.hostId); });
}

// ---------------------------------------------------------------------------
// Check-out + duration
// ---------------------------------------------------------------------------
function checkOutVisit(id) {
  return safeCall('checkOutVisit', function(){
    const me = requireRole(['Guard','Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const log = getSheetAsObjects('VisitLogs').filter(function(r){ return String(r.id)===String(id); })[0];
      if (!log) return fail('Visit not found');
      if (log.status !== 'Inside') return fail('Only visitors currently Inside can be checked out');
      const now = new Date();
      const inTime = log.checkInTime ? new Date(log.checkInTime) : now;
      const durationMins = Math.max(0, Math.round((now - inTime)/60000));
      const updated = updateRowById('VisitLogs','id',id,{ status:'Checked Out', checkOutTime: now.toISOString(), durationMins: durationMins });
      logAudit(me.email,'checkout','VisitLogs',id,{ durationMins:durationMins });
      invalidateEntityCache('VisitLogs');
      return success(updated);
    } finally { lock.releaseLock(); }
  });
}

// ---------------------------------------------------------------------------
// Live counts (Guard board) — cached, refreshed on action + poll
// ---------------------------------------------------------------------------
function getLiveCounts() {
  return safeCall('getLiveCounts', function(){
    return cachedRead('live:counts', function(){
      const today = new Date().toISOString().split('T')[0];
      const logs = getSheetAsObjects('VisitLogs').filter(function(r){ return !r.deleted; });
      const todays = logs.filter(function(l){ return String(l.createdAt||'').split('T')[0] === today; });
      return {
        todayTotal: todays.length,
        inside: logs.filter(function(l){ return l.status==='Inside'; }).length,
        checkedOut: logs.filter(function(l){ return l.status==='Checked Out'; }).length,
        pending: logs.filter(function(l){ return l.status==='Pending Approval'; }).length
      };
    });
  });
}
function getInsideList() {
  return safeCall('getInsideList', function(){
    const rows = enrichLogs(getSheetAsObjects('VisitLogs').filter(function(r){ return !r.deleted && r.status==='Inside'; }));
    rows.sort(function(a,b){ return (a.checkInTime < b.checkInTime) ? 1 : -1; });
    return success(rows);
  });
}
function getPendingApprovals() {
  return safeCall('getPendingApprovals', function(){
    const me = getCurrentUser().data;
    let rows = enrichLogs(getSheetAsObjects('VisitLogs').filter(function(r){ return !r.deleted && r.status==='Pending Approval'; }));
    if (me.role === 'Host') rows = rows.filter(function(r){ return hostOwnsVisit_(me, r); });
    rows.sort(function(a,b){ return (a.createdAt < b.createdAt) ? 1 : -1; });
    return success(rows);
  });
}


// ---------------------------------------------------------------------------
// Backend_Notify — pluggable host notification (Option A email / Option B BSP)
// ---------------------------------------------------------------------------
const Backend_Notify = (function(){
  function hostContact_(hostId) {
    const hosts = getSheetAsObjects('Hosts');
    const h = hosts.filter(function(x){ return String(x.id)===String(hostId); })[0] || {};
    const users = getSheetAsObjects('Users');
    const u = users.filter(function(x){ return String(x.name||'').toLowerCase() === String(h.name||'').toLowerCase() && x.role==='Host'; })[0] || {};
    return { name:h.name||'', phone:h.phone||'', email:u.email||'' };
  }
  function notifyHost(visitLogId) {
    try {
      if (String(getConfig('notify.enabled')||'true').toLowerCase() === 'false') return;
      const channel = String(getConfig('notify.channel')||'A').toUpperCase();
      if (channel === 'B') return optionB_(visitLogId);
      return optionA_(visitLogId);
    } catch(e){ logError('Backend_Notify.notifyHost', e); }
  }
  // Option A — free email via MailApp with approve/reject deep link (+ optional wa.me)
  function optionA_(visitLogId) {
    const enriched = enrichLogs(getSheetAsObjects('VisitLogs').filter(function(r){ return String(r.id)===String(visitLogId); }))[0];
    if (!enriched) return;
    const host = hostContact_(enriched.hostId);
    if (!host.email) { logError('Backend_Notify.optionA_','Host has no email in Users sheet: '+host.name); return; }
    const base = getWebAppUrl();
    const link = base ? (base + '?#visit=' + visitLogId) : '(open GateDesk)';
    const wa = host.phone ? ('https://wa.me/' + String(host.phone).replace(/\D/g,'')) : '';
    const html = '<div style="font-family:Arial,sans-serif;max-width:520px">' +
      '<h2 style="color:#0B1220">Visitor awaiting your approval</h2>' +
      '<p><b>' + escapeHtml_(enriched.visitorName) + '</b> from <b>' + escapeHtml_(enriched.company||'—') + '</b> is at the gate to meet you.</p>' +
      '<table style="font-size:14px;border-collapse:collapse">' +
      '<tr><td style="color:#64748b;padding:4px 12px 4px 0">Phone</td><td>' + escapeHtml_(enriched.visitorPhone) + '</td></tr>' +
      '<tr><td style="color:#64748b;padding:4px 12px 4px 0">Purpose</td><td>' + escapeHtml_(enriched.purpose||'—') + '</td></tr>' +
      '</table>' +
      '<p style="margin-top:16px"><a href="' + link + '" style="background:#0B1220;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none">Open GateDesk to Approve / Reject</a></p>' +
      (wa ? '<p style="font-size:12px;color:#64748b">Or reply on WhatsApp: <a href="' + wa + '">' + wa + '</a></p>' : '') +
      '<p style="font-size:11px;color:#94a3b8;margin-top:20px">Automated notification · GateDesk</p></div>';
    MailApp.sendEmail({ to: host.email, subject: 'GateDesk · Visitor for you: ' + enriched.visitorName, htmlBody: html });
  }
  // Option B — real WhatsApp via a BSP REST API (scaffold; token in Script Properties)
  function optionB_(visitLogId) {
    const endpoint = getConfig('notify.bspEndpoint');
    const token = PropertiesService.getScriptProperties().getProperty('BSP_TOKEN');
    if (!endpoint || !token) { logError('Backend_Notify.optionB_','Missing notify.bspEndpoint (Config) or BSP_TOKEN (Script Property) — falling back to email'); return optionA_(visitLogId); }
    const enriched = enrichLogs(getSheetAsObjects('VisitLogs').filter(function(r){ return String(r.id)===String(visitLogId); }))[0];
    if (!enriched) return;
    const host = hostContact_(enriched.hostId);
    const payload = {
      to: String(host.phone).replace(/\D/g,''),
      template: 'visitor_approval',
      params: { visitor: enriched.visitorName, company: enriched.company||'', purpose: enriched.purpose||'', visitId: String(visitLogId) }
    };
    UrlFetchApp.fetch(endpoint, {
      method:'post', contentType:'application/json',
      headers:{ Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload), muteHttpExceptions:true
    });
  }
  return { notifyHost: notifyHost };
})();

function escapeHtml_(v){ return String(v==null?'':v).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }

// doPost webhook — Option B: BSP posts the host's WhatsApp reply here
function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) ? JSON.parse(e.postData.contents) : {};
    const visitId = body.visitId || (body.params && body.params.visitId);
    const decision = String(body.decision||body.reply||'').toLowerCase();
    if (visitId && decision) {
      const lock = LockService.getScriptLock(); lock.waitLock(10000);
      try {
        if (decision.indexOf('approve') !== -1 || decision === 'yes' || decision === '1') {
          updateRowById('VisitLogs','id',visitId,{ status:'Inside', checkInTime:new Date().toISOString(), approvalMethod:'WhatsApp' });
        } else if (decision.indexOf('reject') !== -1 || decision === 'no' || decision === '2') {
          updateRowById('VisitLogs','id',visitId,{ status:'Rejected', approvalMethod:'WhatsApp' });
        }
        invalidateEntityCache('VisitLogs');
      } finally { lock.releaseLock(); }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok:true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    logError('doPost', err);
    return ContentService.createTextOutput(JSON.stringify({ ok:false, error:err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ---------------------------------------------------------------------------
// Admin: analytics, CSV export, QR
// ---------------------------------------------------------------------------
function getAnalytics() {
  return safeCall('getAnalytics', function(){
    requireRole(['Admin']);
    return cachedRead('analytics:v1', function(){
      const logs = getSheetAsObjects('VisitLogs').filter(function(r){ return !r.deleted; });
      const done = logs.filter(function(l){ return l.status==='Checked Out' && Number(l.durationMins)>0; });
      const avgDuration = done.length ? Math.round(done.reduce(function(s,l){ return s+Number(l.durationMins); },0)/done.length) : 0;
      const byHour = {}; for (let h=0;h<24;h++) byHour[h]=0;
      logs.forEach(function(l){ const t = l.checkInTime||l.createdAt; if (t){ const hr = new Date(t).getHours(); byHour[hr] = (byHour[hr]||0)+1; } });
      let busiest = 0, busiestCount = -1;
      Object.keys(byHour).forEach(function(h){ if (byHour[h] > busiestCount){ busiestCount = byHour[h]; busiest = Number(h); } });
      const byStatus = {}; GD_STATUS.forEach(function(s){ byStatus[s]=0; });
      logs.forEach(function(l){ byStatus[l.status] = (byStatus[l.status]||0)+1; });
      return {
        totalVisits: logs.length, avgDurationMins: avgDuration,
        busiestHour: busiest, busiestHourCount: Math.max(0,busiestCount),
        byStatus: byStatus,
        hourly: Object.keys(byHour).map(function(h){ return { hour:Number(h), count:byHour[h] }; })
      };
    });
  });
}
function exportVisitLogsCsv(filters) {
  return safeCall('exportVisitLogsCsv', function(){
    requireRole(['Admin']);
    let rows = enrichLogs(getSheetAsObjects('VisitLogs').filter(function(r){ return !r.deleted; }));
    if (filters) {
      if (filters.status) rows = rows.filter(function(r){ return r.status===filters.status; });
      if (filters.hostId) rows = rows.filter(function(r){ return String(r.hostId)===String(filters.hostId); });
      if (filters._from) rows = rows.filter(function(r){ return String(r.createdAt||'') >= filters._from; });
      if (filters._to)   rows = rows.filter(function(r){ return String(r.createdAt||'') <= (filters._to+'T23:59:59'); });
    }
    const cols = ['id','visitorName','visitorPhone','company','hostName','purpose','status','source','checkInTime','checkOutTime','durationMins','createdAt'];
    const csv = cols.join(',') + '\n' + rows.map(function(r){
      return cols.map(function(c){ const v = r[c]==null?'':String(r[c]); return '"'+v.replace(/"/g,'""')+'"'; }).join(',');
    }).join('\n');
    return success({ filename:'gatedesk-visitlogs.csv', csv:csv, count:rows.length });
  });
}
function getQrInfo() {
  return safeCall('getQrInfo', function(){
    requireRole(['Admin']);
    const url = getConfig('checkin.url') || (getWebAppUrl() ? getWebAppUrl()+'?page=checkin' : '');
    const qr = url ? ('https://api.qrserver.com/v1/create-qr-code/?size=280x280&data=' + encodeURIComponent(url)) : '';
    return success({ checkinUrl:url, qrImageUrl:qr, siteName: getConfig('site.name')||'GateDesk' });
  });
}

// ---------------------------------------------------------------------------
// Soft-delete + restore (Admin) for VisitLogs
// ---------------------------------------------------------------------------
function softDeleteVisit(id) {
  return safeCall('softDeleteVisit', function(){
    const me = requireRole(['Admin']);
    const updated = updateRowById('VisitLogs','id',id,{ deleted:true });
    logAudit(me.email,'delete','VisitLogs',id,'soft-delete');
    invalidateEntityCache('VisitLogs');
    return success(updated);
  });
}
function restoreVisit(id) {
  return safeCall('restoreVisit', function(){
    const me = requireRole(['Admin']);
    const updated = updateRowById('VisitLogs','id',id,{ deleted:false });
    logAudit(me.email,'restore','VisitLogs',id,'undo');
    invalidateEntityCache('VisitLogs');
    return success(updated);
  });
}

// ---------------------------------------------------------------------------
// Triggers — overstay sweep
// ---------------------------------------------------------------------------
function overstaySweep() {
  return safeCall('overstaySweep', function(){
    const cutoff = String(getConfig('overstay.cutoff')||'20:00');
    const parts = cutoff.split(':'); const ch = Number(parts[0])||20, cm = Number(parts[1])||0;
    const now = new Date();
    const cutoffToday = new Date(now); cutoffToday.setHours(ch, cm, 0, 0);
    const logs = enrichLogs(getSheetAsObjects('VisitLogs').filter(function(r){ return !r.deleted && r.status==='Inside'; }));
    const overstayers = logs.filter(function(l){ return now > cutoffToday && l.checkInTime && new Date(l.checkInTime) < cutoffToday; });
    if (overstayers.length) {
      logAudit('system','overstay-sweep','VisitLogs','',{ count:overstayers.length, names:overstayers.map(function(o){return o.visitorName;}) });
      const admins = getSheetAsObjects('Users').filter(function(u){ return u.role==='Admin' && u.email; });
      if (admins.length && String(getConfig('notify.enabled')||'true').toLowerCase() !== 'false') {
        const list = overstayers.map(function(o){ return '• '+o.visitorName+' ('+o.company+') — in since '+o.checkInTime; }).join('<br>');
        MailApp.sendEmail({ to: admins.map(function(a){return a.email;}).join(','), subject:'GateDesk · '+overstayers.length+' visitor(s) still inside after cutoff', htmlBody:'<p>The following visitors are still marked Inside past '+cutoff+':</p><p>'+list+'</p>' });
      }
    }
    return success({ overstayers: overstayers.length });
  });
}
function installOverstayTrigger() {
  removeOverstayTrigger();
  ScriptApp.newTrigger('overstaySweep').timeBased().atHour(21).everyDays(1).create();
  SpreadsheetApp.getUi().alert('Overstay trigger installed (runs daily ~21:00).');
}
function removeOverstayTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction()==='overstaySweep') ScriptApp.deleteTrigger(t); });
}
