/**
 * ============================================================================
 *  ORBIT CRM — Google Apps Script Web App (single-spreadsheet backend)
 * ============================================================================
 *
 *  DEPLOYMENT
 *  ----------
 *  1. Open Google Sheets → Extensions → Apps Script.
 *  2. Paste this file as Code.gs and paste Index.html as an HTML file named
 *     exactly "Index".
 *  3. Reload the sheet; a custom menu "⚙️ Setup" appears. Click
 *     "Initialize System" once (safe to run again — idempotent).
 *  4. Deploy → New deployment → Web app → Execute as: Me, Access: Anyone with
 *     the link (or your domain). Open the resulting URL.
 *  5. Optional: run "Install Daily SLA Trigger" from the same menu once to
 *     wire the 07:00 daily follow-up sweep.
 *
 *  SHEET STRUCTURE (created by initializeSheets)
 *  ---------------------------------------------
 *   Users            id | email | name | role | active | createdAt
 *   Leads            id | name | email | phone | company | source | stage |
 *                    score | ownerEmail | lastContactedDate | createdAt |
 *                    deleted
 *   Accounts         id | name | industry | website | phone | ownerEmail |
 *                    createdAt | deleted
 *   Contacts         id | name | email | phone | title | accountId |
 *                    ownerEmail | createdAt | deleted
 *   Deals            id | title | leadId | accountId | value | probability |
 *                    expectedCloseDate | stage | ownerEmail | createdAt |
 *                    deleted
 *   Activities       id | type | subject | body | leadId | contactId |
 *                    accountId | dealId | createdBy | createdAt
 *   Config           key | value
 *   AuditLog         id | ts | actor | action | entity | entityId | details
 *   ErrorLog         id | ts | fn | message | stack
 *   UserPreferences  email | prefs
 *
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const SHEETS = {
  Users: {
    headers: ['id','email','name','role','active','createdAt'],
    widths:  [60,220,180,140,80,180]
  },
  Leads: {
    headers: ['id','name','email','phone','company','source','stage','score','ownerEmail','lastContactedDate','createdAt','deleted'],
    widths:  [60,180,220,140,180,120,140,80,220,180,180,80]
  },
  Accounts: {
    headers: ['id','name','industry','website','phone','ownerEmail','createdAt','deleted'],
    widths:  [60,200,160,200,140,220,180,80]
  },
  Contacts: {
    headers: ['id','name','email','phone','title','accountId','ownerEmail','createdAt','deleted'],
    widths:  [60,180,220,140,160,80,220,180,80]
  },
  Deals: {
    headers: ['id','title','leadId','accountId','value','probability','expectedCloseDate','stage','ownerEmail','createdAt','deleted'],
    widths:  [60,220,80,80,120,100,160,140,220,180,80]
  },
  Activities: {
    headers: ['id','type','subject','body','leadId','contactId','accountId','dealId','createdBy','createdAt'],
    widths:  [60,110,220,320,80,80,80,80,220,180]
  },
  Config: {
    headers: ['key','value'],
    widths:  [260,320]
  },
  AuditLog: {
    headers: ['id','ts','actor','action','entity','entityId','details'],
    widths:  [60,180,220,140,120,80,420]
  },
  ErrorLog: {
    headers: ['id','ts','fn','message','stack'],
    widths:  [60,180,220,420,520]
  },
  UserPreferences: {
    headers: ['email','prefs'],
    widths:  [260,520]
  }
};

const STAGES  = ['New','Contacted','Qualified','Proposal Sent','Won','Lost'];
const SOURCES = ['Website','Referral','Cold Call','Event','LinkedIn','Partner','Ads'];
const ROLES   = ['Sales Rep','Sales Manager','Admin'];

// ---------------------------------------------------------------------------
// Web app entry + menu
// ---------------------------------------------------------------------------
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Orbit CRM')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Setup')
    .addItem('Initialize System', 'menuInitialize')
    .addItem('Reset Dummy Data',  'menuResetDummy')
    .addSeparator()
    .addItem('Install Daily SLA Trigger', 'installSlaTrigger')
    .addItem('Remove SLA Trigger',        'removeSlaTrigger')
    .addToUi();
}

function menuInitialize() {
  const r = initializeSheets();
  SpreadsheetApp.getUi().alert('Initialized.\n\nCreated: ' + r.data.sheetsCreated.join(', ') +
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
    Object.keys(SHEETS).forEach(function(name){
      let sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
        const spec = SHEETS[name];
        sh.getRange(1,1,1,spec.headers.length).setValues([spec.headers])
          .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
        sh.setFrozenRows(1);
        spec.widths.forEach(function(w,i){ sh.setColumnWidth(i+1, w); });
        created.push(name);
      } else {
        skipped.push(name);
      }
    });
    // strip the default Sheet1 if present and empty
    const s1 = ss.getSheetByName('Sheet1');
    if (s1 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s1); } catch(e){}
    }

    // ---- seed Config with defaults (only if empty) ----
    seedConfigDefaults();

    // ---- seed each core sheet if it is currently empty (only headers) ----
    let dummyRowsAdded = 0;
    dummyRowsAdded += seedUsers();
    dummyRowsAdded += seedAccounts();
    dummyRowsAdded += seedContacts();
    dummyRowsAdded += seedLeads();
    dummyRowsAdded += seedDeals();
    dummyRowsAdded += seedActivities();

    return success({ sheetsCreated:created, sheetsSkipped:skipped, dummyRowsAdded:dummyRowsAdded });
  });
}

function seedConfigDefaults() {
  const sh = getSheet('Config');
  const existing = getSheetAsObjects('Config').reduce(function(acc,r){acc[r.key]=true;return acc;},{});
  const defaults = [
    ['sla.followUpDays','5'],
    ['sla.escalateMultiplier','2'],
    ['roundRobin.pointer','0'],
    ['app.name','Orbit CRM'],
    ['app.version','1.0.0']
  ];
  const rows = defaults.filter(function(kv){ return !existing[kv[0]]; });
  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,2).setValues(rows);
}

function seedUsers() {
  const sh = getSheet('Users');
  if (sh.getLastRow() > 1) return 0;
  const now = new Date().toISOString();
  const rows = [
    [1,'rep1@example.com','Alex Rivera','Sales Rep',true,now],
    [2,'rep2@example.com','Priya Nair','Sales Rep',true,now],
    [3,'rep3@example.com','Diego Sato','Sales Rep',true,now],
    [4,'manager@example.com','Morgan Lee','Sales Manager',true,now],
    [5,'admin@example.com','Root Admin','Admin',true,now]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

function seedAccounts() {
  const sh = getSheet('Accounts');
  if (sh.getLastRow() > 1) return 0;
  const now = new Date().toISOString();
  const rows = [
    [1,'Northwind Traders','Retail','northwind.com','+1-415-555-0110','rep1@example.com',now,false],
    [2,'Contoso Ltd','Manufacturing','contoso.com','+1-415-555-0112','rep2@example.com',now,false],
    [3,'Fabrikam Inc','Aerospace','fabrikam.com','+1-415-555-0113','rep3@example.com',now,false],
    [4,'Adventure Works','Sports','adventureworks.com','+1-415-555-0114','rep1@example.com',now,false],
    [5,'Litware','Software','litware.com','+1-415-555-0115','rep2@example.com',now,false],
    [6,'Tailspin Toys','Retail','tailspintoys.com','+1-415-555-0116','rep3@example.com',now,false],
    [7,'Wingtip Partners','Consulting','wingtip.com','+1-415-555-0117','rep1@example.com',now,false],
    [8,'Proseware','Media','proseware.com','+1-415-555-0118','rep2@example.com',now,false]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

function seedContacts() {
  const sh = getSheet('Contacts');
  if (sh.getLastRow() > 1) return 0;
  const now = new Date().toISOString();
  const rows = [
    [1,'Julia Ortiz','julia@northwind.com','+1-415-555-1101','VP Sales',1,'rep1@example.com',now,false],
    [2,'Ken Watanabe','ken@contoso.com','+1-415-555-1102','CTO',2,'rep2@example.com',now,false],
    [3,'Lena Park','lena@fabrikam.com','+1-415-555-1103','Procurement',3,'rep3@example.com',now,false],
    [4,'Marco Bianchi','marco@adventureworks.com','+1-415-555-1104','COO',4,'rep1@example.com',now,false],
    [5,'Nadia Haddad','nadia@litware.com','+1-415-555-1105','Head of Ops',5,'rep2@example.com',now,false],
    [6,'Oscar Kim','oscar@tailspintoys.com','+1-415-555-1106','Buyer',6,'rep3@example.com',now,false],
    [7,'Petra Schwarz','petra@wingtip.com','+1-415-555-1107','Managing Partner',7,'rep1@example.com',now,false],
    [8,'Quinn Adebayo','quinn@proseware.com','+1-415-555-1108','CEO',8,'rep2@example.com',now,false]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

function seedLeads() {
  const sh = getSheet('Leads');
  if (sh.getLastRow() > 1) return 0;
  const today = new Date();
  function iso(d){ return d.toISOString(); }
  function daysAgo(n){ const d=new Date(today); d.setDate(d.getDate()-n); return d; }
  const rows = [
    [1,'Ivy Chen','ivy@aperture.io','+1-415-555-2001','Aperture Science','Website','New',72,'rep1@example.com',iso(daysAgo(1)),iso(daysAgo(1)),false],
    [2,'Marcus Reid','marcus@umbra.co','+1-415-555-2002','Umbra Co','Referral','Contacted',65,'rep2@example.com',iso(daysAgo(3)),iso(daysAgo(6)),false],
    [3,'Sofia Alvarez','sofia@lumenlabs.com','+1-415-555-2003','Lumen Labs','LinkedIn','Qualified',88,'rep3@example.com',iso(daysAgo(10)),iso(daysAgo(14)),false],
    [4,'Ravi Kapoor','ravi@vertexworks.io','+1-415-555-2004','Vertex Works','Event','Proposal Sent',91,'rep1@example.com',iso(daysAgo(12)),iso(daysAgo(20)),false],
    [5,'Elena Popa','elena@northgate.io','+1-415-555-2005','Northgate','Cold Call','Won',95,'rep2@example.com',iso(daysAgo(4)),iso(daysAgo(25)),false],
    [6,'Tomás Ferreira','tomas@bluewire.co','+1-415-555-2006','Bluewire','Partner','Lost',22,'rep3@example.com',iso(daysAgo(30)),iso(daysAgo(45)),false],
    [7,'Hana Yamada','hana@quantic.io','+1-415-555-2007','Quantic','Website','Contacted',54,'rep1@example.com',iso(daysAgo(15)),iso(daysAgo(18)),false], // overdue
    [8,'Noah Weber','noah@fjord.io','+1-415-555-2008','Fjord Digital','Ads','New',48,'rep2@example.com',iso(daysAgo(20)),iso(daysAgo(20)),false],           // 2x overdue
    [9,'Amelia Brooks','amelia@atlaspay.com','+1-415-555-2009','Atlas Pay','Referral','Qualified',77,'rep3@example.com',iso(daysAgo(2)),iso(daysAgo(9)),false],
    [10,'Daniel Osei','daniel@helios.tech','+1-415-555-2010','Helios Tech','LinkedIn','Proposal Sent',83,'rep1@example.com',iso(daysAgo(6)),iso(daysAgo(22)),false]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

function seedDeals() {
  const sh = getSheet('Deals');
  if (sh.getLastRow() > 1) return 0;
  const today = new Date();
  function fut(n){ const d=new Date(today); d.setDate(d.getDate()+n); return d.toISOString(); }
  const now = new Date().toISOString();
  const rows = [
    [1,'Aperture — Starter Plan',1,null,12000,30,fut(20),'Qualified','rep1@example.com',now,false],
    [2,'Umbra — Growth Bundle',2,null,28000,50,fut(35),'Proposal Sent','rep2@example.com',now,false],
    [3,'Lumen — Enterprise',3,null,86000,70,fut(45),'Qualified','rep3@example.com',now,false],
    [4,'Vertex — Renewal',4,null,42000,80,fut(15),'Proposal Sent','rep1@example.com',now,false],
    [5,'Northgate — Onboarding',5,null,54000,100,fut(-5),'Won','rep2@example.com',now,false],
    [6,'Bluewire — POC',6,null,15000,0,fut(-2),'Lost','rep3@example.com',now,false],
    [7,'Quantic — Expansion',7,null,33000,45,fut(28),'Contacted','rep1@example.com',now,false],
    [8,'Fjord — Discovery',8,null,9000,20,fut(60),'New','rep2@example.com',now,false],
    [9,'Atlas Pay — Team Plan',9,null,22000,55,fut(40),'Qualified','rep3@example.com',now,false],
    [10,'Helios — Pilot',10,null,17500,60,fut(25),'Proposal Sent','rep1@example.com',now,false]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

function seedActivities() {
  const sh = getSheet('Activities');
  if (sh.getLastRow() > 1) return 0;
  const now = new Date();
  function ago(n){ const d=new Date(now); d.setDate(d.getDate()-n); return d.toISOString(); }
  const rows = [
    [1,'call','Intro call','Great chat, sending deck.',1,null,null,null,'rep1@example.com',ago(1)],
    [2,'email','Follow up','Sent proposal draft v1.',2,null,null,null,'rep2@example.com',ago(3)],
    [3,'meeting','Discovery meeting','Mapped stakeholders across 3 teams.',3,null,null,null,'rep3@example.com',ago(10)],
    [4,'note','Signal','Champion moved from Bluewire — warm handoff.',6,null,null,null,'rep3@example.com',ago(28)],
    [5,'email','Proposal sent','Sent MSA + pricing.',4,null,null,null,'rep1@example.com',ago(12)],
    [6,'call','Kickoff','Onboarding kickoff scheduled next Tue.',5,null,null,null,'rep2@example.com',ago(4)],
    [7,'meeting','Product demo','Demoed core workflows.',9,null,null,null,'rep3@example.com',ago(2)],
    [8,'note','Budget note','Q2 budget confirmed at 20k+.',10,null,null,null,'rep1@example.com',ago(5)]
  ];
  sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  return rows.length;
}

// ---------------------------------------------------------------------------
// resetDummyData — clears seeded rows (leaves anything you added yourself IF
// your ids are > 1000)
// ---------------------------------------------------------------------------
function resetDummyData() {
  return safeCall('resetDummyData', function(){
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      let removed = 0;
      const targets = ['Leads','Accounts','Contacts','Deals','Activities','Users'];
      targets.forEach(function(name){
        const sh = getSheet(name);
        const last = sh.getLastRow();
        if (last <= 1) return;
        const data = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
        const keep = data.filter(function(row){ return Number(row[0]) > 1000; });
        sh.getRange(2,1,last-1,sh.getLastColumn()).clearContent();
        if (keep.length) sh.getRange(2,1,keep.length,keep[0].length).setValues(keep);
        removed += (data.length - keep.length);
      });
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
    headers.forEach(function(h,i){ o[h] = r[i]; });
    return o;
  });
}

function appendRowFromObject(sheetName, obj) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  // auto id if not provided
  if (headers[0] === 'id' && (obj.id === undefined || obj.id === null || obj.id === '')) {
    obj.id = nextId(sheetName);
  }
  const row = headers.map(function(h){
    const v = obj[h];
    return (v === undefined || v === null) ? '' : v;
  });
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
      headers.forEach(function(h,i){
        if (updated.hasOwnProperty(h)) data[r][i] = updated[h];
      });
      sh.getRange(r+2,1,1,headers.length).setValues([data[r]]);
      const obj = {}; headers.forEach(function(h,i){ obj[h]=data[r][i]; });
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
  for (let i=0;i<rows.length;i++) {
    if (rows[i].key === key) { sh.getRange(i+2,2).setValue(value); return; }
  }
  sh.appendRow([key,value]);
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
      // fall back to a viewer identity if the current user is not in the Users sheet
      me = { id:0, email:email||'anonymous', name:email||'Guest', role:'Sales Rep', active:true };
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
      ts: new Date().toISOString(),
      actor: actor, action: action, entity: entity, entityId: entityId,
      details: (typeof details === 'string') ? details : JSON.stringify(details)
    });
  } catch(e) { /* never throw from audit */ }
}

function logError(fn, err) {
  try {
    appendRowFromObject('ErrorLog', {
      ts: new Date().toISOString(),
      fn: fn,
      message: (err && err.message) ? err.message : String(err),
      stack:   (err && err.stack)   ? err.stack   : ''
    });
  } catch(e){}
}

// ---------------------------------------------------------------------------
// Bootstrap: single endpoint the front-end calls at startup
// ---------------------------------------------------------------------------
function bootstrap() {
  return safeCall('bootstrap', function(){
    const me = getCurrentUser().data;
    const users = getSheetAsObjects('Users').filter(function(u){ return u.active; });
    const cfg = {};
    getSheetAsObjects('Config').forEach(function(r){ cfg[r.key] = r.value; });
    return success({
      me: me,
      users: users,
      config: cfg,
      stages: STAGES,
      sources: SOURCES,
      roles: ROLES
    });
  });
}

// ---------------------------------------------------------------------------
// Pagination
// ---------------------------------------------------------------------------
function getPagedData(entity, page, pageSize, filters, sortKey, sortDir) {
  return safeCall('getPagedData', function(){
    page = Math.max(1, Number(page)||1);
    pageSize = Math.min(200, Math.max(5, Number(pageSize)||25));
    const cacheKey = 'page:' + entity + ':' + page + ':' + pageSize + ':' +
                     JSON.stringify(filters||{}) + ':' + (sortKey||'') + ':' + (sortDir||'');
    const cache = CacheService.getUserCache();
    const cached = cache.get(cacheKey);
    if (cached) return success(JSON.parse(cached));

    let all = getSheetAsObjects(entity).filter(function(r){ return !r.deleted; });
    if (filters) {
      Object.keys(filters).forEach(function(k){
        const v = filters[k];
        if (v === '' || v === null || v === undefined) return;
        all = all.filter(function(r){ return String(r[k]||'').toLowerCase() === String(v).toLowerCase(); });
      });
    }
    if (sortKey) {
      const dir = (sortDir === 'desc') ? -1 : 1;
      all.sort(function(a,b){
        const av=a[sortKey], bv=b[sortKey];
        if (av === bv) return 0;
        return (av > bv ? 1 : -1) * dir;
      });
    }
    const total = all.length;
    const start = (page-1)*pageSize;
    const rows = all.slice(start, start+pageSize);
    const payload = { rows:rows, total:total, page:page, pageSize:pageSize };
    cache.put(cacheKey, JSON.stringify(payload), 45); // 45s
    return success(payload);
  });
}

function invalidateEntityCache(entity) {
  // Coarse: nuke the whole user cache. Cheap and effective for a CRM this size.
  try { CacheService.getUserCache().removeAll([]); } catch(e){}
}

// ---------------------------------------------------------------------------
// LEADS
// ---------------------------------------------------------------------------
function createLead(payload) {
  return safeCall('createLead', function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      // duplicate
      const dup = findDuplicate('Leads', payload);
      if (dup) return fail('Duplicate lead detected: ' + dup.field + ' already exists (id ' + dup.row.id + ')');

      const now = new Date().toISOString();
      const owner = payload.ownerEmail || roundRobinAssign();
      const lead = {
        name: payload.name, email: payload.email||'', phone: payload.phone||'',
        company: payload.company||'', source: payload.source||'Website',
        stage: payload.stage || 'New', score: Number(payload.score)||50,
        ownerEmail: owner, lastContactedDate: now, createdAt: now, deleted: false
      };
      const saved = appendRowFromObject('Leads', lead);
      logAudit(me.email,'create','Leads',saved.id,{name:saved.name,stage:saved.stage});
      invalidateEntityCache('Leads');
      return success(saved);
    } finally { lock.releaseLock(); }
  });
}

function updateLead(id, patch) {
  return safeCall('updateLead', function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const before = getById('Leads', id);
      const updated = updateRowById('Leads','id',id, patch);
      logAudit(me.email,'update','Leads',id,{before:before,after:updated});
      invalidateEntityCache('Leads');
      return success(updated);
    } finally { lock.releaseLock(); }
  });
}

function changeLeadStage(id, newStage) {
  return safeCall('changeLeadStage', function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    if (STAGES.indexOf(newStage) === -1) return fail('Invalid stage');
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      const before = getById('Leads', id);
      const updated = updateRowById('Leads','id',id,{ stage:newStage, lastContactedDate:new Date().toISOString() });
      logAudit(me.email,'stage-change','Leads',id,{from:before && before.stage, to:newStage});
      invalidateEntityCache('Leads');
      return success(updated);
    } finally { lock.releaseLock(); }
  });
}

function deleteLead(id) {
  return safeCall('deleteLead', function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    const updated = updateRowById('Leads','id',id,{ deleted:true });
    logAudit(me.email,'delete','Leads',id,'soft-delete');
    invalidateEntityCache('Leads');
    return success(updated);
  });
}

function restoreLead(id) {
  return safeCall('restoreLead', function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    const updated = updateRowById('Leads','id',id,{ deleted:false });
    logAudit(me.email,'restore','Leads',id,'undo');
    invalidateEntityCache('Leads');
    return success(updated);
  });
}

// ---------------------------------------------------------------------------
// Accounts / Contacts / Deals — generic CRUD builders
// ---------------------------------------------------------------------------
function _createGeneric(entity, payload, dupField) {
  return safeCall('create:'+entity, function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    const lock = LockService.getScriptLock(); lock.waitLock(15000);
    try {
      if (dupField) {
        const dup = findDuplicate(entity, payload);
        if (dup) return fail('Duplicate: ' + dup.field + ' already exists (id ' + dup.row.id + ')');
      }
      payload.createdAt = payload.createdAt || new Date().toISOString();
      payload.deleted = false;
      const saved = appendRowFromObject(entity, payload);
      logAudit(me.email,'create',entity,saved.id,payload);
      invalidateEntityCache(entity);
      return success(saved);
    } finally { lock.releaseLock(); }
  });
}
function _updateGeneric(entity, id, patch) {
  return safeCall('update:'+entity, function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    const before = getById(entity, id);
    const updated = updateRowById(entity,'id',id,patch);
    logAudit(me.email,'update',entity,id,{before:before,after:updated});
    invalidateEntityCache(entity);
    return success(updated);
  });
}
function _deleteGeneric(entity, id) {
  return safeCall('delete:'+entity, function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    const updated = updateRowById(entity,'id',id,{deleted:true});
    logAudit(me.email,'delete',entity,id,'soft-delete');
    invalidateEntityCache(entity);
    return success(updated);
  });
}
function _restoreGeneric(entity, id) {
  return safeCall('restore:'+entity, function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    const updated = updateRowById(entity,'id',id,{deleted:false});
    logAudit(me.email,'restore',entity,id,'undo');
    invalidateEntityCache(entity);
    return success(updated);
  });
}

function createAccount(p){ return _createGeneric('Accounts', p, false); }
function updateAccount(id,p){ return _updateGeneric('Accounts', id, p); }
function deleteAccount(id){ return _deleteGeneric('Accounts', id); }
function restoreAccount(id){ return _restoreGeneric('Accounts', id); }

function createContact(p){ return _createGeneric('Contacts', p, true); }
function updateContact(id,p){ return _updateGeneric('Contacts', id, p); }
function deleteContact(id){ return _deleteGeneric('Contacts', id); }
function restoreContact(id){ return _restoreGeneric('Contacts', id); }

function createDeal(p){ return _createGeneric('Deals', p, false); }
function updateDeal(id,p){ return _updateGeneric('Deals', id, p); }
function deleteDeal(id){ return _deleteGeneric('Deals', id); }
function restoreDeal(id){ return _restoreGeneric('Deals', id); }

// ---------------------------------------------------------------------------
// Activities
// ---------------------------------------------------------------------------
function addActivity(payload) {
  return safeCall('addActivity', function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    payload.createdAt = new Date().toISOString();
    payload.createdBy = me.email;
    const saved = appendRowFromObject('Activities', payload);
    // touch lastContactedDate if attached to a lead
    if (payload.leadId) {
      updateRowById('Leads','id',payload.leadId,{ lastContactedDate: payload.createdAt });
      invalidateEntityCache('Leads');
    }
    logAudit(me.email,'create','Activities',saved.id,{type:saved.type,subject:saved.subject});
    invalidateEntityCache('Activities');
    return success(saved);
  });
}

function getActivitiesFor(entity, id) {
  return safeCall('getActivitiesFor', function(){
    const key = { Leads:'leadId', Contacts:'contactId', Accounts:'accountId', Deals:'dealId' }[entity];
    if (!key) return fail('Bad entity');
    const rows = getSheetAsObjects('Activities')
      .filter(function(r){ return String(r[key]||'') === String(id); })
      .sort(function(a,b){ return (a.createdAt < b.createdAt) ? 1 : -1; });
    return success(rows);
  });
}

// ---------------------------------------------------------------------------
// Global search
// ---------------------------------------------------------------------------
function globalSearch(q) {
  return safeCall('globalSearch', function(){
    q = String(q||'').trim().toLowerCase();
    if (!q) return success({ leads:[], contacts:[], accounts:[] });
    function match(row, fields){
      for (let i=0;i<fields.length;i++) {
        if (String(row[fields[i]]||'').toLowerCase().indexOf(q) !== -1) return true;
      }
      return false;
    }
    const leads = getSheetAsObjects('Leads').filter(function(r){ return !r.deleted && match(r,['name','email','company','phone']); }).slice(0,25);
    const contacts = getSheetAsObjects('Contacts').filter(function(r){ return !r.deleted && match(r,['name','email','phone','title']); }).slice(0,25);
    const accounts = getSheetAsObjects('Accounts').filter(function(r){ return !r.deleted && match(r,['name','industry','website','phone']); }).slice(0,25);
    return success({ leads:leads, contacts:contacts, accounts:accounts });
  });
}

// ---------------------------------------------------------------------------
// Reports
// ---------------------------------------------------------------------------
function getReports() {
  return safeCall('getReports', function(){
    const cache = CacheService.getUserCache();
    const cached = cache.get('reports:v1');
    if (cached) return success(JSON.parse(cached));

    const leads = getSheetAsObjects('Leads').filter(function(r){return !r.deleted;});
    const users = getSheetAsObjects('Users');
    const funnel = STAGES.map(function(s){ return { stage:s, count: leads.filter(function(l){return l.stage===s;}).length }; });

    const bySource = {};
    SOURCES.forEach(function(s){ bySource[s] = { source:s, total:0, won:0 }; });
    leads.forEach(function(l){
      const s = l.source || 'Website';
      if (!bySource[s]) bySource[s] = { source:s, total:0, won:0 };
      bySource[s].total += 1;
      if (l.stage === 'Won') bySource[s].won += 1;
    });
    const sourceRoi = Object.keys(bySource).map(function(k){
      const b = bySource[k];
      b.winRate = b.total ? Math.round((b.won/b.total)*100) : 0;
      return b;
    }).sort(function(a,b){ return b.winRate - a.winRate; });

    const byRep = {};
    users.forEach(function(u){ byRep[u.email] = { name:u.name, email:u.email, total:0, won:0, lost:0 }; });
    leads.forEach(function(l){
      if (!byRep[l.ownerEmail]) byRep[l.ownerEmail] = { name:l.ownerEmail, email:l.ownerEmail, total:0, won:0, lost:0 };
      byRep[l.ownerEmail].total += 1;
      if (l.stage === 'Won')  byRep[l.ownerEmail].won  += 1;
      if (l.stage === 'Lost') byRep[l.ownerEmail].lost += 1;
    });
    const leaderboard = Object.keys(byRep).map(function(e){
      const r = byRep[e];
      r.winRate = r.total ? Math.round((r.won/r.total)*100) : 0;
      return r;
    }).sort(function(a,b){ return b.won - a.won || b.winRate - a.winRate; });

    const payload = { funnel:funnel, sourceRoi:sourceRoi, leaderboard:leaderboard };
    cache.put('reports:v1', JSON.stringify(payload), 60);
    return success(payload);
  });
}

// ---------------------------------------------------------------------------
// CSV Import
// ---------------------------------------------------------------------------
function importLeads(rows)   { return _importGeneric('Leads', rows, ['name']); }
function importContacts(rows){ return _importGeneric('Contacts', rows, ['name']); }
function importAccounts(rows){ return _importGeneric('Accounts', rows, ['name']); }

function _importGeneric(entity, rows, required) {
  return safeCall('import:'+entity, function(){
    const me = requireRole(['Sales Rep','Sales Manager','Admin']);
    if (!Array.isArray(rows) || rows.length === 0) return fail('No rows provided');
    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    try {
      const sh = getSheet(entity);
      const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
      const existing = getSheetAsObjects(entity);
      let startId = 1;
      if (existing.length) startId = Math.max.apply(null, existing.map(function(r){return Number(r.id)||0;})) + 1;

      const out = [];
      const skipped = [];
      const now = new Date().toISOString();
      const existingEmails = new Set(existing.map(function(r){ return String(r.email||'').toLowerCase(); }).filter(Boolean));
      const existingPhones = new Set(existing.map(function(r){ return String(r.phone||'').toLowerCase(); }).filter(Boolean));

      rows.forEach(function(raw, i){
        for (let k=0;k<required.length;k++) {
          if (!raw[required[k]]) { skipped.push({ row:i+2, reason:'missing '+required[k] }); return; }
        }
        if (raw.email && existingEmails.has(String(raw.email).toLowerCase())) { skipped.push({ row:i+2, reason:'duplicate email' }); return; }
        if (raw.phone && existingPhones.has(String(raw.phone).toLowerCase())) { skipped.push({ row:i+2, reason:'duplicate phone' }); return; }
        const obj = Object.assign({}, raw, { id:startId++, createdAt:now, deleted:false });
        if (entity === 'Leads') {
          obj.lastContactedDate = obj.lastContactedDate || now;
          obj.stage  = obj.stage  || 'New';
          obj.score  = Number(obj.score) || 50;
          obj.source = obj.source || 'Website';
          obj.ownerEmail = obj.ownerEmail || roundRobinAssign();
        }
        const rowArr = headers.map(function(h){
          const v = obj[h];
          return (v===undefined||v===null) ? '' : v;
        });
        out.push(rowArr);
        if (obj.email) existingEmails.add(String(obj.email).toLowerCase());
        if (obj.phone) existingPhones.add(String(obj.phone).toLowerCase());
      });

      if (out.length) {
        sh.getRange(sh.getLastRow()+1, 1, out.length, headers.length).setValues(out);
      }
      logAudit(me.email,'import',entity,'',{ imported:out.length, skipped:skipped.length });
      invalidateEntityCache(entity);
      return success({ imported: out.length, skipped: skipped.length, skippedReasons: skipped });
    } finally { lock.releaseLock(); }
  });
}

// ---------------------------------------------------------------------------
// Round-robin
// ---------------------------------------------------------------------------
function roundRobinAssign() {
  const users = getSheetAsObjects('Users').filter(function(u){ return u.active && u.role === 'Sales Rep'; });
  if (!users.length) return '';
  const p = Number(getConfig('roundRobin.pointer')) || 0;
  const pick = users[p % users.length];
  setConfig('roundRobin.pointer', String((p+1) % users.length));
  return pick.email;
}

// ---------------------------------------------------------------------------
// Duplicate detection
// ---------------------------------------------------------------------------
function findDuplicate(entity, payload) {
  const rows = getSheetAsObjects(entity).filter(function(r){ return !r.deleted; });
  if (payload.email) {
    const e = String(payload.email).toLowerCase();
    for (let i=0;i<rows.length;i++) if (String(rows[i].email||'').toLowerCase() === e) return { field:'email', row:rows[i] };
  }
  if (payload.phone) {
    const p = String(payload.phone).toLowerCase();
    for (let i=0;i<rows.length;i++) if (String(rows[i].phone||'').toLowerCase() === p) return { field:'phone', row:rows[i] };
  }
  return null;
}

function getById(entity, id) {
  const rows = getSheetAsObjects(entity);
  for (let i=0;i<rows.length;i++) if (String(rows[i].id) === String(id)) return rows[i];
  return null;
}
function getRecord(entity, id){ return safeCall('getRecord', function(){ return success(getById(entity, id)); }); }

// ---------------------------------------------------------------------------
// SLA / follow-up sweep — daily trigger
// ---------------------------------------------------------------------------
function slaSweep() {
  return safeCall('slaSweep', function(){
    const threshold = Number(getConfig('sla.followUpDays')) || 5;
    const mult      = Number(getConfig('sla.escalateMultiplier')) || 2;
    const users = getSheetAsObjects('Users');
    const managerEmail = (users.filter(function(u){return u.role==='Sales Manager';})[0]||{}).email;
    const now = new Date();
    const leads = getSheetAsObjects('Leads').filter(function(r){ return !r.deleted && ['Won','Lost'].indexOf(r.stage) === -1; });

    let notified = 0;
    leads.forEach(function(l){
      if (!l.lastContactedDate) return;
      const last = new Date(l.lastContactedDate);
      const days = Math.floor((now - last) / (1000*60*60*24));
      if (days < threshold) return;
      const cc = (days >= threshold*mult && managerEmail) ? managerEmail : '';
      try {
        MailApp.sendEmail({
          to: l.ownerEmail, cc: cc || undefined,
          subject: '[Orbit CRM] Follow-up overdue: ' + l.name + ' (' + days + ' days)',
          htmlBody: '<p>Lead <b>' + l.name + '</b> at <b>' + (l.company||'') + '</b> has had no activity for <b>' + days + '</b> days.</p>' +
                    '<p>Stage: ' + l.stage + ' • Score: ' + l.score + '</p>' +
                    '<p>Open Orbit CRM to log the next touch.</p>'
        });
        notified++;
      } catch(e) { logError('slaSweep.send', e); }
    });
    logAudit('system','sla-sweep','Leads','',{ notified:notified });
    return success({ notified:notified });
  });
}

function installSlaTrigger() {
  removeSlaTrigger();
  ScriptApp.newTrigger('slaSweep').timeBased().atHour(7).everyDays(1).create();
  SpreadsheetApp.getUi().alert('Daily SLA trigger installed (07:00).');
}
function removeSlaTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'slaSweep') ScriptApp.deleteTrigger(t);
  });
}

// ---------------------------------------------------------------------------
// Notifications (in-app)
// ---------------------------------------------------------------------------
function getNotifications() {
  return safeCall('getNotifications', function(){
    const me = getCurrentUser().data;
    const threshold = Number(getConfig('sla.followUpDays')) || 5;
    const now = new Date();
    const leads = getSheetAsObjects('Leads').filter(function(r){
      return !r.deleted && ['Won','Lost'].indexOf(r.stage) === -1 && r.ownerEmail === me.email;
    });
    const out = [];
    leads.forEach(function(l){
      if (!l.lastContactedDate) return;
      const days = Math.floor((now - new Date(l.lastContactedDate)) / (1000*60*60*24));
      if (days >= threshold) {
        out.push({
          id: 'sla-' + l.id,
          type: (days >= threshold*2) ? 'critical' : 'warning',
          title: 'Follow up: ' + l.name,
          body: (l.company||'') + ' • ' + days + ' days since last touch',
          leadId: l.id,
          ts: new Date().toISOString()
        });
      }
    });
    return success(out);
  });
}

// ---------------------------------------------------------------------------
// User preferences (theme, page size)
// ---------------------------------------------------------------------------
function savePreferences(prefs) {
  return safeCall('savePreferences', function(){
    const me = getCurrentUser().data;
    const sh = getSheet('UserPreferences');
    const rows = getSheetAsObjects('UserPreferences');
    const json = JSON.stringify(prefs||{});
    for (let i=0;i<rows.length;i++) {
      if (rows[i].email === me.email) { sh.getRange(i+2,2).setValue(json); return success(true); }
    }
    sh.appendRow([me.email, json]);
    return success(true);
  });
}
function loadPreferences() {
  return safeCall('loadPreferences', function(){
    const me = getCurrentUser().data;
    const rows = getSheetAsObjects('UserPreferences');
    for (let i=0;i<rows.length;i++) {
      if (rows[i].email === me.email) {
        try { return success(JSON.parse(rows[i].prefs||'{}')); } catch(e){ return success({}); }
      }
    }
    return success({});
  });
}
