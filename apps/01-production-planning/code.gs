/**
 * ============================================================================
 *  PRODUCTION PLANNING & SHOP-FLOOR EXECUTION SYSTEM — Google Apps Script
 * ============================================================================
 *  DEPLOYMENT
 *  ----------
 *  1. Open Google Sheets → Extensions → Apps Script.
 *  2. Paste this file as code.gs and index.html as an HTML file named "index".
 *  3. Reload the sheet; use menu "Production System" → "Initialize System".
 *  4. Deploy → New deployment → Web app → Execute as: Me, Access: domain/anyone.
 *
 *  SHEETS (created by setup)
 *  Machines, Products, ProductionPlans, ProductionEntries, DowntimeEvents,
 *  Shifts, OEEResults, Users, Config, AuditLog, JobRuns
 * ============================================================================
 */

var SHEETS = {
  Machines: ['code','name','section','ratedCapacityHr','status'],
  Products: ['sku','name','stdCycleSec','stdScrapPct'],
  ProductionPlans: ['planId','date','shift','machine','sku','plannedQty','priority','dispatchDate','status','createdBy','createdAt'],
  ProductionEntries: ['entryId','planId','hourSlot','qtyOk','qtyRejected','qtyRework','operator','remarks','lateFlag','loggedAt'],
  DowntimeEvents: ['dtId','machine','startTime','endTime','category','subReason','minutes','loggedBy','open'],
  Shifts: ['name','startHH','endHH'],
  OEEResults: ['date','shift','machine','availability','performance','quality','oee','plannedMin','downtimeMin','totalQty','okQty'],
  Users: ['email','name','role','department','active','notify'],
  Config: ['key','value','description'],
  AuditLog: ['ts','user','action','entity','recordId','oldValue','newValue','screen'],
  JobRuns: ['ts','job','durationMs','outcome','detail']
};

var ROLES = ['Admin','Planner','Section Head','Operator','Plant Head'];

/* ------------------------------ core plumbing ----------------------------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Production System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Production System')
    .addItem('Initialize System', 'setup')
    .addItem('Reset Demo Data', 'resetDemoData')
    .addItem('Run OEE Engine Now', 'runOeeEngine')
    .addToUi();
}

function ok(data) { return { ok: true, data: data, error: null }; }
function fail(msg) { return { ok: false, data: null, error: String(msg) }; }

function safe(fn) {
  try { return fn(); }
  catch (e) { logJob('safeCall', 0, 'error', e.message); return fail(e.message); }
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet(name) {
  var sh = ss().getSheetByName(name);
  if (!sh) throw new Error('Sheet missing: ' + name + '. Run Initialize System.');
  return sh;
}

function headerMap(sh) {
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var m = {};
  h.forEach(function (c, i) { m[c] = i; });
  return m;
}

function rows(name) {
  var sh = sheet(name);
  if (sh.getLastRow() < 2) return [];
  var hm = headerMap(sh);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  return vals.map(function (r, idx) {
    var o = { _row: idx + 2 };
    Object.keys(hm).forEach(function (k) { o[k] = r[hm[k]]; });
    return o;
  });
}

function appendObj(name, obj) {
  var sh = sheet(name);
  var hm = headerMap(sh);
  var r = new Array(sh.getLastColumn()).fill('');
  Object.keys(obj).forEach(function (k) { if (hm[k] !== undefined) r[hm[k]] = obj[k]; });
  sh.appendRow(r);
}

function updateByKey(name, keyCol, keyVal, fields) {
  var sh = sheet(name);
  var hm = headerMap(sh);
  var all = rows(name);
  var found = null;
  for (var i = 0; i < all.length; i++) if (String(all[i][keyCol]) === String(keyVal)) { found = all[i]; break; }
  if (!found) throw new Error(name + ' record not found: ' + keyVal);
  var old = {};
  Object.keys(fields).forEach(function (k) {
    old[k] = found[k];
    sh.getRange(found._row, hm[k] + 1).setValue(fields[k]);
  });
  return old;
}

function getConfig(key) {
  var c = rows('Config').filter(function (r) { return r.key === key; })[0];
  return c ? c.value : null;
}
function getConfigNum(key) { return Number(getConfig(key)); }

function nextId(prefix) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var key = 'SEQ_' + prefix;
    var cur = getConfig(key);
    var n = cur ? Number(cur) + 1 : 1;
    var found = rows('Config').filter(function (r) { return r.key === key; })[0];
    if (found) updateByKey('Config', 'key', key, { value: n });
    else appendObj('Config', { key: key, value: n, description: 'ID counter' });
    var yr = new Date().getFullYear();
    return prefix + '-' + yr + '-' + ('0000' + n).slice(-4);
  } finally { lock.releaseLock(); }
}

function currentUser() {
  var email = Session.getActiveUser().getEmail() || 'demo.operator@plant.local';
  var u = rows('Users').filter(function (r) { return r.email === email && r.active === true; })[0];
  if (!u) {
    // unknown users get read-only Plant Head visibility for demo safety
    return { email: email, name: email.split('@')[0], role: 'Plant Head', department: '-', notify: false };
  }
  return u;
}

function requireRole(allowed) {
  var u = currentUser();
  if (allowed.indexOf(u.role) === -1 && u.role !== 'Admin') throw new Error('Permission denied for role: ' + u.role);
  return u;
}

function audit(action, entity, id, oldV, newV, screen) {
  appendObj('AuditLog', {
    ts: new Date(), user: currentUser().email, action: action, entity: entity,
    recordId: id, oldValue: JSON.stringify(oldV || {}), newValue: JSON.stringify(newV || {}), screen: screen || ''
  });
}

function logJob(job, durationMs, outcome, detail) {
  try { appendObj('JobRuns', { ts: new Date(), job: job, durationMs: durationMs, outcome: outcome, detail: detail || '' }); } catch (e) {}
}

/* --------------------------------- setup ---------------------------------- */

function setup() {
  var s = ss();
  Object.keys(SHEETS).forEach(function (name) {
    var sh = s.getSheetByName(name);
    if (!sh) {
      sh = s.insertSheet(name);
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]])
        .setFontWeight('bold').setBackground('#f1f1ef');
      sh.setFrozenRows(1);
    }
  });
  seedConfig();
  seedDemoData();
  installTriggers();
  return ok('Setup complete');
}

function seedConfig() {
  var defaults = [
    ['APP_VERSION', 'v1.0', 'Version shown on About page'],
    ['BUILD_DATE', new Date().toDateString(), 'Build date'],
    ['BEHIND_PLAN_PCT', '10', 'Board: amber when behind plan by more than this %'],
    ['DOWNTIME_EMAIL_MIN', '30', 'Downtime > this (min) emails section head'],
    ['DOWNTIME_ESCALATE_MIN', '60', 'Downtime > this (min) escalates to plant head'],
    ['LATE_ENTRY_HOURS', '2', 'Entries logged later than this are flagged'],
    ['DISPATCH_RISK_TOLERANCE_PCT', '15', 'SKU behind cumulative plan by > this % = commitment at risk'],
    ['CAPACITY_OVERLOAD_PCT', '100', 'Plan slot flagged when load exceeds this % of rated capacity'],
    ['SECTION_HEAD_EMAIL', 'sectionhead@plant.local', 'Downtime alert recipient'],
    ['PLANT_HEAD_EMAIL', 'planthead@plant.local', 'Escalation + digest recipient'],
    ['DT_REASONS', JSON.stringify({
      breakdown: ['Mechanical', 'Electrical', 'Hydraulic'],
      changeover: ['Die change', 'SKU change', 'Setup adjust'],
      'no-material': ['RM shortage', 'WIP waiting'],
      'no-operator': ['Absent', 'Break overlap'],
      power: ['Grid failure', 'DG switchover'],
      'quality-hold': ['NCR hold', 'First-piece pending']
    }), 'Two-level downtime reason tree']
  ];
  var existing = rows('Config').map(function (r) { return r.key; });
  defaults.forEach(function (d) {
    if (existing.indexOf(d[0]) === -1) appendObj('Config', { key: d[0], value: d[1], description: d[2] });
  });
}

function resetDemoData() {
  ['Machines','Products','ProductionPlans','ProductionEntries','DowntimeEvents','Shifts','OEEResults','Users'].forEach(function (n) {
    var sh = sheet(n);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  });
  seedDemoData();
  return ok('Demo data reset');
}

function seedDemoData() {
  if (rows('Machines').length > 0) return;
  var machines = [
    ['M-01','Press 250T','Press Shop',120,'running'],['M-02','Press 400T','Press Shop',90,'running'],
    ['M-03','CNC Lathe 1','Machining',60,'down'],['M-04','CNC Lathe 2','Machining',60,'running'],
    ['M-05','Welding Cell A','Fabrication',45,'running'],['M-06','Welding Cell B','Fabrication',45,'idle'],
    ['M-07','Paint Booth','Finishing',200,'running'],['M-08','Assembly Line 1','Assembly',80,'running']
  ];
  machines.forEach(function (m) { appendObj('Machines', { code: m[0], name: m[1], section: m[2], ratedCapacityHr: m[3], status: m[4] }); });

  var products = [
    ['SKU-A100','Bracket A100',28,2],['SKU-B200','Housing B200',45,3],['SKU-C300','Shaft C300',60,1.5],
    ['SKU-D400','Panel D400',18,2.5],['SKU-E500','Frame E500',90,4]
  ];
  products.forEach(function (p) { appendObj('Products', { sku: p[0], name: p[1], stdCycleSec: p[2], stdScrapPct: p[3] }); });

  [['A',6,14],['B',14,22],['C',22,6]].forEach(function (s) { appendObj('Shifts', { name: s[0], startHH: s[1], endHH: s[2] }); });

  [['admin@plant.local','Asha Admin','Admin','IT',true,true],
   ['planner@plant.local','Prakash Planner','Planner','Planning',true,true],
   ['sectionhead@plant.local','Suresh Section','Section Head','Press Shop',true,true],
   ['op1@plant.local','Omkar Operator','Operator','Press Shop',true,false],
   ['planthead@plant.local','Priya PlantHead','Plant Head','Management',true,true]
  ].forEach(function (u) { appendObj('Users', { email: u[0], name: u[1], role: u[2], department: u[3], active: u[4], notify: u[5] }); });

  // 30 days of plans + entries + downtime so OEE trends render
  var today = new Date();
  var skus = ['SKU-A100','SKU-B200','SKU-C300','SKU-D400','SKU-E500'];
  var pn = 0;
  for (var d = 30; d >= 0; d--) {
    var date = new Date(today.getTime() - d * 86400000);
    var ds = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    machines.forEach(function (m, mi) {
      if (mi === 5 && d < 3) return; // M-06 no plan recently → grey tile
      var sku = skus[(mi + d) % skus.length];
      var planned = Math.round(m[3] * 8 * (mi === 1 && d === 0 ? 1.4 : 0.85)); // one overloaded slot today
      pn++;
      var planId = 'PL-' + ds.replace(/-/g, '') + '-' + m[0];
      appendObj('ProductionPlans', {
        planId: planId, date: ds, shift: 'A', machine: m[0], sku: sku, plannedQty: planned,
        priority: (d < 5 && mi < 2) ? 'HIGH' : 'NORMAL',
        dispatchDate: Utilities.formatDate(new Date(date.getTime() + 3 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        status: d === 0 ? 'active' : 'closed', createdBy: 'planner@plant.local', createdAt: date
      });
      // entries: M-03 had a 90-min breakdown shift (missed plan); M-01 beat plan once
      var perf = mi === 2 ? 0.62 : (mi === 0 && d === 1 ? 1.08 : 0.8 + ((mi + d) % 5) * 0.04);
      var hours = d === 0 ? Math.min(8, new Date().getHours() - 6) : 8;
      for (var h = 0; h < Math.max(0, hours); h++) {
        var okQ = Math.round((planned / 8) * perf);
        appendObj('ProductionEntries', {
          entryId: 'EN-' + ds.replace(/-/g, '') + '-' + m[0] + '-' + h, planId: planId,
          hourSlot: (6 + h) + ':00', qtyOk: okQ, qtyRejected: Math.round(okQ * 0.02),
          qtyRework: Math.round(okQ * 0.01), operator: 'op1@plant.local', remarks: '',
          lateFlag: (h === 3 && d === 2), loggedAt: new Date(date.getTime() + (7 + h) * 3600000)
        });
      }
      if (mi === 2 && d % 4 === 1) {
        appendObj('DowntimeEvents', {
          dtId: 'DT-' + ds.replace(/-/g, '') + '-' + m[0], machine: m[0],
          startTime: new Date(date.getTime() + 9 * 3600000), endTime: new Date(date.getTime() + 10.5 * 3600000),
          category: 'breakdown', subReason: 'Mechanical', minutes: 90, loggedBy: 'op1@plant.local', open: false
        });
      }
      if ((mi === 0 || mi === 4) && d % 6 === 2) {
        appendObj('DowntimeEvents', {
          dtId: 'DT-' + ds.replace(/-/g, '') + '-' + m[0] + 'C', machine: m[0],
          startTime: new Date(date.getTime() + 12 * 3600000), endTime: new Date(date.getTime() + 12.5 * 3600000),
          category: 'changeover', subReason: 'SKU change', minutes: 30, loggedBy: 'op1@plant.local', open: false
        });
      }
    });
  }
  // one machine currently DOWN (open downtime)
  appendObj('DowntimeEvents', {
    dtId: 'DT-OPEN-M03', machine: 'M-03', startTime: new Date(Date.now() - 47 * 60000),
    endTime: '', category: 'breakdown', subReason: 'Electrical', minutes: '', loggedBy: 'op1@plant.local', open: true
  });
  runOeeEngine();
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runOeeEngine').timeBased().everyDays(1).atHour(1).create();
  ScriptApp.newTrigger('sendMorningDigest').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('checkOpenDowntime').timeBased().everyMinutes(30).create();
}

/* ------------------------------- bootstrap -------------------------------- */

function bootstrap() {
  return safe(function () {
    var me = currentUser();
    var cfg = {};
    rows('Config').forEach(function (r) { cfg[r.key] = r.value; });
    return ok({
      me: me,
      config: cfg,
      machines: rows('Machines'),
      products: rows('Products'),
      shifts: rows('Shifts'),
      dtReasons: JSON.parse(getConfig('DT_REASONS') || '{}')
    });
  });
}

/* ---------------------------- plans & capacity ----------------------------- */

function savePlan(p) {
  return safe(function () {
    requireRole(['Planner']);
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      var m = rows('Machines').filter(function (x) { return x.code === p.machine; })[0];
      if (!m) return fail('Unknown machine');
      var loadPct = (Number(p.plannedQty) / (m.ratedCapacityHr * 8)) * 100;
      var over = loadPct > getConfigNum('CAPACITY_OVERLOAD_PCT');
      var planId = 'PL-' + String(p.date).replace(/-/g, '') + '-' + p.machine + '-' + Math.floor(Math.random() * 90 + 10);
      appendObj('ProductionPlans', {
        planId: planId, date: p.date, shift: p.shift, machine: p.machine, sku: p.sku,
        plannedQty: Number(p.plannedQty), priority: p.priority || 'NORMAL',
        dispatchDate: p.dispatchDate || '', status: 'active',
        createdBy: currentUser().email, createdAt: new Date()
      });
      audit('CREATE', 'ProductionPlans', planId, null, p, 'Plan Builder');
      return ok({ planId: planId, loadPct: Math.round(loadPct), overloaded: over });
    } finally { lock.releaseLock(); }
  });
}

function getPlans(dateStr) {
  return safe(function () {
    var plans = rows('ProductionPlans').filter(function (p) { return String(p.date) === dateStr || Utilities.formatDate(new Date(p.date), Session.getScriptTimeZone(), 'yyyy-MM-dd') === dateStr; });
    var machines = rows('Machines');
    plans.forEach(function (p) {
      var m = machines.filter(function (x) { return x.code === p.machine; })[0];
      p.loadPct = m ? Math.round((p.plannedQty / (m.ratedCapacityHr * 8)) * 100) : 0;
      p.overloaded = p.loadPct > getConfigNum('CAPACITY_OVERLOAD_PCT');
    });
    return ok(plans);
  });
}

/* --------------------------- production logging ---------------------------- */

function logProduction(e) {
  return safe(function () {
    requireRole(['Operator', 'Section Head']);
    var openDt = rows('DowntimeEvents').filter(function (d) { return d.machine === e.machine && d.open === true; });
    if (openDt.length) return fail('Machine ' + e.machine + ' has an OPEN downtime event (' + openDt[0].dtId + '). Close it before logging production.');
    var plan = rows('ProductionPlans').filter(function (p) { return p.planId === e.planId; })[0];
    if (!plan) return fail('Plan not found');
    if (Number(e.qtyOk) < 0 || Number(e.qtyRejected) < 0 || Number(e.qtyRework) < 0) return fail('Quantities cannot be negative');
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      var slotHour = Number(String(e.hourSlot).split(':')[0]);
      var late = (new Date().getHours() - slotHour) > getConfigNum('LATE_ENTRY_HOURS');
      var id = 'EN-' + Date.now();
      appendObj('ProductionEntries', {
        entryId: id, planId: e.planId, hourSlot: e.hourSlot,
        qtyOk: Number(e.qtyOk), qtyRejected: Number(e.qtyRejected || 0), qtyRework: Number(e.qtyRework || 0),
        operator: currentUser().email, remarks: e.remarks || '', lateFlag: late, loggedAt: new Date()
      });
      audit('CREATE', 'ProductionEntries', id, null, e, 'Hourly Logging');
      return ok({ entryId: id, lateFlag: late });
    } finally { lock.releaseLock(); }
  });
}

/* -------------------------------- downtime --------------------------------- */

function startDowntime(d) {
  return safe(function () {
    requireRole(['Operator', 'Section Head']);
    var open = rows('DowntimeEvents').filter(function (x) { return x.machine === d.machine && x.open === true; });
    if (open.length) return fail('Machine already has an open downtime event.');
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      var id = nextId('DT');
      appendObj('DowntimeEvents', {
        dtId: id, machine: d.machine, startTime: new Date(), endTime: '',
        category: d.category, subReason: d.subReason, minutes: '', loggedBy: currentUser().email, open: true
      });
      updateByKey('Machines', 'code', d.machine, { status: 'down' });
      audit('CREATE', 'DowntimeEvents', id, null, d, 'Downtime');
      return ok({ dtId: id });
    } finally { lock.releaseLock(); }
  });
}

function endDowntime(dtId) {
  return safe(function () {
    requireRole(['Operator', 'Section Head']);
    var lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      var dt = rows('DowntimeEvents').filter(function (x) { return x.dtId === dtId; })[0];
      if (!dt || dt.open !== true) return fail('Open downtime event not found');
      var mins = Math.round((Date.now() - new Date(dt.startTime).getTime()) / 60000);
      var old = updateByKey('DowntimeEvents', 'dtId', dtId, { endTime: new Date(), minutes: mins, open: false });
      updateByKey('Machines', 'code', dt.machine, { status: 'running' });
      audit('UPDATE', 'DowntimeEvents', dtId, old, { minutes: mins }, 'Downtime');
      return ok({ minutes: mins });
    } finally { lock.releaseLock(); }
  });
}

function endDowntimeByMachine(machine) {
  return safe(function () {
    var dt = rows('DowntimeEvents').filter(function (x) { return x.machine === machine && x.open === true; })[0];
    if (!dt) return fail('No open downtime for ' + machine);
    return endDowntime(dt.dtId);
  });
}

function checkOpenDowntime() {
  var t0 = Date.now();
  try {
    var alertMin = getConfigNum('DOWNTIME_EMAIL_MIN');
    var escMin = getConfigNum('DOWNTIME_ESCALATE_MIN');
    rows('DowntimeEvents').filter(function (d) { return d.open === true; }).forEach(function (d) {
      var mins = Math.round((Date.now() - new Date(d.startTime).getTime()) / 60000);
      if (mins > escMin) notify(getConfig('PLANT_HEAD_EMAIL'), 'ESCALATION: ' + d.machine + ' down ' + mins + ' min',
        'Machine ' + d.machine + ' has been down for ' + mins + ' minutes (' + d.category + ' / ' + d.subReason + '). Immediate attention required.');
      else if (mins > alertMin) notify(getConfig('SECTION_HEAD_EMAIL'), 'Downtime alert: ' + d.machine + ' down ' + mins + ' min',
        'Machine ' + d.machine + ' down for ' + mins + ' minutes (' + d.category + ' / ' + d.subReason + ').');
    });
    logJob('checkOpenDowntime', Date.now() - t0, 'ok', '');
  } catch (e) { logJob('checkOpenDowntime', Date.now() - t0, 'error', e.message); }
}

function notify(to, subject, body) {
  try {
    if (to) MailApp.sendEmail({ to: to, subject: '[Production System] ' + subject, htmlBody:
      '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1a18"><p>' + body + '</p>' +
      '<p style="color:#6b6b66;font-size:12px">Production Planning & Shop-Floor Execution System</p></div>' });
  } catch (e) { logJob('notify', 0, 'error', e.message); }
}

/* ------------------------------- plant board -------------------------------- */

function getPlantBoard() {
  return safe(function () {
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var machines = rows('Machines');
    var plans = rows('ProductionPlans').filter(function (p) { return fmtD(p.date) === todayStr; });
    var entries = rows('ProductionEntries');
    var openDt = rows('DowntimeEvents').filter(function (d) { return d.open === true; });
    var behindPct = getConfigNum('BEHIND_PLAN_PCT');
    var tiles = machines.map(function (m) {
      var plan = plans.filter(function (p) { return p.machine === m.code; })[0];
      var dt = openDt.filter(function (d) { return d.machine === m.code; })[0];
      var tile = { code: m.code, name: m.name, section: m.section, state: 'grey', planned: 0, actual: 0, pct: 0, downMin: 0 };
      if (dt) {
        tile.state = 'red';
        tile.downMin = Math.round((Date.now() - new Date(dt.startTime).getTime()) / 60000);
        tile.dtReason = dt.category + ' / ' + dt.subReason;
      }
      if (plan) {
        tile.planned = Number(plan.plannedQty);
        var hoursElapsed = Math.max(1, Math.min(8, new Date().getHours() - 6));
        var expected = tile.planned * (hoursElapsed / 8);
        tile.actual = entries.filter(function (e) { return e.planId === plan.planId; })
          .reduce(function (s, e) { return s + Number(e.qtyOk || 0); }, 0);
        tile.pct = tile.planned ? Math.round((tile.actual / tile.planned) * 100) : 0;
        if (!dt) tile.state = (tile.actual < expected * (1 - behindPct / 100)) ? 'amber' : 'green';
      }
      return tile;
    });
    var sections = {};
    tiles.forEach(function (t) {
      if (!sections[t.section]) sections[t.section] = { planned: 0, actual: 0 };
      sections[t.section].planned += t.planned; sections[t.section].actual += t.actual;
    });
    return ok({ tiles: tiles, sections: sections, asOf: new Date() });
  });
}

function fmtD(d) {
  if (!d) return '';
  if (typeof d === 'string' && d.match(/^\d{4}-\d{2}-\d{2}/)) return d.slice(0, 10);
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* --------------------------------- OEE engine ------------------------------- */

function runOeeEngine() {
  var t0 = Date.now();
  try {
    var sh = sheet('OEEResults');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    var plans = rows('ProductionPlans');
    var entries = rows('ProductionEntries');
    var dts = rows('DowntimeEvents').filter(function (d) { return d.open !== true; });
    var products = rows('Products');
    var machines = rows('Machines');
    var out = [];
    plans.forEach(function (p) {
      var pe = entries.filter(function (e) { return e.planId === p.planId; });
      if (!pe.length) return;
      var plannedMin = 8 * 60;
      var dMin = dts.filter(function (d) { return d.machine === p.machine && fmtD(d.startTime) === fmtD(p.date); })
        .reduce(function (s, d) { return s + Number(d.minutes || 0); }, 0);
      var runMin = Math.max(1, plannedMin - dMin);
      var totalQty = pe.reduce(function (s, e) { return s + Number(e.qtyOk || 0) + Number(e.qtyRejected || 0) + Number(e.qtyRework || 0); }, 0);
      var okQty = pe.reduce(function (s, e) { return s + Number(e.qtyOk || 0); }, 0);
      var prod = products.filter(function (x) { return x.sku === p.sku; })[0];
      var cycleSec = prod ? Number(prod.stdCycleSec) : 30;
      var availability = runMin / plannedMin;
      var performance = Math.min(1.2, (totalQty * cycleSec / 60) / runMin);
      var quality = totalQty ? okQty / totalQty : 0;
      out.push([fmtD(p.date), p.shift, p.machine,
        round2(availability), round2(performance), round2(quality),
        round2(availability * performance * quality), plannedMin, dMin, totalQty, okQty]);
    });
    if (out.length) sh.getRange(2, 1, out.length, out[0].length).setValues(out);
    logJob('runOeeEngine', Date.now() - t0, 'ok', out.length + ' rows');
  } catch (e) { logJob('runOeeEngine', Date.now() - t0, 'error', e.message); }
}

function round2(n) { return Math.round(n * 10000) / 10000; }

function getOeeTrend() {
  return safe(function () {
    var res = rows('OEEResults');
    var byMachine = {};
    res.forEach(function (r) {
      if (!byMachine[r.machine]) byMachine[r.machine] = [];
      byMachine[r.machine].push({ date: fmtD(r.date), oee: Number(r.oee) });
    });
    var avg = Object.keys(byMachine).map(function (m) {
      var arr = byMachine[m];
      return { machine: m, avgOee: round2(arr.reduce(function (s, x) { return s + x.oee; }, 0) / arr.length), points: arr.slice(-30) };
    }).sort(function (a, b) { return a.avgOee - b.avgOee; });
    return ok({ worst5: avg.slice(0, 5), all: avg });
  });
}

/* ---------------------------- variance & dispatch --------------------------- */

function getVarianceReport(dateStr) {
  return safe(function () {
    var plans = rows('ProductionPlans').filter(function (p) { return fmtD(p.date) === dateStr; });
    var entries = rows('ProductionEntries');
    var dts = rows('DowntimeEvents');
    var rep = plans.map(function (p) {
      var pe = entries.filter(function (e) { return e.planId === p.planId; });
      var produced = pe.reduce(function (s, e) { return s + Number(e.qtyOk || 0); }, 0);
      var rejected = pe.reduce(function (s, e) { return s + Number(e.qtyRejected || 0); }, 0);
      var dtByCat = {};
      dts.filter(function (d) { return d.machine === p.machine && fmtD(d.startTime) === dateStr && d.open !== true; })
        .forEach(function (d) { dtByCat[d.category] = (dtByCat[d.category] || 0) + Number(d.minutes || 0); });
      return {
        machine: p.machine, shift: p.shift, sku: p.sku, planned: Number(p.plannedQty),
        produced: produced, shortfall: Number(p.plannedQty) - produced,
        rejPct: produced ? round2(rejected / (produced + rejected)) * 100 : 0, dtByCat: dtByCat
      };
    });
    return ok(rep);
  });
}

function getDispatchRisk() {
  return safe(function () {
    var tol = getConfigNum('DISPATCH_RISK_TOLERANCE_PCT');
    var plans = rows('ProductionPlans');
    var entries = rows('ProductionEntries');
    var bySku = {};
    plans.forEach(function (p) {
      if (!bySku[p.sku]) bySku[p.sku] = { planned: 0, actual: 0, nearestDispatch: null, priority: 'NORMAL' };
      bySku[p.sku].planned += Number(p.plannedQty);
      bySku[p.sku].actual += entries.filter(function (e) { return e.planId === p.planId; })
        .reduce(function (s, e) { return s + Number(e.qtyOk || 0); }, 0);
      if (p.priority === 'HIGH') bySku[p.sku].priority = 'HIGH';
      if (p.dispatchDate && (!bySku[p.sku].nearestDispatch || fmtD(p.dispatchDate) < bySku[p.sku].nearestDispatch))
        bySku[p.sku].nearestDispatch = fmtD(p.dispatchDate);
    });
    var risk = [];
    Object.keys(bySku).forEach(function (sku) {
      var s = bySku[sku];
      var gapPct = s.planned ? ((s.planned - s.actual) / s.planned) * 100 : 0;
      if (gapPct > tol) risk.push({ sku: sku, planned: s.planned, actual: s.actual, gapPct: Math.round(gapPct), dispatchDate: s.nearestDispatch, priority: s.priority });
    });
    return ok(risk.sort(function (a, b) { return b.gapPct - a.gapPct; }));
  });
}

function getChangeoverReport() {
  return safe(function () {
    var co = rows('DowntimeEvents').filter(function (d) { return d.category === 'changeover' && d.open !== true; });
    var byMachine = {};
    co.forEach(function (d) {
      if (!byMachine[d.machine]) byMachine[d.machine] = { count: 0, totalMin: 0 };
      byMachine[d.machine].count++; byMachine[d.machine].totalMin += Number(d.minutes || 0);
    });
    return ok(Object.keys(byMachine).map(function (m) {
      return { machine: m, count: byMachine[m].count, avgMin: Math.round(byMachine[m].totalMin / byMachine[m].count), totalMin: byMachine[m].totalMin };
    }));
  });
}

/* --------------------------------- digest ----------------------------------- */

function sendMorningDigest() {
  var t0 = Date.now();
  try {
    var yesterday = Utilities.formatDate(new Date(Date.now() - 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var vr = getVarianceReport(yesterday).data || [];
    var risk = getDispatchRisk().data || [];
    var html = '<div style="font-family:Arial;font-size:14px"><h3>Morning production digest — ' + yesterday + '</h3>' +
      '<table border="1" cellpadding="6" style="border-collapse:collapse;font-size:13px"><tr><th>Machine</th><th>SKU</th><th>Planned</th><th>Produced</th><th>Shortfall</th></tr>' +
      vr.map(function (r) { return '<tr><td>' + r.machine + '</td><td>' + r.sku + '</td><td align="right">' + r.planned + '</td><td align="right">' + r.produced + '</td><td align="right">' + r.shortfall + '</td></tr>'; }).join('') +
      '</table><h4>Commitments at risk</h4><ul>' +
      (risk.length ? risk.map(function (r) { return '<li>' + r.sku + ': ' + r.gapPct + '% behind, dispatch ' + r.dispatchDate + '</li>'; }).join('') : '<li>None</li>') +
      '</ul></div>';
    notifyRaw(getConfig('PLANT_HEAD_EMAIL'), 'Morning production digest ' + yesterday, html);
    logJob('sendMorningDigest', Date.now() - t0, 'ok', '');
  } catch (e) { logJob('sendMorningDigest', Date.now() - t0, 'error', e.message); }
}

function notifyRaw(to, subject, html) {
  try { if (to) MailApp.sendEmail({ to: to, subject: '[Production System] ' + subject, htmlBody: html }); }
  catch (e) { logJob('notifyRaw', 0, 'error', e.message); }
}

/* ---------------------------------- about ------------------------------------ */

function getAboutData() {
  return safe(function () {
    var cfg = {};
    rows('Config').forEach(function (r) { cfg[r.key] = r.value; });
    return ok({
      version: cfg.APP_VERSION, buildDate: cfg.BUILD_DATE, config: cfg,
      sheets: Object.keys(SHEETS).map(function (n) { return { name: n, columns: SHEETS[n] }; }),
      roles: ROLES, lastVerified: new Date()
    });
  });
}
