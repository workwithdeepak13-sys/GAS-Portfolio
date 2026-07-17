/*******************************************************************************
 * EXPENSE & REIMBURSEMENT TRACKER — Backend (ExpenseCode.gs)
 * ---------------------------------------------------------------------------
 * Google Apps Script backend for the expense/reimbursement claim workflow:
 *   • Role-based login (Admin / Finance / Manager / Employee)
 *   • Employees raise claims with category, amount, receipts, description
 *   • Manager / Admin / Finance can Approve / Reject with remarks
 *   • Every state change → automated email via MailApp/GmailApp
 *   • Categories master (Travel, Food, Accommodation, Office, Client, Other)
 *   • Full audit trail (activity log) + advance requests
 *   • initializeSheets() creates every sheet AND seeds rich demo data
 *
 * DEPLOY
 *  1. Paste this as `ExpenseCode.gs` in Apps Script
 *  2. Add HTML file named **`expense`** with expense.html contents
 *  3. Run `initializeSheets` once (grant Gmail + Sheets perms)
 *  4. Deploy → Web app → Execute as *Me*, Access *Anyone*
 *  5. To disable emails temporarily: Project Settings → Script Properties →
 *     set `EMAIL_ENABLED = false`
 ******************************************************************************/

const EXP_CFG = {
  CACHE_TTL_SEC: 60, LOCK_TIMEOUT_MS: 15000, DEFAULT_APPROVER_EMAIL: 'admin@expense.com',
  SHEETS: {
    USERS:'ExpUsers', CATEGORIES:'ExpCategories', CLAIMS:'Claims',
    ADVANCES:'Advances', APPROVALS:'ExpApprovals', EMAIL_LOG:'EmailLog', ACTIVITY_LOG:'ExpActivityLog'
  }
};
const EXP_SCHEMAS = {
  ExpUsers:['id','email','password','name','role','department','manager','managerEmail','active','createdDate'],
  ExpCategories:['id','code','name','maxLimit','requiresReceipt','description','active','createdDate'],
  Claims:['id','claimNumber','employeeId','employeeName','employeeEmail','category','categoryId',
    'amount','currency','claimDate','expenseDate','description','receipts','paymentMode',
    'status','approver','approverEmail','approvedAmount','approvedDate','remarks',
    'reimbursedDate','reimbursedRef','createdDate','updatedDate'],
  Advances:['id','advanceNumber','employeeId','employeeName','employeeEmail','purpose',
    'requestedAmount','requiredDate','status','approvedAmount','approvedBy','approvedDate',
    'settledDate','remarks','createdDate','updatedDate'],
  ExpApprovals:['id','entityType','entityId','requestedBy','requestedDate','approver',
    'status','remarks','decidedDate','createdDate'],
  EmailLog:['ts','to','subject','type','entityId','status','error'],
  ExpActivityLog:['ts','user','action','entity','entityId','payload']
};

// ─── WEB ENTRY ──────────────────────────────────────────────────────────────
function doGet(e){
  return HtmlService.createHtmlOutputFromFile('expense')
    .setTitle('Expense & Reimbursement').addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ─── INIT + SEED ────────────────────────────────────────────────────────────
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock(); lock.tryLock(EXP_CFG.LOCK_TIMEOUT_MS);
  try {
    Object.entries(EXP_SCHEMAS).forEach(([name, headers]) => {
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      exp_ensureHeaders_(sh, headers);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#0F172A').setFontColor('#FFFFFF');
    });
    const s1 = ss.getSheetByName('Sheet1');
    if (s1 && ss.getSheets().length > 1 && s1.getLastRow() <= 0) ss.deleteSheet(s1);
    seedExpenseDemoData_();
    exp_clearCache_();
    return { success:true, message:'Expense sheets initialised & demo data seeded', sheets:Object.keys(EXP_SCHEMAS) };
  } finally { try { lock.releaseLock(); } catch(e){} }
}

function seedExpenseDemoData_() {
  if (exp_readAll_('ExpUsers').length > 0) return;
  const now = new Date();
  const iso = d => new Date(d).toISOString().split('T')[0];
  const daysAgo = n => { const d = new Date(now); d.setDate(d.getDate()-n); return iso(d); };
  const uuid = () => Utilities.getUuid().slice(0,8).toUpperCase();

  // Users
  [
    { email:'admin@expense.com',   password:'admin123', name:'Admin Finance', role:'Admin',    department:'Finance', manager:'',              managerEmail:'' },
    { email:'finance@expense.com', password:'fin123',   name:'Finance Team',  role:'Finance',  department:'Finance', manager:'Admin Finance', managerEmail:'admin@expense.com' },
    { email:'manager@expense.com', password:'mgr123',   name:'Sales Manager', role:'Manager',  department:'Sales',   manager:'Admin Finance', managerEmail:'admin@expense.com' },
    { email:'emp@expense.com',     password:'emp123',   name:'John Employee', role:'Employee', department:'Sales',   manager:'Sales Manager', managerEmail:'manager@expense.com' },
    { email:'priya@expense.com',   password:'emp123',   name:'Priya Menon',   role:'Employee', department:'HR',      manager:'Admin Finance', managerEmail:'admin@expense.com' },
    { email:'rahul@expense.com',   password:'emp123',   name:'Rahul Sharma',  role:'Employee', department:'IT',      manager:'Sales Manager', managerEmail:'manager@expense.com' }
  ].forEach(u => exp_insert_('ExpUsers',{ id:'EU-'+uuid(), active:true, createdDate:new Date().toISOString(), ...u }));

  // Categories
  [
    ['CAT-TRV','TRV','Travel',           30000,true, 'Flights, cabs, trains, fuel'],
    ['CAT-FOD','FOD','Food & Meals',     5000, true, 'Business meals & team lunches'],
    ['CAT-ACC','ACC','Accommodation',    25000,true, 'Hotel stays during business trips'],
    ['CAT-OFF','OFF','Office Supplies',  3000, true, 'Stationery, small equipment'],
    ['CAT-CLN','CLN','Client Entertainment',10000,true,'Client meetings & entertainment'],
    ['CAT-INT','INT','Internet & Phone', 2000, true, 'Business phone / internet reimbursement'],
    ['CAT-TRN','TRN','Training',         15000,true, 'Courses, certifications, books'],
    ['CAT-OTH','OTH','Other',            5000, false,'Miscellaneous']
  ].forEach(c => exp_insert_('ExpCategories',{ id:c[0], code:c[1], name:c[2], maxLimit:c[3], requiresReceipt:c[4], description:c[5], active:true, createdDate:daysAgo(90) }));

  // Claims (10 records across statuses)
  const claims = [
    ['emp@expense.com',    'John Employee','Travel',              'CAT-TRV', 4500, daysAgo(15), 'Cab to client meeting in Bandra','https://example.com/rcpt/001.jpg','Credit Card','Approved','Sales Manager','manager@expense.com',4500, daysAgo(12),'All good'],
    ['emp@expense.com',    'John Employee','Food & Meals',        'CAT-FOD', 1200, daysAgo(10), 'Team lunch after Q2 close',      'https://example.com/rcpt/002.jpg','Cash','Approved','Sales Manager','manager@expense.com',1200, daysAgo(8), ''],
    ['emp@expense.com',    'John Employee','Accommodation',       'CAT-ACC', 8500, daysAgo(20), 'Pune client visit — 2 nights',   'https://example.com/rcpt/003.jpg','Credit Card','Reimbursed','Sales Manager','manager@expense.com',8500, daysAgo(15),''],
    ['priya@expense.com',  'Priya Menon',  'Office Supplies',     'CAT-OFF', 850,  daysAgo(6),  'Stationery for interview panel',  'https://example.com/rcpt/004.jpg','Cash','Pending','Admin Finance','admin@expense.com',0,'',''],
    ['rahul@expense.com',  'Rahul Sharma', 'Internet & Phone',    'CAT-INT', 1200, daysAgo(4),  'Home internet bill – Sept',       'https://example.com/rcpt/005.jpg','Bank Transfer','Pending','Sales Manager','manager@expense.com',0,'',''],
    ['rahul@expense.com',  'Rahul Sharma', 'Training',            'CAT-TRN', 12500,daysAgo(30), 'AWS Solutions Architect course',  'https://example.com/rcpt/006.jpg','Credit Card','Approved','Sales Manager','manager@expense.com',12500,daysAgo(25),''],
    ['priya@expense.com',  'Priya Menon',  'Client Entertainment','CAT-CLN', 3500, daysAgo(8),  'Candidate lunch – senior hire',   'https://example.com/rcpt/007.jpg','Credit Card','Approved','Admin Finance','admin@expense.com',3200, daysAgo(6),'₹300 above policy trimmed'],
    ['emp@expense.com',    'John Employee','Travel',              'CAT-TRV', 15000,daysAgo(2),  'Delhi client trip flights',       'https://example.com/rcpt/008.jpg','Credit Card','Pending','Sales Manager','manager@expense.com',0,'',''],
    ['rahul@expense.com',  'Rahul Sharma', 'Food & Meals',        'CAT-FOD', 650,  daysAgo(3),  'Working lunch on release day',    'https://example.com/rcpt/009.jpg','Cash','Rejected','Sales Manager','manager@expense.com',0, daysAgo(2),'Not a business purpose'],
    ['priya@expense.com',  'Priya Menon',  'Other',               'CAT-OTH', 800,  daysAgo(1),  'Reimbursement for office plants', 'https://example.com/rcpt/010.jpg','Cash','Pending','Admin Finance','admin@expense.com',0,'','']
  ];
  claims.forEach((c,i) => {
    const cn = 'CLM/'+now.getFullYear()+'/'+String(i+1).padStart(4,'0');
    exp_insert_('Claims',{
      id:'CLM-'+uuid(), claimNumber:cn,
      employeeId: (exp_readAll_('ExpUsers').find(u => u.email === c[0]) || {}).id || '',
      employeeName:c[1], employeeEmail:c[0],
      category:c[2], categoryId:c[3], amount:c[4], currency:'INR',
      claimDate:c[5], expenseDate:c[5], description:c[6], receipts:c[7], paymentMode:c[8],
      status:c[9], approver:c[10], approverEmail:c[11],
      approvedAmount:c[12], approvedDate:c[13], remarks:c[14],
      reimbursedDate: c[9] === 'Reimbursed' ? daysAgo(10) : '',
      reimbursedRef:  c[9] === 'Reimbursed' ? 'NEFT/UTR/'+uuid() : '',
      createdDate:c[5]
    });
  });

  // Advances (3 records)
  [
    ['emp@expense.com', 'John Employee','Chennai client visit', 20000, daysAgo(2),      'Approved',20000,'Sales Manager', daysAgo(1),  ''],
    ['priya@expense.com','Priya Menon',  'HR event catering',    15000, daysAgo(5),     'Pending', 0, '',                '',           ''],
    ['rahul@expense.com','Rahul Sharma', 'Training workshop',    8000,  daysAgo(20),    'Settled', 8000,'Admin Finance', daysAgo(18),  'Settled with expense CLM-006']
  ].forEach((a,i) => exp_insert_('Advances',{
    id:'ADV-'+uuid(), advanceNumber:'ADV/'+now.getFullYear()+'/'+String(i+1).padStart(4,'0'),
    employeeId:(exp_readAll_('ExpUsers').find(u=>u.email===a[0])||{}).id||'',
    employeeName:a[1], employeeEmail:a[0], purpose:a[2],
    requestedAmount:a[3], requiredDate:a[4], status:a[5],
    approvedAmount:a[6], approvedBy:a[7], approvedDate:a[8], settledDate: a[5]==='Settled' ? daysAgo(15) : '',
    remarks:a[9], createdDate:a[4]
  }));

  exp_clearCache_();
}

// ─── AUTH ───────────────────────────────────────────────────────────────────
function authenticateUser(email, password) {
  try {
    if (!email || !password) return { success:false, message:'Missing credentials' };
    const u = exp_readAll_('ExpUsers').find(x =>
      String(x.email||'').toLowerCase() === String(email).toLowerCase() &&
      String(x.password) === String(password) &&
      (x.active === true || String(x.active).toLowerCase() === 'true' || x.active === '' || x.active == null));
    if (!u) return { success:false, message:'Invalid credentials' };
    const user = { id:u.id, email:u.email, name:u.name, role:u.role, department:u.department, manager:u.manager, managerEmail:u.managerEmail };
    PropertiesService.getUserProperties().setProperty('_expUser', JSON.stringify(user));
    exp_log_(user.email,'login','ExpUsers',u.id,null);
    return { success:true, user };
  } catch (e){ return { success:false, message:e.message }; }
}
function getCurrentUser(){ try { const r = PropertiesService.getUserProperties().getProperty('_expUser'); return r ? JSON.parse(r) : null; } catch(e){ return null; } }
function signOut(){ PropertiesService.getUserProperties().deleteProperty('_expUser'); return { success:true }; }

// ─── CATEGORIES ─────────────────────────────────────────────────────────────
function getAllCategories(){ return exp_ok_(exp_readAll_('ExpCategories')); }
function addCategory(d){ return exp_insertWithId_('ExpCategories','CAT',d); }
function updateCategory(id,u){ return exp_updateById_('ExpCategories',id,u); }
function deleteCategory(id){ return exp_deleteById_('ExpCategories',id); }

// ─── CLAIMS ────────────────────────────────────────────────────────────────
function getAllClaims(){ return exp_ok_(exp_readAll_('Claims')); }
function getMyClaims() {
  const u = getCurrentUser(); if (!u) return exp_ok_([]);
  return exp_ok_(exp_readAll_('Claims').filter(c => c.employeeId === u.id || c.employeeEmail === u.email));
}
function getPendingClaimsForApprover() {
  const u = getCurrentUser(); if (!u) return exp_ok_([]);
  const isApprover = ['Admin','Finance','Manager'].indexOf(u.role) !== -1;
  if (!isApprover) return exp_ok_([]);
  return exp_ok_(exp_readAll_('Claims').filter(c => c.status === 'Pending' &&
    (u.role === 'Admin' || u.role === 'Finance' || c.approverEmail === u.email)));
}

function submitClaim(data) {
  try {
    const u = getCurrentUser(); if (!u) return { success:false, message:'Not logged in' };
    // Force employee identity from session (server-side guard)
    data.employeeId = u.id; data.employeeName = u.name; data.employeeEmail = u.email;
    // Approver = manager, else Admin
    if (!data.approver) {
      data.approver = u.manager || 'Admin Finance';
      data.approverEmail = u.managerEmail || EXP_CFG.DEFAULT_APPROVER_EMAIL;
    }
    data.status = 'Pending';
    data.claimNumber = data.claimNumber || generateSeqNumber_('CLM','Claims','claimNumber');
    data.claimDate = data.claimDate || new Date().toISOString().split('T')[0];
    data.currency = data.currency || 'INR';
    const res = exp_insertWithId_('Claims','CLM',data);
    if (res.success) {
      exp_insert_('ExpApprovals',{ id:'APR-'+Utilities.getUuid().slice(0,8).toUpperCase(), entityType:'Claim', entityId:res.id, requestedBy:u.name, requestedDate:new Date().toISOString(), approver:data.approver, status:'Pending', remarks:'', createdDate:new Date().toISOString() });
      sendClaimSubmittedEmail_(data);
    }
    return res;
  } catch(e){ return { success:false, message:e.message }; }
}

function approveClaim(id, approvedAmount, remarks) {
  const u = getCurrentUser();
  if (!u || ['Admin','Finance','Manager'].indexOf(u.role) === -1) return { success:false, message:'Not authorised' };
  const claim = exp_readAll_('Claims').find(c => c.id === id);
  if (!claim) return { success:false, message:'Claim not found' };
  const amt = Number(approvedAmount || claim.amount || 0);
  const updated = { status:'Approved', approvedAmount:amt, approvedDate:new Date().toISOString(), remarks:remarks||'', approver:u.name, approverEmail:u.email };
  exp_updateById_('Claims', id, updated);
  exp_updateApprovalByEntity_('Claim', id, { status:'Approved', remarks:remarks||'', decidedDate:new Date().toISOString() });
  sendClaimDecisionEmail_({ ...claim, ...updated }, 'Approved');
  return { success:true };
}

function rejectClaim(id, remarks) {
  const u = getCurrentUser();
  if (!u || ['Admin','Finance','Manager'].indexOf(u.role) === -1) return { success:false, message:'Not authorised' };
  const claim = exp_readAll_('Claims').find(c => c.id === id);
  if (!claim) return { success:false, message:'Claim not found' };
  const updated = { status:'Rejected', approvedAmount:0, approvedDate:new Date().toISOString(), remarks:remarks||'', approver:u.name, approverEmail:u.email };
  exp_updateById_('Claims', id, updated);
  exp_updateApprovalByEntity_('Claim', id, { status:'Rejected', remarks:remarks||'', decidedDate:new Date().toISOString() });
  sendClaimDecisionEmail_({ ...claim, ...updated }, 'Rejected');
  return { success:true };
}

function markClaimReimbursed(id, ref) {
  const u = getCurrentUser();
  if (!u || ['Admin','Finance'].indexOf(u.role) === -1) return { success:false, message:'Not authorised' };
  const claim = exp_readAll_('Claims').find(c => c.id === id);
  if (!claim) return { success:false, message:'Claim not found' };
  if (claim.status !== 'Approved') return { success:false, message:'Claim must be approved first' };
  exp_updateById_('Claims', id, { status:'Reimbursed', reimbursedDate:new Date().toISOString(), reimbursedRef:ref||'' });
  sendClaimDecisionEmail_({ ...claim, reimbursedRef:ref||'' }, 'Reimbursed');
  return { success:true };
}

function deleteClaim(id) {
  const u = getCurrentUser();
  const claim = exp_readAll_('Claims').find(c => c.id === id);
  if (!claim) return { success:false, message:'Not found' };
  if (u && u.role === 'Employee' && claim.employeeEmail !== u.email) return { success:false, message:'Not authorised' };
  if (claim.status !== 'Pending' && u.role === 'Employee') return { success:false, message:'Cannot delete a decided claim' };
  return exp_deleteById_('Claims', id);
}

// ─── ADVANCES ──────────────────────────────────────────────────────────────
function getAllAdvances(){ return exp_ok_(exp_readAll_('Advances')); }
function getMyAdvances() {
  const u = getCurrentUser(); if (!u) return exp_ok_([]);
  return exp_ok_(exp_readAll_('Advances').filter(a => a.employeeEmail === u.email));
}
function requestAdvance(data) {
  const u = getCurrentUser(); if (!u) return { success:false, message:'Not logged in' };
  data.employeeId = u.id; data.employeeName = u.name; data.employeeEmail = u.email;
  data.status = 'Pending';
  data.advanceNumber = data.advanceNumber || generateSeqNumber_('ADV','Advances','advanceNumber');
  const res = exp_insertWithId_('Advances','ADV',data);
  if (res.success) sendAdvanceRequestedEmail_(data);
  return res;
}
function approveAdvance(id, amount, remarks) {
  const u = getCurrentUser();
  if (!u || ['Admin','Finance','Manager'].indexOf(u.role) === -1) return { success:false, message:'Not authorised' };
  const adv = exp_readAll_('Advances').find(a => a.id === id);
  if (!adv) return { success:false, message:'Advance not found' };
  const amt = Number(amount || adv.requestedAmount || 0);
  const updated = { status:'Approved', approvedAmount:amt, approvedBy:u.name, approvedDate:new Date().toISOString(), remarks:remarks||'' };
  exp_updateById_('Advances', id, updated);
  sendAdvanceDecisionEmail_({ ...adv, ...updated }, 'Approved');
  return { success:true };
}
function rejectAdvance(id, remarks) {
  const u = getCurrentUser();
  if (!u || ['Admin','Finance','Manager'].indexOf(u.role) === -1) return { success:false, message:'Not authorised' };
  const adv = exp_readAll_('Advances').find(a => a.id === id);
  if (!adv) return { success:false, message:'Advance not found' };
  const updated = { status:'Rejected', approvedAmount:0, approvedBy:u.name, approvedDate:new Date().toISOString(), remarks:remarks||'' };
  exp_updateById_('Advances', id, updated);
  sendAdvanceDecisionEmail_({ ...adv, ...updated }, 'Rejected');
  return { success:true };
}
function deleteAdvance(id){ return exp_deleteById_('Advances', id); }

// ─── DASHBOARD ─────────────────────────────────────────────────────────────
function getDashboardStats() {
  const u = getCurrentUser();
  const claims = exp_readAll_('Claims');
  const my = u ? claims.filter(c => c.employeeEmail === u.email) : [];
  const isApprover = u && ['Admin','Finance','Manager'].indexOf(u.role) !== -1;
  const pending = isApprover ? claims.filter(c => c.status === 'Pending' && (u.role !== 'Manager' || c.approverEmail === u.email)) : [];
  return { success:true, data:{
    myTotalClaims: my.length,
    myPending: my.filter(c=>c.status==='Pending').length,
    myApproved: my.filter(c=>c.status==='Approved').length,
    myReimbursed: my.filter(c=>c.status==='Reimbursed').length,
    myPendingAmount: my.filter(c=>c.status==='Pending').reduce((s,c)=>s+Number(c.amount||0),0),
    myApprovedAmount: my.filter(c=>c.status==='Approved' || c.status==='Reimbursed').reduce((s,c)=>s+Number(c.approvedAmount||0),0),
    myReimbursedAmount: my.filter(c=>c.status==='Reimbursed').reduce((s,c)=>s+Number(c.approvedAmount||0),0),
    pendingForMe: pending.length,
    pendingForMeAmount: pending.reduce((s,c)=>s+Number(c.amount||0),0),
    totalClaims: claims.length,
    totalPending: claims.filter(c=>c.status==='Pending').length,
    totalApproved: claims.filter(c=>c.status==='Approved').length,
    totalReimbursed: claims.filter(c=>c.status==='Reimbursed').length,
    totalReimbursedAmount: claims.filter(c=>c.status==='Reimbursed').reduce((s,c)=>s+Number(c.approvedAmount||0),0)
  }};
}

// ─── EMAIL AUTOMATION ──────────────────────────────────────────────────────
function emailsEnabled_() {
  const v = PropertiesService.getScriptProperties().getProperty('EMAIL_ENABLED');
  return v === null || v === undefined || String(v).toLowerCase() !== 'false';
}
function sendMailSafe_(to, subject, htmlBody, type, entityId) {
  try {
    if (!to) throw new Error('No recipient');
    if (!emailsEnabled_()) {
      exp_log_email_(to, subject, type, entityId, 'skipped-disabled', '');
      return { success:true, skipped:true };
    }
    MailApp.sendEmail({ to:to, subject:subject, htmlBody:htmlBody });
    exp_log_email_(to, subject, type, entityId, 'sent', '');
    return { success:true };
  } catch(e){
    exp_log_email_(to, subject, type, entityId, 'error', e.message);
    Logger.log('Mail failed to ' + to + ': ' + e.message);
    return { success:false, message:e.message };
  }
}
function exp_log_email_(to, subject, type, entityId, status, err){
  try { exp_getSheet_('EmailLog').appendRow([new Date().toISOString(), to||'', subject||'', type||'', entityId||'', status||'', err||'']); } catch(e){}
}

function baseEmailStyle_(){ return `
  body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;color:#1e293b;background:#f8fafc;padding:0;margin:0}
  .box{max-width:560px;margin:24px auto;background:#fff;border-radius:14px;overflow:hidden;border:1px solid #e2e8f0}
  .hd{background:#0F172A;color:#fff;padding:18px 24px;font-size:18px;font-weight:600}
  .bd{padding:24px}
  .row{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px dashed #e2e8f0}
  .row:last-child{border:0}
  .k{color:#64748b;font-size:13px}
  .v{color:#0f172a;font-weight:600;font-size:13px}
  .amt{font-size:22px;font-weight:700;color:#0F172A;margin:12px 0}
  .badge{display:inline-block;padding:4px 10px;border-radius:20px;font-size:11px;font-weight:600;letter-spacing:.4px;text-transform:uppercase}
  .ok{background:#dcfce7;color:#166534}.warn{background:#fef3c7;color:#92400e}.bad{background:#fee2e2;color:#991b1b}.info{background:#dbeafe;color:#1e40af}
  .ft{padding:14px 24px;background:#f1f5f9;font-size:11px;color:#64748b;text-align:center}
`; }
function money_(n){ return '₹' + Number(n||0).toLocaleString('en-IN'); }

function sendClaimSubmittedEmail_(claim) {
  const to = claim.approverEmail || EXP_CFG.DEFAULT_APPROVER_EMAIL;
  const subject = `[Expense] New claim ${claim.claimNumber} · ${money_(claim.amount)} from ${claim.employeeName}`;
  const html = `<html><head><style>${baseEmailStyle_()}</style></head><body>
    <div class="box">
      <div class="hd">💰 New Expense Claim Submitted</div>
      <div class="bd">
        <p>Hi ${claim.approver||''},</p>
        <p><b>${claim.employeeName}</b> has submitted a new expense claim for your approval.</p>
        <div class="amt">${money_(claim.amount)} <span class="badge warn">Pending</span></div>
        <div class="row"><span class="k">Claim #</span><span class="v">${claim.claimNumber}</span></div>
        <div class="row"><span class="k">Category</span><span class="v">${claim.category}</span></div>
        <div class="row"><span class="k">Expense Date</span><span class="v">${claim.expenseDate||claim.claimDate}</span></div>
        <div class="row"><span class="k">Payment Mode</span><span class="v">${claim.paymentMode||'-'}</span></div>
        <div class="row"><span class="k">Description</span><span class="v" style="max-width:60%;text-align:right">${claim.description||''}</span></div>
        ${claim.receipts ? `<div class="row"><span class="k">Receipt</span><span class="v"><a href="${claim.receipts}">View receipt</a></span></div>` : ''}
        <p style="margin-top:16px;color:#64748b;font-size:13px">Please log in to the Expense Tracker to approve or reject.</p>
      </div>
      <div class="ft">Automated notification · Expense & Reimbursement Tracker</div>
    </div></body></html>`;
  return sendMailSafe_(to, subject, html, 'claim-submitted', claim.id);
}
function sendClaimDecisionEmail_(claim, decision) {
  const to = claim.employeeEmail;
  const emoji = decision === 'Approved' ? '✅' : decision === 'Rejected' ? '❌' : '💸';
  const badgeClass = decision === 'Approved' ? 'ok' : decision === 'Rejected' ? 'bad' : 'info';
  const subject = `[Expense] Your claim ${claim.claimNumber} is ${decision}`;
  const html = `<html><head><style>${baseEmailStyle_()}</style></head><body>
    <div class="box">
      <div class="hd">${emoji} Expense Claim ${decision}</div>
      <div class="bd">
        <p>Hi ${claim.employeeName},</p>
        <p>Your expense claim <b>${claim.claimNumber}</b> has been <b>${decision.toLowerCase()}</b> by ${claim.approver}.</p>
        <div class="amt">${money_(decision === 'Rejected' ? 0 : (claim.approvedAmount||claim.amount))} <span class="badge ${badgeClass}">${decision}</span></div>
        <div class="row"><span class="k">Original</span><span class="v">${money_(claim.amount)}</span></div>
        ${decision !== 'Rejected' ? `<div class="row"><span class="k">Approved</span><span class="v">${money_(claim.approvedAmount||claim.amount)}</span></div>` : ''}
        <div class="row"><span class="k">Category</span><span class="v">${claim.category}</span></div>
        <div class="row"><span class="k">Description</span><span class="v" style="max-width:60%;text-align:right">${claim.description||''}</span></div>
        ${claim.remarks ? `<div class="row"><span class="k">Remarks</span><span class="v" style="max-width:60%;text-align:right">${claim.remarks}</span></div>` : ''}
        ${claim.reimbursedRef ? `<div class="row"><span class="k">Payment Ref</span><span class="v">${claim.reimbursedRef}</span></div>` : ''}
        <p style="margin-top:16px;color:#64748b;font-size:13px">Log in to the Expense Tracker to view the full history.</p>
      </div>
      <div class="ft">Automated notification · Expense & Reimbursement Tracker</div>
    </div></body></html>`;
  return sendMailSafe_(to, subject, html, 'claim-'+decision.toLowerCase(), claim.id);
}
function sendAdvanceRequestedEmail_(adv){
  const to = EXP_CFG.DEFAULT_APPROVER_EMAIL;
  const subject = `[Advance] New request ${adv.advanceNumber} · ${money_(adv.requestedAmount)} from ${adv.employeeName}`;
  const html = `<html><head><style>${baseEmailStyle_()}</style></head><body>
    <div class="box"><div class="hd">💼 Advance Requested</div>
    <div class="bd"><p>${adv.employeeName} has requested an advance.</p>
      <div class="amt">${money_(adv.requestedAmount)} <span class="badge warn">Pending</span></div>
      <div class="row"><span class="k">Purpose</span><span class="v">${adv.purpose||''}</span></div>
      <div class="row"><span class="k">Required by</span><span class="v">${adv.requiredDate||''}</span></div>
    </div><div class="ft">Automated · Expense Tracker</div></div></body></html>`;
  return sendMailSafe_(to, subject, html, 'advance-requested', adv.id);
}
function sendAdvanceDecisionEmail_(adv, decision){
  const subject = `[Advance] ${adv.advanceNumber} ${decision}`;
  const badge = decision === 'Approved' ? 'ok' : 'bad';
  const html = `<html><head><style>${baseEmailStyle_()}</style></head><body>
    <div class="box"><div class="hd">💼 Advance ${decision}</div>
    <div class="bd"><p>Hi ${adv.employeeName},</p>
      <p>Your advance request has been ${decision.toLowerCase()} by ${adv.approvedBy}.</p>
      <div class="amt">${money_(decision==='Rejected'?0:(adv.approvedAmount||adv.requestedAmount))} <span class="badge ${badge}">${decision}</span></div>
      ${adv.remarks ? `<div class="row"><span class="k">Remarks</span><span class="v">${adv.remarks}</span></div>` : ''}
    </div><div class="ft">Automated · Expense Tracker</div></div></body></html>`;
  return sendMailSafe_(adv.employeeEmail, subject, html, 'advance-'+decision.toLowerCase(), adv.id);
}

// Manual re-send / test helpers exposed to UI
function resendClaimEmail(id){ const c = exp_readAll_('Claims').find(x=>x.id===id); if(!c) return { success:false, message:'Not found' }; return sendClaimSubmittedEmail_(c); }
function sendReminderForClaim(id){ const c = exp_readAll_('Claims').find(x=>x.id===id); if(!c) return { success:false, message:'Not found' }; return sendMailSafe_(c.approverEmail, `[Reminder] Claim ${c.claimNumber} awaiting your approval`, `<p>Hi ${c.approver}, this is a reminder that ${c.employeeName}'s claim of ${money_(c.amount)} is pending your review.</p>`, 'reminder', c.id); }

// ─── HELPERS ───────────────────────────────────────────────────────────────
function exp_updateApprovalByEntity_(entityType, entityId, updates){
  const rec = exp_readAll_('ExpApprovals').find(a => a.entityType === entityType && a.entityId === entityId && a.status === 'Pending');
  if (rec) exp_updateById_('ExpApprovals', rec.id, updates);
}
function generateSeqNumber_(prefix, sheet, field){
  const rows = exp_readAll_(sheet); const yr = new Date().getFullYear();
  const seq = rows.filter(r => String(r[field]||'').indexOf(prefix+'/'+yr) === 0).length + 1;
  return prefix+'/'+yr+'/'+String(seq).padStart(4,'0');
}

function exp_getSheet_(name){ const ss = SpreadsheetApp.getActiveSpreadsheet(); let sh = ss.getSheetByName(name); if (!sh){ sh = ss.insertSheet(name); exp_ensureHeaders_(sh, EXP_SCHEMAS[name]||['id']); } return sh; }
function exp_ensureHeaders_(sh, wanted){ const lc = Math.max(sh.getLastColumn(),1); const existing = sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v||'')); const present = existing.filter(Boolean); const missing = wanted.filter(h => present.indexOf(h) === -1); if(present.length===0 && missing.length){ sh.getRange(1,1,1,missing.length).setValues([missing]); return; } if(missing.length) sh.getRange(1,present.length+1,1,missing.length).setValues([missing]); }
function exp_getHeaders_(name){ const sh = exp_getSheet_(name); const w = EXP_SCHEMAS[name]||[]; if(w.length) exp_ensureHeaders_(sh,w); const lc = Math.max(sh.getLastColumn(),1); return sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v||'')).filter(Boolean); }
function exp_readAll_(name){ const cache = CacheService.getScriptCache(); const key = 'EXP::'+name; const hit = cache.get(key); if(hit){ try { return JSON.parse(hit); } catch(e){} } const sh = exp_getSheet_(name); const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2 || lc < 1){ cache.put(key,'[]',EXP_CFG.CACHE_TTL_SEC); return []; } const vals = sh.getRange(1,1,lr,lc).getValues(); const hdrs = vals[0].map(v=>String(v||'')); const rows = []; for (let i=1;i<vals.length;i++){ const r = vals[i]; if (r.every(v=>v===''||v===null)) continue; const o = {}; for (let j=0;j<hdrs.length;j++){ if(!hdrs[j]) continue; let v = r[j]; if (v instanceof Date) v = v.toISOString(); o[hdrs[j]] = v; } rows.push(o); } try { cache.put(key, JSON.stringify(rows), EXP_CFG.CACHE_TTL_SEC); } catch(e){} return rows; }
function exp_serialize_(name, obj){ const hdrs = exp_getHeaders_(name); const extra = Object.keys(obj).filter(k => hdrs.indexOf(k) === -1); if (extra.length){ const sh = exp_getSheet_(name); sh.getRange(1, hdrs.length+1, 1, extra.length).setValues([extra]); extra.forEach(h => hdrs.push(h)); } return hdrs.map(h => { let v = obj[h]; if (v === undefined || v === null) return ''; if (v instanceof Date) return v.toISOString(); return v; }); }
function exp_insert_(name, obj){ const sh = exp_getSheet_(name); const row = exp_serialize_(name, obj); sh.appendRow(row); exp_clearKey_(name); return obj; }
function exp_insertWithId_(name, prefix, data){ try { if(!data) return { success:false, message:'No data' }; const lock = LockService.getScriptLock(); lock.tryLock(EXP_CFG.LOCK_TIMEOUT_MS); try { const now = new Date().toISOString(); if (!data.id) data.id = prefix+'-'+Utilities.getUuid().slice(0,8).toUpperCase(); if (!data.createdDate) data.createdDate = now; exp_insert_(name, data); exp_log_((getCurrentUser()||{}).email||'system','create',name,data.id,null); return { success:true, id:data.id, data }; } finally { try { lock.releaseLock(); } catch(e){} } } catch (e){ return { success:false, message:e.message }; } }
function exp_findRow_(name, field, val){ const sh = exp_getSheet_(name); const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return { sheet:sh, rowIndex:-1, headers:[], row:null }; const vals = sh.getRange(1,1,lr,lc).getValues(); const hdrs = vals[0].map(v=>String(v||'')); const idx = hdrs.indexOf(field); if (idx === -1) return { sheet:sh, rowIndex:-1, headers:hdrs, row:null }; for (let i=1;i<vals.length;i++){ if (String(vals[i][idx]) === String(val)) return { sheet:sh, rowIndex:i+1, headers:hdrs, row:vals[i] }; } return { sheet:sh, rowIndex:-1, headers:hdrs, row:null }; }
function exp_updateById_(name, id, updates){ try { if (!id) return { success:false, message:'Missing id' }; const lock = LockService.getScriptLock(); lock.tryLock(EXP_CFG.LOCK_TIMEOUT_MS); try { const loc = exp_findRow_(name,'id',id); if (loc.rowIndex === -1) return { success:false, message:'Not found' }; const cur = {}; loc.headers.forEach((h,j)=>{ if(h) cur[h] = loc.row[j]; }); Object.assign(cur, updates); if (!cur.updatedDate) cur.updatedDate = new Date().toISOString(); const nr = exp_serialize_(name, cur); loc.sheet.getRange(loc.rowIndex, 1, 1, nr.length).setValues([nr]); exp_clearKey_(name); exp_log_((getCurrentUser()||{}).email||'system','update',name,id,Object.keys(updates)); return { success:true, id }; } finally { try { lock.releaseLock(); } catch(e){} } } catch(e){ return { success:false, message:e.message }; } }
function exp_deleteById_(name, id){ try { const lock = LockService.getScriptLock(); lock.tryLock(EXP_CFG.LOCK_TIMEOUT_MS); try { const loc = exp_findRow_(name,'id',id); if (loc.rowIndex === -1) return { success:false, message:'Not found' }; loc.sheet.deleteRow(loc.rowIndex); exp_clearKey_(name); exp_log_((getCurrentUser()||{}).email||'system','delete',name,id,null); return { success:true }; } finally { try { lock.releaseLock(); } catch(e){} } } catch(e){ return { success:false, message:e.message }; } }
function exp_ok_(rows){ return { success:true, data:rows }; }
function exp_clearKey_(name){ try { CacheService.getScriptCache().remove('EXP::'+name); } catch(e){} }
function exp_clearCache_(){ try { CacheService.getScriptCache().removeAll(Object.keys(EXP_SCHEMAS).map(s=>'EXP::'+s)); } catch(e){} }
function exp_log_(user, action, entity, entityId, payload){ try { exp_getSheet_('ExpActivityLog').appendRow([new Date().toISOString(), user||'anon', action||'', entity||'', entityId||'', payload ? JSON.stringify(payload).slice(0,4000) : '']); } catch(e){} }

function resetCache(){ exp_clearCache_(); return { success:true }; }
function healthCheck(){ const s = {}; Object.keys(EXP_SCHEMAS).forEach(n => { const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); s[n] = sh ? { rows:Math.max(0, sh.getLastRow()-1), cols:sh.getLastColumn() } : 'missing'; }); return { success:true, sheets:s, timestamp:new Date().toISOString() }; }

// ─── BULK CSV IMPORT ───────────────────────────────────────────────────────
// Accepts sheetName + array of records. Batches server-side under a lock.
// Returns { success, ok, failed, errors:[{row, message}] }.
function bulkImport(sheetName, rows) {
  try {
    if (!sheetName) return { success:false, message:'Missing sheetName' };
    if (!Array.isArray(rows) || rows.length === 0) return { success:false, message:'No rows to import' };
    if (!EXP_SCHEMAS[sheetName]) return { success:false, message:'Unknown sheet: '+sheetName };
    const prefix = { ExpCategories:'CAT', Claims:'CLM', Advances:'ADV' }[sheetName] || 'BLK';
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
          exp_insert_(sheetName, rec);
          result.ok++;
        } catch (e) {
          result.failed++;
          result.errors.push({ row: i + 1, message: String(e.message || e) });
        }
      });
      exp_clearKey_(sheetName);
      exp_log_((getCurrentUser()||{}).email||'system', 'bulkImport', sheetName, '', { ok:result.ok, failed:result.failed });
    } finally { try { lock.releaseLock(); } catch(e){} }
    return result;
  } catch (e) { return { success:false, message: e.message }; }
}
