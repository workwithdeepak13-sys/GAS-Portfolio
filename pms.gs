/*******************************************************************************
 * FLOWPULSE PMS — Backend (pms.gs)
 * ---------------------------------------------------------------------------
 * Full-featured Project Management System on Google Apps Script + Sheets:
 *   • Login / role-based access (Admin / PM / Team / Client)
 *   • Projects, Phases, Tasks (with dependencies & hierarchy), Milestones
 *   • Kanban board, Gantt timeline, Calendar view, Burndown chart
 *   • Timesheet logging with approval workflow
 *   • Comments & attachments per task
 *   • Project templates (save/load structure)
 *   • Automated risk detection, workload heatmap, budget tracking
 *   • Email notifications (assignments, @mentions, reminders)
 *   • Daily time-based trigger for overdue task reminders
 *   • initializeSheets() — Creates all 12 sheets & seeds demo data
 *
 * DEPLOY:
 *   1. Paste this as pms.gs in Apps Script
 *   2. Add HTML file named `pms` with pms.html contents
 *   3. Run initializeSheets once (grants perms + seeds demo data)
 *   4. Deploy → New deployment → Web app → Execute as Me, Access Anyone
 ******************************************************************************/

const PMS_CFG = {
  CACHE_TTL_SEC: 60,
  LOCK_TIMEOUT_MS: 15000,
  DEFAULT_APPROVER: 'admin@pms.com',
  SHEETS: {
    USERS:'PmsUsers', PROJECTS:'Projects', PHASES:'Phases', TASKS:'Tasks',
    MILESTONES:'Milestones', TIMESHEETS:'Timesheets', COMMENTS:'Comments',
    ATTACHMENTS:'Attachments', NOTIFICATIONS:'Notifications',
    TEMPLATES:'ProjectTemplates',     APPROVALS:'PmsApprovals', ACTIVITY_LOG:'PmsActivityLog', SESSIONS:'PmsSessions'
  }
};

const PMS_SCHEMAS = {
  PmsUsers:['id','email','password','name','role','active','createdDate'],
  Projects:['id','code','name','description','pmId','pmName','status','priority',
    'startDate','endDate','budget','actualCost','templateId','createdDate','updatedDate'],
  Phases:['id','projectId','name','description','order','startDate','endDate','status','createdDate','updatedDate'],
  Tasks:['id','projectId','phaseId','parentTaskId','title','description','assigneeId','assigneeName',
    'priority','status','estimatedHours','actualHours','startDate','dueDate','dependsOn',
    'createdDate','updatedDate'],
  Milestones:['id','projectId','name','targetDate','achievedDate','status','createdDate'],
  Timesheets:['id','taskId','userId','userName','date','hours','description','status','approvedBy','createdDate'],
  Comments:['id','taskId','userId','userName','role','body','createdAt'],
  Attachments:['id','entityType','entityId','fileName','driveFileId','uploadedBy','uploadedAt'],
  Notifications:['id','userId','message','link','isRead','createdAt'],
  ProjectTemplates:['id','name','description','phases','tasks','createdDate'],
  PmsApprovals:['id','entityType','entityId','requestedBy','requestedDate','approver','status','remarks','decidedDate','createdDate'],
  PmsActivityLog:['ts','user','action','entity','entityId','payload'],
  PmsSessions:['token','userId','userEmail','userName','userRole','createdAt','expiresAt']
};

const PMS_JSON_FIELDS = {
  ProjectTemplates:['phases','tasks']
};

// ─── WEB ENTRY ──────────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('pms')
    .setTitle('Flowmative · PMS')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name){ return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ─── INITIALISATION + DEMO DATA ─────────────────────────────────────────────
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock(); lock.tryLock(PMS_CFG.LOCK_TIMEOUT_MS);
  try {
    Object.entries(PMS_SCHEMAS).forEach(([name, headers]) => {
      let sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      pms_ensureHeaders_(sh, headers);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#7C3AED').setFontColor('#FFFFFF');
    });
    const s1 = ss.getSheetByName('Sheet1');
    if (s1 && ss.getSheets().length > 1 && s1.getLastRow() <= 0) ss.deleteSheet(s1);
    seedPMSDemoData_();
    pms_clearCache_();
    return { success:true, message:'PMS sheets initialised & demo data seeded', sheets:Object.keys(PMS_SCHEMAS) };
  } finally { try { lock.releaseLock(); } catch(e){} }
}

function seedPMSDemoData_() {
  if (pms_readAll_('PmsUsers').length > 0) return;
  const now = new Date();
  const iso = d => { const x = new Date(d); return x.toISOString().split('T')[0]; };
  const daysAgo = n => { const d = new Date(now); d.setDate(d.getDate()-n); return iso(d); };
  const daysAhead = n => { const d = new Date(now); d.setDate(d.getDate()+n); return iso(d); };
  const uuid = () => Utilities.getUuid().slice(0,8).toUpperCase();
  const nowIso = new Date().toISOString();

  // ── Users (5) ────────────────────────────────────────────────────────────
  const users = [
    { email:'admin@pms.com',   password:'admin123', name:'Admin User',   role:'Admin' },
    { email:'pm@pms.com',      password:'pm123',     name:'Priya Menon', role:'PM' },
    { email:'pm2@pms.com',     password:'pm123',     name:'Rahul Sharma',role:'PM' },
    { email:'team@pms.com',    password:'team123',   name:'John Employee',role:'Team' },
    { email:'client@pms.com',  password:'client123', name:'ABC Corp',    role:'Client' }
  ];
  users.forEach(u => pms_insert_('PmsUsers',{ id:'USR-'+uuid(), active:true, createdDate:nowIso, ...u }));

  const admin = pms_readAll_('PmsUsers').find(u=>u.role==='Admin');
  const priya = pms_readAll_('PmsUsers').find(u=>u.email==='pm@pms.com');
  const rahul = pms_readAll_('PmsUsers').find(u=>u.email==='pm2@pms.com');
  const neha  = pms_readAll_('PmsUsers').find(u=>u.email==='team@pms.com');

  // ── Projects (3) ─────────────────────────────────────────────────────────
  const projects = [
    { code:'PROJ-001', name:'Website Redesign', description:'Complete overhaul of corporate website with new brand guidelines', pm:priya, status:'InProgress', priority:'High', start:daysAgo(30), end:daysAhead(60), budget:500000 },
    { code:'PROJ-002', name:'Mobile App v2', description:'Version 2 of the customer mobile application', pm:rahul, status:'InProgress', priority:'Medium', start:daysAgo(15), end:daysAhead(75), budget:800000 },
    { code:'PROJ-003', name:'Data Migration', description:'Migrate legacy data to new cloud infrastructure', pm:priya, status:'Planning', priority:'Low', start:daysAhead(10), end:daysAhead(100), budget:300000 }
  ];
  projects.forEach((p,i) => pms_insert_('Projects',{
    id:'PRJ-'+(i+1).toString().padStart(3,'0'), code:p.code, name:p.name, description:p.description,
    pmId:p.pm.id, pmName:p.pm.name, status:p.status, priority:p.priority,
    startDate:p.start, endDate:p.end, budget:p.budget, actualCost:0,
    templateId:'', createdDate:daysAgo(30-i), updatedDate:nowIso
  }));

  const prj1 = pms_readAll_('Projects').find(p=>p.code==='PROJ-001');
  const prj2 = pms_readAll_('Projects').find(p=>p.code==='PROJ-002');

  // ── Phases (per project) ─────────────────────────────────────────────────
  const phaseDefs = [
    {pid:prj1.id, phases:[
      {name:'Discovery & Research',  order:1, start:daysAgo(30), end:daysAgo(15), status:'Completed'},
      {name:'Wireframing',           order:2, start:daysAgo(15), end:daysAgo(5),  status:'Completed'},
      {name:'Visual Design',         order:3, start:daysAgo(5),  end:daysAhead(15),status:'InProgress'},
      {name:'Development',           order:4, start:daysAhead(16),end:daysAhead(45),status:'Pending'},
      {name:'Testing & Launch',      order:5, start:daysAhead(45),end:daysAhead(60),status:'Pending'}
    ]},
    {pid:prj2.id, phases:[
      {name:'Requirements', order:1, start:daysAgo(15), end:daysAgo(5),  status:'Completed'},
      {name:'UI/UX Design', order:2, start:daysAgo(5),  end:daysAhead(15),status:'InProgress'},
      {name:'Backend API',  order:3, start:daysAhead(15),end:daysAhead(50),status:'Pending'},
      {name:'Frontend',     order:4, start:daysAhead(20),end:daysAhead(60),status:'Pending'},
      {name:'QA & Release', order:5, start:daysAhead(60),end:daysAhead(90),status:'Pending'}
    ]}
  ];
  let phaseIdx = 0;
  phaseDefs.forEach(pd => pd.phases.forEach(ph => {
    phaseIdx++;
    pms_insert_('Phases',{
      id:'PH-'+(phaseIdx).toString().padStart(3,'0'), projectId:pd.pid, name:ph.name,
      description:'', order:ph.order, startDate:ph.start, endDate:ph.end,
      status:ph.status, createdDate:daysAgo(30), updatedDate:nowIso
    });
  }));

  const phases1 = pms_readAll_('Phases').filter(p=>p.projectId===prj1.id);
  const phases2 = pms_readAll_('Phases').filter(p=>p.projectId===prj2.id);

  // ── Tasks (with dependencies & hierarchy) ────────────────────────────────
  const taskDefs = [
    // Project 1 tasks
    {p:prj1, ph:phases1[0], tasks:[
      {title:'Stakeholder interviews',      desc:'Interview 5 key stakeholders', assign:priya, priority:'High',   est:16, act:14, s:daysAgo(30), d:daysAgo(22), stat:'Done',       dep:''},
      {title:'Competitor analysis',        desc:'Analyse top 5 competitors',    assign:neha,  priority:'Medium', est:12, act:10, s:daysAgo(28), d:daysAgo(20), stat:'Done',       dep:''},
      {title:'Brand guidelines doc',       desc:'Document brand guidelines',     assign:priya, priority:'High',   est:8,  act:8,  s:daysAgo(22), d:daysAgo(15), stat:'Done',       dep:'TASK-001'}
    ]},
    {p:prj1, ph:phases1[1], tasks:[
      {title:'Homepage wireframe',        desc:'Desktop & mobile wireframes',   assign:neha,  priority:'High',   est:20, act:18, s:daysAgo(15), d:daysAgo(8),  stat:'Done',       dep:''},
      {title:'Inner pages wireframes',    desc:'About, contact, product pages', assign:neha,  priority:'Medium', est:16, act:0,  s:daysAgo(15), d:daysAgo(5),  stat:'InProgress', dep:''}
    ]},
    {p:prj1, ph:phases1[2], tasks:[
      {title:'Design system setup',       desc:'Colors, typography, tokens',   assign:priya, priority:'High',   est:24, act:20, s:daysAgo(5),  d:daysAhead(5), stat:'InProgress', dep:'TASK-5'},
      {title:'High-fidelity mockups',     desc:'All screens in Figma',         assign:neha,  priority:'High',   est:40, act:0,  s:daysAhead(3),d:daysAhead(15),stat:'Todo',       dep:'TASK-7'},
      {title:'Prototype & user testing',  desc:'Interactive prototype + test',  assign:priya, priority:'Medium', est:16, act:0,  s:daysAhead(15),d:daysAhead(20),stat:'Todo',    dep:'TASK-8'}
    ]},
    // Project 2 tasks
    {p:prj2, ph:phases2[0], tasks:[
      {title:'User stories workshop',    desc:'2-day workshop with product',   assign:rahul, priority:'High',   est:16, act:14, s:daysAgo(15), d:daysAgo(8),  stat:'Done',       dep:''},
      {title:'API spec document',         desc:'OpenAPI spec for all endpoints',   assign:rahul, priority:'High',   est:20, act:16, s:daysAgo(12), d:daysAgo(5),  stat:'Done',       dep:''}
    ]},
    {p:prj2, ph:phases2[1], tasks:[
      {title:'Mobile UI kit',             desc:'Component library for mobile',  assign:neha,  priority:'Medium', est:24, act:12, s:daysAgo(5),  d:daysAhead(5), stat:'InProgress', dep:''},
      {title:'Navigation & onboarding',  desc:'Splash, login, nav flow',       assign:neha,  priority:'High',   est:20, act:0,  s:daysAhead(2),d:daysAhead(12),stat:'Todo',       dep:'TASK-12'}
    ]}
  ];
  let taskIdx = 0;
  taskDefs.forEach(td => td.tasks.forEach(t => {
    taskIdx++;
    const pid = taskIdx < 10 ? 'TASK-'+taskIdx : 'TASK-'+taskIdx;
    pms_insert_('Tasks',{
      id:pid, projectId:td.p.id, phaseId:td.ph.id,
      parentTaskId:'', title:t.title, description:t.desc,
      assigneeId:t.assign.id, assigneeName:t.assign.name,
      priority:t.priority, status:t.status, estimatedHours:t.est,
      actualHours:t.act, startDate:t.s, dueDate:t.d,
      dependsOn:t.dep, createdDate:t.s, updatedDate:nowIso
    });
  }));

  // ── Milestones ───────────────────────────────────────────────────────────
  [
    {pid:prj1.id, name:'Design Phase Complete', target:daysAhead(20), achieved:'', status:'Pending'},
    {pid:prj1.id, name:'MVP Launch',            target:daysAhead(60), achieved:'', status:'Pending'},
    {pid:prj2.id, name:'API Beta Available',     target:daysAhead(40), achieved:'', status:'Pending'},
    {pid:prj1.id, name:'Research Complete',       target:daysAgo(15),   achieved:daysAgo(15), status:'Achieved'}
  ].forEach((m,i) => pms_insert_('Milestones',{
    id:'MS-'+(i+1).toString().padStart(3,'0'), projectId:m.pid, name:m.name,
    targetDate:m.target, achievedDate:m.achieved, status:m.status, createdDate:nowIso
  }));

  // ── Timesheets ───────────────────────────────────────────────────────────
  const allTasks = pms_readAll_('Tasks');
  allTasks.forEach(t => {
    if (t.actualHours > 0) {
      // Spread actual hours across multiple days
      const days = Math.ceil(Number(t.actualHours)/4);
      for (let d=0; d<days && d<10; d++) {
        const hrs = d===days-1 ? Number(t.actualHours)%4||4 : 4;
        const dt = new Date(t.startDate||now); dt.setDate(dt.getDate()+d);
        pms_insert_('Timesheets',{
          id:'TS-'+uuid(), taskId:t.id, userId:t.assigneeId, userName:t.assigneeName,
          date:iso(dt), hours:hrs, description:'Work on '+t.title,
          status:'Approved', approvedBy:t.assigneeName,
          createdDate:iso(dt)
        });
      }
    }
  });

  // ── Comments ────────────────────────────────────────────────────────────
  const doneTasks = allTasks.filter(t => t.status==='Done');
  doneTasks.slice(0,3).forEach((t,i) => pms_insert_('Comments',{
    id:'COM-'+(i+1).toString().padStart(3,'0'), taskId:t.id, userId:t.assigneeId,
    userName:t.assigneeName, role:'Team', body:'Completed this task ahead of schedule.',
    createdAt:daysAgo(i*2)
  }));

  // ── Notifications ───────────────────────────────────────────────────────
  [neha, priya].forEach((u,i) => pms_insert_('Notifications',{
    id:'NOTIF-'+uuid(), userId:u.id,
    message: i===0 ? 'You have 2 tasks due this week' : 'Website Redesign is 60% complete',
    link:'#tasks', isRead:false, createdAt:daysAgo(1)
  }));

  // ── Project Templates ────────────────────────────────────────────────────
  pms_insert_('ProjectTemplates',{
    id:'TPL-001', name:'Standard Web Project', description:'Default phases for web development',
    phases:JSON.stringify([
      {name:'Discovery',order:1},{name:'Design',order:2},{name:'Development',order:3},{name:'Testing',order:4},{name:'Deploy',order:5}
    ]),
    tasks:JSON.stringify([
      {title:'Requirements gathering',phase:1,estHours:16},{title:'Wireframes',phase:2,estHours:24},{title:'Frontend dev',phase:3,estHours:80}
    ]),
    createdDate:nowIso
  });

  pms_clearCache_();
}

// ─── AUTH ───────────────────────────────────────────────────────────────────
function pms_createSession_(user) {
  const token = Utilities.getUuid();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
  pms_insert_('PmsSessions', {
    token, userId:user.id, userEmail:user.email, userName:user.name,
    userRole:user.role, createdAt:now.toISOString(), expiresAt:expiresAt.toISOString()
  });
  return token;
}
function pms_revokeSession_(token) {
  if (!token) return;
  try {
    const loc = pms_findRow_('PmsSessions', 'token', token);
    if (loc.rowIndex > -1) loc.sheet.deleteRow(loc.rowIndex);
    pms_clearKey_('PmsSessions');
  } catch(e){}
}
function authenticateUser(email, password) {
  try {
    if (!email || !password) return { success:false, message:'Missing credentials' };
    const u = pms_readAll_('PmsUsers').find(x =>
      String(x.email||'').toLowerCase() === String(email).toLowerCase() &&
      String(x.password) === String(password) &&
      (x.active === true || String(x.active).toLowerCase() === 'true'));
    if (!u) return { success:false, message:'Invalid credentials' };
    const user = { id:u.id, email:u.email, name:u.name, role:u.role };
    const token = pms_createSession_(user);
    pms_log_(user.email, 'login', 'PmsUsers', u.id, null);
    return { success:true, user, token };
  } catch (err) { return { success:false, message:err.message }; }
}
function validateSession(token) {
  try {
    if (!token) return { success:false, message:'No token' };
    const sessions = pms_readAll_('PmsSessions');
    const s = sessions.find(x => x.token === token);
    if (!s) return { success:false, message:'Invalid session' };
    const now = new Date();
    const exp = new Date(s.expiresAt);
    if (now > exp) {
      pms_revokeSession_(token);
      return { success:false, message:'Session expired' };
    }
    const user = { id:s.userId, email:s.userEmail, name:s.userName, role:s.userRole };
    return { success:true, user };
  } catch (err) { return { success:false, message:err.message }; }
}
function signOut(token) {
  pms_revokeSession_(token);
  return { success:true };
}
function registerClient(name, email, password, company) {
  try {
    if (!name || !email || !password) return { success:false, message:'Name, Email and Password are required' };
    const existing = pms_readAll_('PmsUsers').find(x => String(x.email||'').toLowerCase() === String(email).toLowerCase());
    if (existing) return { success:false, message:'Email already registered' };
    const user = {
      id:'USR-'+Utilities.getUuid().slice(0,8).toUpperCase(),
      email, password, name, role:'Client', active:true,
      company: company || '',
      createdDate: new Date().toISOString()
    };
    pms_insert_('PmsUsers', user);
    pms_log_(email, 'register', 'PmsUsers', user.id, { company });
    return { success:true, message:'Registration successful. You can now sign in.' };
  } catch (err) { return { success:false, message:err.message }; }
}
function sendResetLink(email) {
  try {
    if (!email) return { success:false, message:'Email is required' };
    const users = pms_readAll_('PmsUsers');
    const u = users.find(x => String(x.email||'').toLowerCase() === String(email).toLowerCase());
    if (!u) return { success:false, message:'No account found with that email' };
    const resetToken = Utilities.getUuid().slice(0,12);
    const now = new Date();
    const exp = new Date(now.getTime() + 60 * 60 * 1000); // 1 hour expiry
    pms_insert_('PmsSessions', {
      token:'RESET-'+resetToken, userId:u.id, userEmail:u.email, userName:u.name,
      userRole:u.role, createdAt:now.toISOString(), expiresAt:exp.toISOString()
    });
    const resetUrl = ScriptApp.getService().getUrl() + '?reset=' + resetToken;
    try {
      pmsSendEmail_(u.email, '[Flowmative] Password Reset',
        `<div style="font-family:Fira Sans,sans-serif;max-width:480px;margin:24px auto;background:#fff;border-radius:14px;border:1px solid #e2e8f0;overflow:hidden">
          <div style="background:#7C3AED;color:#fff;padding:18px 24px;font-size:18px;font-weight:600">Password Reset</div>
          <div style="padding:24px"><p>Hi <b>${u.name}</b>,</p><p>Click below to reset your password. This link expires in 1 hour.</p>
          <a href="${resetUrl}" style="display:inline-block;margin:16px 0;padding:12px 24px;background:#7C3AED;color:#fff;text-decoration:none;border-radius:8px;font-weight:600">Reset Password</a>
          <p style="color:#64748b;font-size:12px;margin-top:16px">If you didn't request this, ignore this email.</p></div>
          <div style="padding:14px 24px;background:#f1f5f9;font-size:11px;color:#64748b;text-align:center">Flowmative · Project Management System</div></div>`);
    } catch(e) {}
    return { success:true, message:'Reset link sent to your email' };
  } catch (err) { return { success:false, message:err.message }; }
}

// ─── PROJECTS ───────────────────────────────────────────────────────────────
function getAllProjects(){ return pms_ok_(pms_readAll_('Projects')); }
function addProject(d){ return pms_insertWithId('Projects','PRJ',d); }
function updateProject(id,u){ return pms_updateById_('Projects',id,u); }
function deleteProject(id){ return pms_deleteById_('Projects',id); }

// ─── PHASES ─────────────────────────────────────────────────────────────────
function getAllPhases(){ return pms_ok_(pms_readAll_('Phases')); }
function getPhasesForProject(pid){ return pms_ok_(pms_readAll_('Phases').filter(p=>p.projectId===pid)); }
function addPhase(d){ return pms_insertWithId('Phases','PH',d); }
function updatePhase(id,u){ return pms_updateById_('Phases',id,u); }

// ─── TASKS ──────────────────────────────────────────────────────────────────
function getAllTasks(){ return pms_ok_(pms_readAll_('Tasks')); }
function getTasksForProject(pid){ return pms_ok_(pms_readAll_('Tasks').filter(t=>t.projectId===pid)); }
function getTasksForPhase(pid){ return pms_ok_(pms_readAll_('Tasks').filter(t=>t.phaseId===pid)); }
function addTask(data){
  const res = pms_insertWithId('Tasks','TASK',data);
  if (res.success && data.assigneeEmail) {
    pms_sendEmail_(data.assigneeEmail,
      `[Flowmative] New task: ${data.title}`,
      `<p>Hi ${data.assigneeName},</p><p>You've been assigned <b>${data.title}</b> in project ${data.projectId}. Due: ${data.dueDate||'N/A'}.</p>`);
  }
  return res;
} 
function updateTask(id,upd){ return pms_updateById_('Tasks',id,upd); }
function deleteTask(id){ return pms_deleteById_('Tasks',id); }
function updateTaskStatus(id, status) {
  const upd = { status, updatedDate:new Date().toISOString() };
  return pms_updateById_('Tasks', id, upd);
}

// ─── MILESTONES ─────────────────────────────────────────────────────────────
function getAllMilestones(){ return pms_ok_(pms_readAll_('Milestones')); }
function addMilestone(d){ return pms_insertWithId('Milestones','MS',d); }
function updateMilestone(id,u){ return pms_updateById_('Milestones',id,u); }

// ─── TIMESHEETS ─────────────────────────────────────────────────────────────
function getAllTimesheets(){ return pms_ok_(pms_readAll_('Timesheets')); }
function getMyTimesheets() {
  const u = getCurrentUser(); if (!u) return pms_ok_([]);
  return pms_ok_(pms_readAll_('Timesheets').filter(t=>t.userId===u.id));
}
function addTimesheet(d){ return pms_insertWithId('Timesheets','TS',d); }
function approveTimesheet(id){
  const u = getCurrentUser();
  return pms_updateById_('Timesheets', id, { status:'Approved', approvedBy:u?.name||'' });
}
function rejectTimesheet(id){ return pms_updateById_('Timesheets', id, { status:'Rejected' }); }

// ─── COMMENTS ───────────────────────────────────────────────────────────────
function getCommentsForTask(tid){ return pms_ok_(pms_readAll_('Comments').filter(c=>c.taskId===tid)); }
function addComment(data){ return pms_insertWithId('Comments','COM',data); }

// ─── ATTACHMENTS ───────────────────────────────────────────────────────────
function getAttachmentsForEntity(etype, eid){ return pms_ok_(pms_readAll_('Attachments').filter(a=>a.entityType===etype&&a.entityId===eid)); }
function addAttachment(d){ return pms_insertWithId('Attachments','ATT',d); }

// ─── NOTIFICATIONS ─────────────────────────────────────────────────────────
function getMyNotifications() {
  const u = getCurrentUser(); if (!u) return pms_ok_([]);
  return pms_ok_(pms_readAll_('Notifications').filter(n=>n.userId===u.id).sort((a,b)=>(b.createdAt||'').localeCompare(a.createdAt||'')));
}
function markNotifRead(id){ return pms_updateById_('Notifications',id,{isRead:true}); }

// ─── PROJECT TEMPLATES ───────────────────────────────────────────────────────
function getAllTemplates(){ return pms_ok_(pms_readAll_('ProjectTemplates')); }
function addTemplate(d){ return pms_insertWithId('ProjectTemplates','TPL',d); }

// ─── DERIVED DATA / ANALYTICS ─────────────────────────────────────────────────────
function getDashboardStats() {
  const projects = pms_readAll_('Projects');
  const tasks = pms_readAll_('Tasks');
  const timesheets = pms_readAll_('Timesheets');
  const milestones = pms_readAll_('Milestones');
  const u = getCurrentUser();
  const myTasks = u ? tasks.filter(t=>t.assigneeId===u.id) : [];

  // Compute burndown: remaining hours per day for last 30 days
  const burndown = [];
  const totalHrs = tasks.reduce((s,t)=>s+Number(t.estimatedHours||0), 0);
  const doneHrs = tasks.filter(t=>t.status==='Done').reduce((s,t)=>s+Number(t.estimatedHours||0), 0);
  for (let i=29; i>=0; i--) {
    const d = new Date(); d.setDate(d.getDate()-i);
    const dayStr = d.toISOString().split('T')[0];
    const doneByDay = timesheets.filter(ts=>ts.date===dayStr).reduce((s,ts)=>s+Number(ts.hours||0), 0);
    burndown.push({ date:dayStr, remaining: totalHrs - doneHrs - (burndown.reduce((a,b)=>a+b.logged,0)), logged:doneByDay });
  }

  // Risk detection: overdue tasks that block others
  const nowIso = new Date().toISOString().split('T')[0];
  const overdueTasks = tasks.filter(t => t.status!=='Done' && t.dueDate && t.dueDate < nowIso);
  const atRiskTasks = overdueTasks.filter(t => {
    const dependents = tasks.filter(ot => ot.dependsOn && ot.dependsOn.split(',').includes(t.id));
    return dependents.length > 0;
  });

  return { success:true, data:{
    totalProjects: projects.length,
    activeProjects: projects.filter(p=>p.status==='InProgress').length,
    totalTasks: tasks.length,
    doneTasks: tasks.filter(t=>t.status==='Done').length,
    overdueTasks: overdueTasks.length,
    atRiskTasks: atRiskTasks.length,
    myTasks: myTasks.length,
    myPendingTasks: myTasks.filter(t=>t.status!=='Done').length,
    totalEstimatedHours: totalHrs,
    totalLoggedHours: timesheets.reduce((s,ts)=>s+Number(ts.hours||0), 0),
    progress: totalHrs > 0 ? Math.round(doneHrs/totalHrs*100) : 0,
    burndown,
    atRiskTasks: atRiskTasks.slice(0,10),
    milestones: milestones.filter(m=>m.status!=='Achieved').length
  }};
}

function getWorkloadData() {
  const users = pms_readAll_('PmsUsers').filter(u=>u.role==='Team'||u.role==='PM'||u.role==='Admin');
  const tasks = pms_readAll_('Tasks');
  return { success:true, data: users.map(u => {
    const assigned = tasks.filter(t=>t.assigneeId===u.id);
    const totalHrs = assigned.reduce((s,t)=>s+Number(t.estimatedHours||0), 0);
    const doneHrs = assigned.filter(t=>t.status==='Done').reduce((s,t)=>s+Number(t.estimatedHours||0), 0);
    return { userId:u.id, name:u.name, role:u.role, taskCount:assigned.length, totalHours:totalHrs, doneHours:doneHrs, loadPct: totalHrs > 0 ? Math.round(doneHrs/totalHrs*100) : 0 };
  })};
}

function getProjectReport(projectId) {
  const project = pms_readAll_('Projects').find(p=>p.id===projectId);
  if (!project) return { success:false, message:'Not found' };
  const phases = pms_readAll_('Phases').filter(p=>p.projectId===projectId);
  const tasks = pms_readAll_('Tasks').filter(t=>t.projectId===projectId);
  const timesheets = pms_readAll_('Timesheets');
  const totalEst = tasks.reduce((s,t)=>s+Number(t.estimatedHours||0),0);
  const totalAct = timesheets.filter(ts=>tasks.some(t=>t.id===ts.taskId)).reduce((s,ts)=>s+Number(ts.hours||0),0);
  return { success:true, data:{
    project, phases, tasks,
    budgetUtilization: project.budget > 0 ? Math.round((Number(project.actualCost||0)/project.budget)*100) : 0,
    totalEstimatedHours: totalEst,
    totalLoggedHours: totalAct,
    progress: totalEst > 0 ? Math.round(tasks.filter(t=>t.status==='Done').reduce((s,t)=>s+Number(t.estimatedHours||0),0)/totalEst*100) : 0,
    taskStatusBreakdown: {
      todo: tasks.filter(t=>t.status==='Todo').length,
      inProgress: tasks.filter(t=>t.status==='InProgress').length,
      review: tasks.filter(t=>t.status==='Review').length,
      done: tasks.filter(t=>t.status==='Done').length
    },
    burndown: computeBurndown_(tasks, timesheets)
  }};
}

function computeBurndown_(tasks, timesheets) {
  const total = tasks.reduce((s,t)=>s+Number(t.estimatedHours||0),0);
  const done = tasks.filter(t=>t.status==='Done').reduce((s,t)=>s+Number(t.estimatedHours||0),0);
  const ideal = []; const actual = []; const labels = [];
  const doneSet = {};
  tasks.forEach(t=>{ if(t.status==='Done') doneSet[t.id]=true; });
  timesheets.filter(ts=>doneSet[ts.taskId]).forEach(ts=>{
    const label = ts.date || '';
    const idx = labels.indexOf(label);
    if(idx===-1){ labels.push(label); ideal.push(0); actual.push(Number(ts.hours||0)); }
    else actual[idx] += Number(ts.hours||0);
  });
  return { labels, idealLine:labels.map((_,i)=>total-(total/labels.length)*i), actualLine:actual.reduce((a,v,i)=>{a.push((a[i-1]||total)-v);return a;},[]) };
}

// ─── HELPERS ───────────────────────────────────────────────────────────────
function computeProjectProgress_(projectId) {
  const tasks = pms_readAll_('Tasks').filter(t=>t.projectId===projectId);
  if (!tasks.length) return 0;
  const total = tasks.reduce((s,t)=>s+Number(t.estimatedHours||0),0);
  const done = tasks.filter(t=>t.status==='Done').reduce((s,t)=>s+Number(t.estimatedHours||0),0);
  return total > 0 ? Math.round(done/total*100) : 0;
}

function pms_getSheet_(name){ const ss = SpreadsheetApp.getActiveSpreadsheet(); let sh = ss.getSheetByName(name); if(!sh){ sh = ss.insertSheet(name); pms_ensureHeaders_(sh, PMS_SCHEMAS[name]||['id']); } return sh; }
function pms_ensureHeaders_(sh, wanted){ const lc = Math.max(sh.getLastColumn(),1); const existing = sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v||'')); const present = existing.filter(Boolean); const missing = wanted.filter(h => present.indexOf(h) === -1); if(present.length===0 && missing.length){ sh.getRange(1,1,1,missing.length).setValues([missing]); return; } if(missing.length) sh.getRange(1,present.length+1,1,missing.length).setValues([missing]); }
function pms_getHeaders_(name){ const sh = pms_getSheet_(name); const w = PMS_SCHEMAS[name]||[]; if(w.length) pms_ensureHeaders_(sh,w); const lc = Math.max(sh.getLastColumn(),1); return sh.getRange(1,1,1,lc).getValues()[0].map(v=>String(v||'')).filter(Boolean); }
function pms_readAll_(name){
  const cache = CacheService.getScriptCache(); const key = 'PMS::'+name;
  const hit = cache.get(key); if(hit){ try { return JSON.parse(hit); } catch(e){} }
  const sh = pms_getSheet_(name); const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 1){ cache.put(key,'[]',PMS_CFG.CACHE_TTL_SEC); return []; }
  const vals = sh.getRange(1,1,lr,lc).getValues();
  const hdrs = vals[0].map(v=>String(v||'')); const jf = PMS_JSON_FIELDS[name]||[]; const rows = [];
  for (let i=1;i<vals.length;i++){ const r = vals[i]; if (r.every(v=>v===''||v===null)) continue;
    const o = {}; for (let j=0;j<hdrs.length;j++){ if(!hdrs[j]) continue; let v = r[j]; if (v instanceof Date) v = v.toISOString(); if (jf.indexOf(hdrs[j])!==-1 && typeof v === 'string' && v){ try { v = JSON.parse(v); } catch(e){} } o[hdrs[j]] = v; } rows.push(o); }
  try { cache.put(key, JSON.stringify(rows), PMS_CFG.CACHE_TTL_SEC); } catch(e){}
  return rows;
}
function pms_serialize_(name, obj){ const hdrs = pms_getHeaders_(name); const extra = Object.keys(obj).filter(k => hdrs.indexOf(k) === -1); if (extra.length){ const sh = pms_getSheet_(name); sh.getRange(1, hdrs.length+1, 1, extra.length).setValues([extra]); extra.forEach(h => hdrs.push(h)); } const jf = PMS_JSON_FIELDS[name]||[]; return hdrs.map(h => { let v = obj[h]; if (v === undefined || v === null) return ''; if (jf.indexOf(h) !== -1 && typeof v === 'object') return JSON.stringify(v); if (v instanceof Date) return v.toISOString(); return v; }); }
function pms_insert_(name, obj){ const sh = pms_getSheet_(name); const row = pms_serialize_(name, obj); sh.appendRow(row); pms_clearKey_(name); return obj; }
function pms_insertWithId(name, prefix, data){ try { if(!data) return { success:false, message:'No data' }; const lock = LockService.getScriptLock(); lock.tryLock(PMS_CFG.LOCK_TIMEOUT_MS); try { const now = new Date().toISOString(); if (!data.id) data.id = prefix+'-'+Utilities.getUuid().slice(0,8).toUpperCase(); if (!data.createdDate) data.createdDate = now; pms_insert_(name, data); pms_log_((getCurrentUser()||{}).email||'system','create',name,data.id,null); return { success:true, id:data.id, data }; } finally { try { lock.releaseLock(); } catch(e){} } } catch (e){ return { success:false, message:e.message }; } }
function pms_findRow_(name, field, val){ const sh = pms_getSheet_(name); const lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 2) return { sheet:sh, rowIndex:-1, headers:[], row:null }; const vals = sh.getRange(1,1,lr,lc).getValues(); const hdrs = vals[0].map(v=>String(v||'')); const idx = hdrs.indexOf(field); if (idx === -1) return { sheet:sh, rowIndex:-1, headers:hdrs, row:null }; for (let i=1;i<vals.length;i++){ if (String(vals[i][idx]) === String(val)) return { sheet:sh, rowIndex:i+1, headers:hdrs, row:vals[i] }; } return { sheet:sh, rowIndex:-1, headers:hdrs, row:null }; }
function pms_updateById_(name, id, updates){ try { if (!id) return { success:false, message:'Missing id' }; if (!updates || typeof updates !== 'object') return { success:false, message:'Missing updates' }; const lock = LockService.getScriptLock(); lock.tryLock(PMS_CFG.LOCK_TIMEOUT_MS); try { const loc = pms_findRow_(name,'id',id); if (loc.rowIndex === -1) return { success:false, message:'Not found' }; const cur = {}; loc.headers.forEach((h,j)=>{ if(h) cur[h] = loc.row[j]; }); Object.assign(cur, updates); if (!cur.updatedDate) cur.updatedDate = new Date().toISOString(); const nr = pms_serialize_(name, cur); loc.sheet.getRange(loc.rowIndex, 1, 1, nr.length).setValues([nr]); pms_clearKey_(name); pms_log_((getCurrentUser()||{}).email||'system','update',name,id,Object.keys(updates)); return { success:true, id }; } finally { try { lock.releaseLock(); } catch(e){} } } catch(e){ return { success:false, message:e.message }; } }
function pms_deleteById_(name, id){ try { if (!id) return { success:false, message:'Missing id' }; const lock = LockService.getScriptLock(); lock.tryLock(PMS_CFG.LOCK_TIMEOUT_MS); try { const loc = pms_findRow_(name,'id',id); if (loc.rowIndex === -1) return { success:false, message:'Not found' }; loc.sheet.deleteRow(loc.rowIndex); pms_clearKey_(name); pms_log_((getCurrentUser()||{}).email||'system','delete',name,id,null); return { success:true }; } finally { try { lock.releaseLock(); } catch(e){} } } catch(e){ return { success:false, message:e.message }; } }
function pms_ok_(rows){ return { success:true, data:rows }; }
function pms_clearKey_(name){ try { CacheService.getScriptCache().remove('PMS::'+name); } catch(e){} }
function pms_clearCache_(){ try { CacheService.getScriptCache().removeAll(Object.keys(PMS_SCHEMAS).map(s=>'PMS::'+s)); } catch(e){} }
function pms_log_(user, action, entity, entityId, payload){ try { pms_getSheet_('PmsActivityLog').appendRow([new Date().toISOString(), user||'anon', action||'', entity||'', entityId||'', payload ? JSON.stringify(payload).slice(0,4000) : '']); } catch(e){} }

// ─── EMAIL NOTIFICATIONS ────────────────────────────────────────────────────
function pmsSendEmail_(to, subject, htmlBody) {
  try {
    MailApp.sendEmail({ to, subject, htmlBody });
    return { success:true };
  } catch(e) { return { success:false, message:e.message }; }
}

function sendTaskAssignmentEmail_(taskData) {
  if (!taskData.assigneeEmail) return { success:false, message:'No assignee email' };
  return pmsSendEmail_(taskData.assigneeEmail,
    `[Flowmative] New task: ${taskData.title}`,
    `<div style="font-family:Fira Sans,sans-serif;max-width:560px;margin:24px auto;background:#fff;border-radius:14px;border:1px solid #e2e8f0;overflow:hidden">
      <div style="background:#7C3AED;color:#fff;padding:18px 24px;font-size:18px;font-weight:600">New Task Assigned</div>
      <div style="padding:24px">
        <p>Hi <b>${taskData.assigneeName}</b>,</p>
        <p>You've been assigned a new task in <b>${taskData.projectId||'Flowmative'}</b>.</p>
        <div style="background:#FAF5FF;padding:16px;border-radius:10px;margin:12px 0">
          <div style="font-weight:600;font-size:16px;color:#4C1D95">${taskData.title}</div>
          <div style="color:#64748b;font-size:13px;margin-top:6px">${taskData.description||''}</div>
        </div>
        <div style="margin-top:12px"><span style="color:#64748b;font-size:13px">Priority: ${taskData.priority||'Medium'}</span> · <span style="color:#64748b;font-size:13px">Due: ${taskData.dueDate||'N/A'}</span></div>
        <p style="margin-top:16px;color:#64748b;font-size:13px">Log in to Flowmative to view and start working.</p>
      </div>
      <div style="padding:14px 24px;background:#f1f5f9;font-size:11px;color:#64748b;text-align:center">Flowmative · Project Management System</div>
    </div>`);
}

// ─── DAILY TRIGGER (overdue reminders) ─────────────────────────────────────
function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='dailyOverdueReminder');
  if (triggers.length === 0) {
    ScriptApp.newTrigger('dailyOverdueReminder').timeBased().atHour(8).everyDays(1).create();
  }
}

function dailyOverdueReminder() {
  const now = new Date().toISOString().split('T')[0];
  const tasks = pms_readAll_('Tasks').filter(t => t.status !== 'Done' && t.dueDate && t.dueDate < now);
  tasks.forEach(t => {
    if (t.assigneeName && t.assigneeId) {
      pms_insert_('Notifications',{
        id:'NOTIF-'+Utilities.getUuid().slice(0,8).toUpperCase(),
        userId:t.assigneeId,
        message:`Task overdue: "${t.title}" was due ${t.dueDate}`,
        link:'/tasks', isRead:false, createdAt:new Date().toISOString()
      });
    }
  });
}

// ─── MAINTENANCE ─────────────────────────────────────────────────────────────
function resetCache(){ pms_clearCache_(); return { success:true }; }
function healthCheck(){ const s = {}; Object.keys(PMS_SCHEMAS).forEach(n => { const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); s[n] = sh ? { rows:Math.max(0, sh.getLastRow()-1), cols:sh.getLastColumn() } : 'missing'; }); return { success:true, sheets:s, timestamp:new Date().toISOString() }; }
function pmsInsertWithId(name, prefix, data) { return pms_insertWithId(name, prefix, data); }

// ─── BULK CSV IMPORT ───────────────────────────────────────────────────────
function bulkImport(sheetName, rows) {
  try {
    if (!sheetName) return { success:false, message:'Missing sheetName' };
    if (!Array.isArray(rows) || rows.length === 0) return { success:false, message:'No rows to import' };
    if (!PMS_SCHEMAS[sheetName]) return { success:false, message:'Unknown sheet: '+sheetName };
    const prefix = { Projects:'PRJ', Tasks:'TASK', Timesheets:'TS', Phases:'PH', Milestones:'MS' }[sheetName] || 'BLK';
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
          pms_insert_(sheetName, rec);
          result.ok++;
        } catch (e) { result.failed++; result.errors.push({ row:i+1, message:String(e.message||e) }); }
      });
      pms_clearKey_(sheetName);
      pms_log_((getCurrentUser()||{}).email||'system','bulkImport',sheetName,'',{ ok:result.ok, failed:result.failed });
    } finally { try { lock.releaseLock(); } catch(e){} }
    return result;
  } catch (e) { return { success:false, message:e.message }; }
}