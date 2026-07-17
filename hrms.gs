/*******************************************************************************
 * AHL · HR Management System — Backend (Code.gs)
 * ---------------------------------------------------------------------------
 * Single-file, highly optimised Google Apps Script backend that serves the
 * `hr.html` front-end. Every function called by `google.script.run` (via the
 * `gsr()` helper in hr.html) is implemented here. Data is persisted to a
 * Google Sheet — one tab per entity — with dynamic column expansion, JSON
 * serialisation for nested objects, and CacheService-backed reads.
 *
 * Key features
 *  • initializeSheets()  – One-click creation / repair of every sheet & headers
 *  • Dynamic schema       – New fields on records auto-add columns
 *  • CacheService reads   – Sub-second responses on hot paths
 *  • Batch I/O            – Single getValues() / setValues() per operation
 *  • Central CRUD helpers – DRY, easy to extend
 *  • JSON aware fields    – stageData / checklist / logs are stringified
 *  • Role-based auth      – Admin / HR / Manager / Employee via Users sheet
 *
 * Deployment
 *  1.  Copy this file as `Code.gs` inside the Apps Script project.
 *  2.  Copy `hr.html` alongside it (File → New → HTML → name it `hr`).
 *  3.  Open the script editor → run `initializeSheets` once (grant perms).
 *  4.  Deploy → New deployment → Web app (Execute as: Me, Access: Anyone).
 *  5.  Optional: paste a WhatsApp API token into Script Properties
 *      (Project settings → Script properties) as `WHATSAPP_API_URL` and
 *      `WHATSAPP_API_TOKEN` for the sendWhatsAppMessage() helper.
 ******************************************************************************/

// ══════════════════════════════════════════════════════════════════════════
//  CONFIG  ·  Sheet names, headers, defaults
// ══════════════════════════════════════════════════════════════════════════
const CFG = {
  CACHE_TTL_SEC: 60, // 1-minute cache
  LOCK_TIMEOUT_MS: 15000,
  DEFAULT_BRAND: 'ahl',
  SHEETS: {
    USERS:            'Users',
    CANDIDATES:       'Candidates',
    EMPLOYEES:        'Employees',
    REQUISITIONS:     'Requisitions',
    LEAVES:           'Leaves',
    HOLIDAYS:         'Holidays',
    FEEDBACK:         'Feedback',
    TRAININGS:        'Trainings',
    PROBLEMS:         'Problems',
    ONBOARDING:       'Onboarding',
    LIFECYCLE:        'Lifecycle',
    INTERVIEW_SCORES: 'InterviewScores',
    ACTIVITY_LOG:     'ActivityLog'
  }
};

// Canonical column order for each sheet. New fields get appended automatically.
const SCHEMAS = {
  Users: ['id','email','password','name','role','employeeId','brand','active','createdDate'],
  Candidates: [
    'id','brand','fullName','email','phone','whatsapp','position','department','experience',
    'skills','appliedThrough','requisitionId','cvLink','portfolio','expectedCtc','currentCtc',
    'noticePeriod','appliedDate','currentStage','status','interviewStatus','interviewScore',
    'archived','stageData','createdDate','updatedDate'
  ],
  Employees: [
    'id','brand','fullName','email','phone','company','department','position','joinDate',
    'status','pdfLink','dateOfBirth','gender','bloodGroup','aadharCardNumber','panNumber',
    'bankName','branch','accountNumber','ifscCode','address','emergencyContact','mentorName',
    'exitDate','exitReason','exitNotes','exitChecklist','createdDate','updatedDate'
  ],
  Requisitions: [
    'id','brand','jobTitle','department','jobLevel','employmentType','minExp','maxExp',
    'openings','budget','workLocation','requiredSkills','hiringManager','recruiterName',
    'priority','targetStart','deadline','reasonType','jobDescription','notes','status',
    'createdDate','hiredDate','hiredCandidateId','updatedDate'
  ],
  Leaves: ['id','employeeId','brand','leaveType','days','fromDate','toDate','reason','status',
    'approvedBy','approvedDate','createdDate'],
  Holidays: ['id','name','date','type','appliesTo','createdDate'],
  Feedback: ['id','employeeId','brand','feedbackType','rating','feedback','problems','author','date'],
  Trainings: ['id','brand','title','department','trainer','plannedStart','duration',
    'description','currentStage','status','stageData','createdDate','updatedDate'],
  Problems: ['id','description','raisedBy','reportedBy','department','priority','status',
    'assignedTo','owner','deadline','nextAction','category','author','stageData',
    'createdAt','updatedAt'],
  Onboarding:       ['employeeId','data','updatedAt'],
  Lifecycle:        ['employeeId','events','updatedAt'],
  InterviewScores:  ['candidateKey','score','status','difficulty','submittedAt','tabSwitches','topics','raw'],
  ActivityLog:      ['ts','user','action','entity','entityId','payload']
};

// Fields that are always JSON-stringified when written and parsed when read.
const JSON_FIELDS = {
  Candidates:      ['stageData'],
  Employees:       ['exitChecklist'],
  Trainings:       ['stageData'],
  Problems:        ['stageData'],
  Onboarding:      ['data'],
  Lifecycle:       ['events'],
  InterviewScores: ['raw']
};

// ══════════════════════════════════════════════════════════════════════════
//  WEB APP ENTRY  ·  doGet / include
// ══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('hr')
    .setTitle('Flowmative · HR Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

// ══════════════════════════════════════════════════════════════════════════
//  INITIALISATION  ·  Create sheets, headers, seed default users
// ══════════════════════════════════════════════════════════════════════════
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.tryLock(CFG.LOCK_TIMEOUT_MS);

  Object.entries(SCHEMAS).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    ensureHeaders_(sh, headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1F2937')
      .setFontColor('#FFFFFF');
  });

  // Drop the default "Sheet1" if empty and other sheets exist.
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1 && s1.getLastRow() <= 0) ss.deleteSheet(s1);

  seedDefaultUsers_();
  seedHRDemoData_();
  clearCache_();
  lock.releaseLock();
  return { success: true, message: 'All sheets initialised & demo data seeded', sheets: Object.keys(SCHEMAS) };
}

function seedDefaultUsers_() {
  const rows = readAll_('Users');
  if (rows.length) return; // already seeded

  const now = new Date().toISOString();
  const defaults = [
    { email: 'admin@hr.com',   password: 'admin123', name: 'Admin HR',      role: 'Admin',    brand: 'all',  employeeId: '' },
    { email: 'hr@hr.com',      password: 'hr123',    name: 'HR Manager',    role: 'HR',       brand: 'all',  employeeId: '' },
    { email: 'manager@hr.com', password: 'mgr123',   name: 'Team Manager',  role: 'Manager',  brand: 'ahl',  employeeId: '' },
    { email: 'emp@hr.com',     password: 'emp123',   name: 'John Employee', role: 'Employee', brand: 'ahl',  employeeId: 'EMP-0001' }
  ];
  defaults.forEach(u => insert_('Users', {
    id: 'USR-' + Utilities.getUuid().slice(0, 8),
    active: true, createdDate: now, ...u
  }));
}

// ══════════════════════════════════════════════════════════════════════════
//  DEMO DATA SEEDING · Employees, Requisitions, Leaves, Holidays, Feedback,
//  Trainings, Problems, Onboarding, Lifecycle
// ══════════════════════════════════════════════════════════════════════════
function seedHRDemoData_() {
  const now = new Date();
  const iso = d => new Date(d).toISOString().split('T')[0];
  const daysAgo = n => { const d = new Date(now); d.setDate(d.getDate()-n); return iso(d); };
  const daysAhead = n => { const d = new Date(now); d.setDate(d.getDate()+n); return iso(d); };
  const nowIso = new Date().toISOString();

  // ── EMPLOYEES (skip if any exist) ────────────────────────────────────────
  if (readAll_('Employees').length === 0) {
    const employees = [
      ['EMP-0001','ahl',     'John Employee',   'emp@hr.com',         '+919812300001','American Hairline','Sales',          'Sales Executive',       daysAgo(240),'Active','1992-04-12','Male','O+','1234-5678-9012','ABCDE1234F','HDFC Bank','Andheri','501234567890','HDFC0000123','12 Marina Heights, Andheri West, Mumbai','+919812300101','Priya Menon'],
      ['EMP-0002','ahl',     'Priya Menon',     'priya@hr.com',       '+919812300002','American Hairline','HR',             'HR Executive',          daysAgo(420),'Active','1990-08-20','Female','A+','1234-5678-9013','ABCDE1234G','ICICI Bank','Bandra','501234567891','ICIC0000123','7B Palm Grove, Bandra West, Mumbai','+919812300102','HR Manager'],
      ['EMP-0003','alchamae','Rahul Sharma',    'rahul@hr.com',       '+919812300003','Alchamae',         'IT',             'Backend Developer',     daysAgo(90), 'Active','1994-01-15','Male','B+','1234-5678-9014','ABCDE1234H','SBI',       'Powai',  '501234567892','SBIN0000123','Flat 402, Powai Lakes, Mumbai',       '+919812300103','Team Manager'],
      ['EMP-0004','ahl',     'Neha Kapoor',     'neha@hr.com',        '+919812300004','American Hairline','Marketing',      'Marketing Lead',        daysAgo(150),'Active','1988-11-30','Female','AB+','1234-5678-9015','ABCDE1234I','Axis Bank','Colaba', '501234567893','UTIB0000123','9 Sea View, Colaba, Mumbai',           '+919812300104','Team Manager'],
      ['EMP-0005','ydigital','Karan Mehta',     'karan@hr.com',       '+919812300005','YDigital',         'AI Team',        'AI Engineer',           daysAgo(60), 'Active','1995-06-22','Male','O-','1234-5678-9016','ABCDE1234J','Kotak',    'Malad',  '501234567894','KKBK0000123','B-102 Malad East, Mumbai',            '+919812300105','Team Manager'],
      ['EMP-0006','ahl',     'Ananya Iyer',     'ananya@hr.com',      '+919812300006','American Hairline','Operations',     'Operations Manager',    daysAgo(720),'Active','1985-03-05','Female','A-','1234-5678-9017','ABCDE1234K','HDFC Bank','Worli',  '501234567895','HDFC0000456','Sea Face Apts, Worli, Mumbai',        '+919812300106','Admin HR'],
      ['EMP-0007','alchamae','Vikram Singh',    'vikram@hr.com',      '+919812300007','Alchamae',         'Automation Team','Automation Engineer',   daysAgo(30), 'Active','1996-09-18','Male','B-','1234-5678-9018','ABCDE1234L','ICICI Bank','Vashi', '501234567896','ICIC0000456','Vashi Sector 17, Navi Mumbai',        '+919812300107','Team Manager'],
      ['EMP-0008','ahl',     'Meera Joshi',     'meera@hr.com',       '+919812300008','American Hairline','Finance',        'Finance Analyst',       daysAgo(500),'Active','1991-07-11','Female','AB-','1234-5678-9019','ABCDE1234M','SBI',      'Thane',  '501234567897','SBIN0000456','Hiranandani Estate, Thane',           '+919812300108','Admin HR'],
      ['EMP-0009','ydigital','Aarav Kapoor',    'aarav@hr.com',       '+919812300009','YDigital',         'Sales',          'Business Development',  daysAgo(180),'Active','1993-12-08','Male','O+','1234-5678-9020','ABCDE1234N','Axis Bank','Andheri','501234567898','UTIB0000456','Lokhandwala, Andheri West',           '+919812300109','Team Manager'],
      ['EMP-0010','ahl',     'Ritika Bansal',   'ritika@hr.com',      '+919812300010','American Hairline','HR',             'Recruitment Specialist',daysAgo(45), 'Active','1994-05-25','Female','A+','1234-5678-9021','ABCDE1234O','Kotak',    'Bandra', '501234567899','KKBK0000456','Bandra Kurla Complex, Mumbai',        '+919812300110','HR Manager']
    ];
    employees.forEach(e => insert_('Employees', {
      id:e[0], brand:e[1], fullName:e[2], email:e[3], phone:e[4], company:e[5], department:e[6], position:e[7], joinDate:e[8], status:e[9],
      pdfLink:'', dateOfBirth:e[10], gender:e[11], bloodGroup:e[12], aadharCardNumber:e[13], panNumber:e[14],
      bankName:e[15], branch:e[16], accountNumber:e[17], ifscCode:e[18], address:e[19], emergencyContact:e[20], mentorName:e[21],
      exitDate:'', exitReason:'', exitNotes:'', exitChecklist:{}, createdDate:e[8]
    }));
  }

  // ── REQUISITIONS ─────────────────────────────────────────────────────────
  if (readAll_('Requisitions').length === 0) {
    const reqs = [
      ['ahl',     'Senior Sales Executive',   'Sales',          'Senior',    'Full-time', 3, 6, 2, '8-12 LPA','Mumbai',      'Sales, Negotiation, CRM',        'Ananya Iyer','Priya Menon','High',   daysAhead(30), daysAhead(45), 'Expansion',    'Drive B2B sales in West India',                            'Prefer FMCG background',  'Open'],
      ['alchamae','Automation QA Engineer',   'Automation Team','Mid',       'Full-time', 2, 4, 1, '10-14 LPA','Bengaluru',   'Selenium, Cypress, CI/CD',       'Team Manager','Ritika Bansal','Urgent', daysAhead(15), daysAhead(30), 'Replacement', 'Own end-to-end automation for Alchamae products',           '',                        'Open'],
      ['ydigital','AI/ML Engineer',           'AI Team',        'Senior',    'Full-time', 4, 8, 1, '20-30 LPA','Remote',      'Python, PyTorch, LLMs, RAG',     'Karan Mehta','HR Manager', 'High',   daysAhead(45), daysAhead(60), 'New role',    'Build production-grade LLM applications',                    'Publications a plus',     'Open'],
      ['ahl',     'HR Business Partner',       'HR',             'Manager',   'Full-time', 5, 8, 1, '12-16 LPA','Mumbai',      'Employee Relations, TA, POSH',   'Admin HR','HR Manager','Normal', daysAhead(60), daysAhead(75), 'Expansion',   'Support 3 business units across Mumbai',                     '',                        'Open'],
      ['ahl',     'Marketing Intern',          'Marketing',      'Intern',    'Internship',0, 1, 3, '25k stipend','Mumbai',    'Content, Social, Copywriting',   'Neha Kapoor','Ritika Bansal','Low',    daysAhead(20), daysAhead(30), 'Seasonal',   '6-month internship, potential to convert',                  '',                        'Open'],
      ['ydigital','Frontend Engineer',         'IT',             'Mid',       'Full-time', 3, 5, 1, '15-20 LPA','Remote',      'React, TypeScript, Tailwind',    'Rahul Sharma','Team Manager','High',  daysAhead(25), daysAhead(40), 'New role',    'Own the Alchamae web console',                              '',                        'Filled']
    ];
    reqs.forEach((r,i) => insert_('Requisitions', {
      id:'REQ-'+String(i+1).padStart(4,'0'), brand:r[0], jobTitle:r[1], department:r[2], jobLevel:r[3], employmentType:r[4],
      minExp:r[5], maxExp:r[6], openings:r[7], budget:r[8], workLocation:r[9], requiredSkills:r[10],
      hiringManager:r[11], recruiterName:r[12], priority:r[13], targetStart:r[14], deadline:r[15],
      reasonType:r[16], jobDescription:r[17], notes:r[18], status:r[19],
      createdDate: daysAgo(10 - i), hiredDate: r[19]==='Filled' ? daysAgo(2) : '', hiredCandidateId:''
    }));
  }

  // ── LEAVES ───────────────────────────────────────────────────────────────
  if (readAll_('Leaves').length === 0) {
    const leaves = [
      ['EMP-0001','ahl',    'Sick Leave',      1,   daysAgo(3),  daysAgo(3),  'Fever',                        'Approved','Team Manager', daysAgo(2)],
      ['EMP-0001','ahl',    'Casual Leave',    2,   daysAhead(5),daysAhead(6),'Family function',              'Pending', '',             ''],
      ['EMP-0002','ahl',    'Earned Leave',    5,   daysAgo(20), daysAgo(16), 'Family vacation to Goa',       'Approved','Admin HR',     daysAgo(25)],
      ['EMP-0003','alchamae','Work From Home',  3,   daysAgo(10), daysAgo(8),  'Home internet setup',          'Approved','Team Manager', daysAgo(11)],
      ['EMP-0004','ahl',    'Emergency Leave', 1,   daysAgo(1),  daysAgo(1),  'Family emergency',             'Approved','Admin HR',     daysAgo(1)],
      ['EMP-0005','ydigital','Sick Leave',      2,   daysAgo(7),  daysAgo(6),  'Viral fever',                  'Approved','Team Manager', daysAgo(8)],
      ['EMP-0007','alchamae','Casual Leave',    1,   daysAhead(2),daysAhead(2),'Personal work',                'Pending', '',             ''],
      ['EMP-0009','ydigital','Earned Leave',    3,   daysAhead(10),daysAhead(12),'Long weekend',              'Pending', '',             '']
    ];
    leaves.forEach((l,i) => insert_('Leaves', {
      id:'LV-'+String(i+1).padStart(4,'0'), employeeId:l[0], brand:l[1], leaveType:l[2], days:l[3],
      fromDate:l[4], toDate:l[5], reason:l[6], status:l[7], approvedBy:l[8], approvedDate:l[9],
      createdDate: daysAgo(15-i)
    }));
  }

  // ── HOLIDAYS (FY 2026) ───────────────────────────────────────────────────
  if (readAll_('Holidays').length === 0) {
    const yr = now.getFullYear();
    const holidays = [
      ['New Year',              yr + '-01-01', 'Company', 'both'],
      ['Republic Day',          yr + '-01-26', 'National','both'],
      ['Holi',                  yr + '-03-14', 'Festival','both'],
      ['Good Friday',           yr + '-04-18', 'Festival','both'],
      ['Independence Day',      yr + '-08-15', 'National','both'],
      ['Ganesh Chaturthi',      yr + '-09-06', 'Festival','ahl'],
      ['Gandhi Jayanti',        yr + '-10-02', 'National','both'],
      ['Diwali',                yr + '-11-01', 'Festival','both'],
      ['Christmas',             yr + '-12-25', 'Festival','both']
    ];
    holidays.forEach((h,i) => insert_('Holidays', {
      id:'HOL-'+String(i+1).padStart(4,'0'), name:h[0], date:h[1], type:h[2], appliesTo:h[3], createdDate:nowIso
    }));
  }

  // ── FEEDBACK ─────────────────────────────────────────────────────────────
  if (readAll_('Feedback').length === 0) {
    const feedback = [
      ['EMP-0001','ahl',     'Performance',5,'Consistently exceeds quarterly targets','',                       'Ananya Iyer',  daysAgo(15)],
      ['EMP-0002','ahl',     'Behavior',   4,'Great team player, supportive to peers', '',                       'Admin HR',     daysAgo(20)],
      ['EMP-0003','alchamae','Achievement',5,'Delivered automation framework 2 weeks early','',                  'Team Manager', daysAgo(5)],
      ['EMP-0004','ahl',     'Performance',4,'Strong quarter for campaigns',            'Needs more delegation','Admin HR',     daysAgo(30)],
      ['EMP-0005','ydigital','Training',   3,'Attended AWS certification','Needs to complete PyTorch course',    'Team Manager', daysAgo(10)],
      ['EMP-0007','alchamae','Issue',      2,'Late to standups thrice this week','Discussed 1-on-1',             'Team Manager', daysAgo(3)]
    ];
    feedback.forEach((f,i) => insert_('Feedback', {
      id:'FB-'+String(i+1).padStart(4,'0'), employeeId:f[0], brand:f[1], feedbackType:f[2], rating:f[3],
      feedback:f[4], problems:f[5], author:f[6], date:f[7]
    }));
  }

  // ── TRAININGS ────────────────────────────────────────────────────────────
  if (readAll_('Trainings').length === 0) {
    const trainings = [
      ['ahl',     'POSH Awareness Session',        'HR',            'External Trainer',daysAhead(10),'2 hours', 'Mandatory annual POSH compliance training','Planning',        'Upcoming'],
      ['ahl',     'Sales Excellence Bootcamp',      'Sales',         'Ananya Iyer',     daysAhead(20),'3 days',  'Deep-dive into consultative selling',      'Curriculum',      'Upcoming'],
      ['alchamae','Selenium 4 Migration',           'Automation Team','Team Manager',   daysAgo(5),   '1 day',   'Migrate legacy Selenium 3 suites',         'Delivery',        'Ongoing'],
      ['ydigital','LLM Prompt Engineering',         'AI Team',       'Karan Mehta',     daysAgo(20),  '4 hours', 'Best practices for production prompts',    'Completed',       'Completed'],
      ['ahl',     'First Aid & Safety',              'Operations',    'External Trainer',daysAhead(45),'1 day',   'Basic first aid for office',               'Planning',        'Upcoming']
    ];
    trainings.forEach((t,i) => insert_('Trainings', {
      id:'TR-'+String(i+1).padStart(4,'0'), brand:t[0], title:t[1], department:t[2], trainer:t[3],
      plannedStart:t[4], duration:t[5], description:t[6], currentStage:t[7], status:t[8],
      stageData:{}, createdDate: daysAgo(30-i)
    }));
  }

  // ── PROBLEMS ─────────────────────────────────────────────────────────────
  if (readAll_('Problems').length === 0) {
    const problems = [
      ['Meeting rooms overbooked on Mondays','Admin HR','Neha Kapoor','Operations','Medium','Open',       'Ananya Iyer','Admin HR', daysAhead(7), 'Roll out room-booking policy',    'Facilities','Admin HR'],
      ['Laptop delivery delayed for new joiners','HR Manager','Ritika Bansal','IT','High','In Progress','Rahul Sharma','Team Manager',daysAhead(3),'Escalate to vendor',              'IT',        'HR Manager'],
      ['Payroll date shifted by 2 days',      'Admin HR','Meera Joshi','Finance','Urgent','Open',       'Meera Joshi','Admin HR', daysAhead(1), 'Communicate to all employees',    'Payroll',   'Admin HR'],
      ['Coffee machine broken 2nd floor',     'Admin HR','Aarav Kapoor','Operations','Low','Resolved',   'Facilities','Admin HR', daysAgo(3),   'Replaced with new unit',          'Facilities','Admin HR']
    ];
    problems.forEach((p,i) => insert_('Problems', {
      id:'PRB-'+String(i+1).padStart(4,'0'), description:p[0], raisedBy:p[1], reportedBy:p[2],
      department:p[3], priority:p[4], status:p[5], assignedTo:p[6], owner:p[7], deadline:p[8],
      nextAction:p[9], category:p[10], author:p[11], stageData:{},
      createdAt: daysAgo(10-i), updatedAt: daysAgo(5-i)
    }));
  }

  // ── ONBOARDING (progress by employeeId) ──────────────────────────────────
  if (readAll_('Onboarding').length === 0) {
    const stages = ['Documents Collected','Laptop Assigned','Email Setup','Orientation Done','Manager Intro','First Task Assigned','30-day Review','60-day Review','90-day Confirmation'];
    const onboarding = [
      ['EMP-0005','Karan Mehta',['Documents Collected','Laptop Assigned','Email Setup','Orientation Done','Manager Intro','First Task Assigned']],   // 6/9 done
      ['EMP-0007','Vikram Singh',['Documents Collected','Laptop Assigned','Email Setup','Orientation Done']],                                        // 4/9 done
      ['EMP-0010','Ritika Bansal',['Documents Collected','Laptop Assigned','Email Setup']]                                                           // 3/9 done
    ];
    onboarding.forEach(o => {
      const checklist = {};
      stages.forEach(s => { checklist[s] = { done: o[2].indexOf(s) !== -1, note:'', ts: o[2].indexOf(s) !== -1 ? daysAgo(20) : '' }; });
      insert_('Onboarding', { employeeId:o[0], data:{ employeeName:o[1], stages:stages, checklist:checklist, meetingNotes:[] }, updatedAt:nowIso });
    });
  }

  // ── LIFECYCLE EVENTS ─────────────────────────────────────────────────────
  if (readAll_('Lifecycle').length === 0) {
    [
      ['EMP-0001', [
        { type:'Joined',     date: daysAgo(240), note:'Onboarded to Sales team' },
        { type:'Promotion',  date: daysAgo(60),  note:'Promoted to Sales Executive' },
        { type:'Achievement',date: daysAgo(15),  note:'Q3 top performer' }
      ]],
      ['EMP-0002', [
        { type:'Joined',    date: daysAgo(420), note:'Joined HR team' },
        { type:'Anniversary',date: daysAgo(60), note:'1-year work anniversary' }
      ]],
      ['EMP-0006', [
        { type:'Joined',    date: daysAgo(720), note:'Founding Operations Manager' },
        { type:'Anniversary',date: daysAgo(720-365), note:'1-year anniversary' },
        { type:'Anniversary',date: daysAgo(720-730), note:'2-year anniversary' }
      ]]
    ].forEach(l => insert_('Lifecycle', { employeeId:l[0], events:l[1], updatedAt:nowIso }));
  }

  // ── CANDIDATES (call existing function, if empty) ────────────────────────
  if (readAll_('Candidates').length === 0) {
    try { seedDemoCandidates(); } catch (e) { /* non-fatal */ }
  }
}

// ══════════════════════════════════════════════════════════════════════════
//  AUTHENTICATION
// ══════════════════════════════════════════════════════════════════════════
function authenticateUser(email, password, brand) {
  try {
    if (!email || !password) return { success: false, message: 'Missing credentials' };
    const users = readAll_('Users');
    const u = users.find(x =>
      String(x.email || '').toLowerCase() === String(email).toLowerCase() &&
      String(x.password) === String(password) &&
      (x.active === true || String(x.active).toLowerCase() === 'true' || x.active === '' || x.active == null)
    );
    if (!u) return { success: false, message: 'Invalid credentials' };
    const user = {
      id: u.id, email: u.email, name: u.name, role: u.role,
      employeeId: u.employeeId || null, brand: u.brand || 'all'
    };
    PropertiesService.getUserProperties().setProperty('_currentUser', JSON.stringify(user));
    logActivity_(user.email, 'login', 'Users', u.id, { brand: brand || u.brand });
    return { success: true, user: user };
  } catch (err) { return { success: false, message: err.message }; }
}

function getCurrentUser() {
  try {
    const raw = PropertiesService.getUserProperties().getProperty('_currentUser');
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

function signOut() {
  PropertiesService.getUserProperties().deleteProperty('_currentUser');
  return { success: true };
}

// ══════════════════════════════════════════════════════════════════════════
//  CANDIDATES
// ══════════════════════════════════════════════════════════════════════════
function getAllCandidates()       { return okData_(readAll_('Candidates')); }
function getCandidates(brand)     { return okData_(readAll_('Candidates').filter(c => !brand || brand === 'all' || (c.brand || 'ahl') === brand)); }
function addCandidate(data)       { return insertWithId_('Candidates', 'CAND', data); }
function updateCandidate(id, upd) { return updateById_('Candidates', id, upd); }

function updateCandidateInterviewStatus(id, status, reason) {
  const upd = { interviewStatus: status, updatedDate: new Date().toISOString() };
  if (status === 'Dropped' || status === 'Rejected') {
    upd.archived = true; upd.status = 'Archived';
  }
  if (reason) upd.rejectionReason = reason;
  return updateById_('Candidates', id, upd);
}

function seedDemoCandidates() {
  try {
    const sh = getSheet_('Candidates');
    // Clear existing rows (keep header)
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    clearCacheKey_('Candidates');

    const today = new Date();
    const iso = d => new Date(d).toISOString().split('T')[0];
    const stages = ['Application Received','Pre-Screening','Confirmation Call','Joining','Documentation','Onboarding'];
    const demo = [
      ['Aarav Kapoor', 'aarav@example.com', '+919812345671', 'Sales Executive', 'Sales', 'ahl', 3, 'Application Received'],
      ['Priya Sharma', 'priya@example.com', '+919812345672', 'Marketing Lead',  'Marketing', 'ahl', 5, 'Pre-Screening'],
      ['Rohit Verma',  'rohit@example.com', '+919812345673', 'AI Engineer',     'AI Team', 'alchamae', 4, 'Confirmation Call'],
      ['Neha Iyer',    'neha@example.com',  '+919812345674', 'HR Executive',    'HR', 'ahl', 2, 'Joining'],
      ['Karan Mehta',  'karan@example.com', '+919812345675', 'Backend Dev',     'IT', 'ydigital', 6, 'Onboarding']
    ];
    demo.forEach(d => {
      const sd = {};
      let cursor = new Date(today); cursor.setDate(cursor.getDate() - 30);
      stages.forEach(s => {
        const id = s.toLowerCase().replace(/[^a-z]/g, '').slice(0, 3);
        sd[id] = { plannedDate: iso(cursor), actualDate: '', notes: '', checklist: {} };
        cursor.setDate(cursor.getDate() + 5);
      });
      insert_('Candidates', {
        id: 'CAND-' + Utilities.getUuid().slice(0, 8),
        brand: d[5], fullName: d[0], email: d[1], phone: d[2],
        position: d[3], department: d[4], experience: d[6], skills: '',
        appliedThrough: 'LinkedIn', requisitionId: '',
        appliedDate: iso(today), currentStage: d[7], status: 'Active',
        stageData: JSON.stringify(sd),
        createdDate: new Date().toISOString()
      });
    });
    clearCacheKey_('Candidates');
    return { success: true, message: 'Seeded ' + demo.length + ' candidates' };
  } catch (e) { return { success: false, message: e.message }; }
}

// ══════════════════════════════════════════════════════════════════════════
//  EMPLOYEES
// ══════════════════════════════════════════════════════════════════════════
function getAllEmployees()        { return okData_(readAll_('Employees')); }
function addEmployee(data)        { return insertWithId_('Employees', 'EMP', data); }
function updateEmployee(id, upd)  { return updateById_('Employees', id, upd); }
function deleteEmployee(id)       { return deleteById_('Employees', id); }

// ══════════════════════════════════════════════════════════════════════════
//  REQUISITIONS
// ══════════════════════════════════════════════════════════════════════════
function getAllRequisitions()          { return okData_(readAll_('Requisitions')); }
function addRequisition(data)          { return insertWithId_('Requisitions', 'REQ', data); }
function updateRequisition(id, upd)    { return updateById_('Requisitions', id, upd); }
function deleteRequisition(id)         { return deleteById_('Requisitions', id); }

// ══════════════════════════════════════════════════════════════════════════
//  LEAVES
// ══════════════════════════════════════════════════════════════════════════
function getAllLeaves()          { return okData_(readAll_('Leaves')); }
function addLeave(data)          { return insertWithId_('Leaves', 'LV', data); }
function updateLeave(id, upd)    {
  if (upd && (upd.status === 'Approved' || upd.status === 'Rejected')) {
    const u = getCurrentUser();
    upd.approvedBy   = u ? (u.name || u.email) : 'HR';
    upd.approvedDate = new Date().toISOString();
  }
  return updateById_('Leaves', id, upd);
}
function deleteLeave(id) { return deleteById_('Leaves', id); }

// ══════════════════════════════════════════════════════════════════════════
//  HOLIDAYS
// ══════════════════════════════════════════════════════════════════════════
function getAllHolidays()   { return okData_(readAll_('Holidays')); }
function addHoliday(data)   { return insertWithId_('Holidays', 'HOL', data); }
function deleteHoliday(id)  { return deleteById_('Holidays', id); }

// ══════════════════════════════════════════════════════════════════════════
//  FEEDBACK
// ══════════════════════════════════════════════════════════════════════════
function getAllFeedback()  { return okData_(readAll_('Feedback')); }
function addFeedback(data) { return insertWithId_('Feedback', 'FB', data); }
function deleteFeedback(id){ return deleteById_('Feedback', id); }

// ══════════════════════════════════════════════════════════════════════════
//  TRAININGS
// ══════════════════════════════════════════════════════════════════════════
function getAllTrainings()          { return okData_(readAll_('Trainings')); }
function addTraining(data)          { return insertWithId_('Trainings', 'TR', data); }
function updateTraining(id, upd)    { return updateById_('Trainings', id, upd); }
function deleteTraining(id)         { return deleteById_('Trainings', id); }

// ══════════════════════════════════════════════════════════════════════════
//  PROBLEMS  ·  saveProblem() handles both create + update
// ══════════════════════════════════════════════════════════════════════════
function getAllProblems() { return okData_(readAll_('Problems')); }

function saveProblem(data) {
  try {
    if (!data) return { success: false, message: 'No data' };
    if (data.id) {
      const existing = readAll_('Problems').find(p => String(p.id) === String(data.id));
      if (existing) return updateById_('Problems', data.id, data);
    }
    // create new
    const now = new Date().toISOString();
    if (!data.createdAt) data.createdAt = now;
    if (!data.updatedAt) data.updatedAt = now;
    return insertWithId_('Problems', 'PRB', data);
  } catch (e) { return { success: false, message: e.message }; }
}
function deleteProblem(id) { return deleteById_('Problems', id); }

// ══════════════════════════════════════════════════════════════════════════
//  ONBOARDING PROGRESS  ·  keyed by employeeId, blob JSON
// ══════════════════════════════════════════════════════════════════════════
function getAllOnboardingProgress() {
  const rows = readAll_('Onboarding');
  const out = {};
  rows.forEach(r => { if (r.employeeId) out[r.employeeId] = r.data || {}; });
  return { success: true, data: out };
}
function saveOnboardingProgress(empId, data) {
  if (!empId) return { success: false, message: 'Missing employeeId' };
  return upsertByKey_('Onboarding', 'employeeId', empId, {
    employeeId: empId, data: data || {}, updatedAt: new Date().toISOString()
  });
}

// ══════════════════════════════════════════════════════════════════════════
//  LIFECYCLE EVENTS  ·  keyed by employeeId, array of events
// ══════════════════════════════════════════════════════════════════════════
function getAllLifecycleEvents() {
  const rows = readAll_('Lifecycle');
  const out = {};
  rows.forEach(r => { if (r.employeeId) out[r.employeeId] = Array.isArray(r.events) ? r.events : []; });
  return { success: true, data: out };
}
function saveLifecycleEvent(empId, event) {
  if (!empId || !event) return { success: false, message: 'Missing empId or event' };
  const rows = readAll_('Lifecycle');
  const existing = rows.find(r => r.employeeId === empId);
  const events = existing && Array.isArray(existing.events) ? existing.events.slice() : [];
  events.push(event);
  return upsertByKey_('Lifecycle', 'employeeId', empId, {
    employeeId: empId, events: events, updatedAt: new Date().toISOString()
  });
}

// ══════════════════════════════════════════════════════════════════════════
//  INTERVIEW SCORES  ·  external sync target (candidateKey → summary)
// ══════════════════════════════════════════════════════════════════════════
function getAllInterviewScores() {
  const rows = readAll_('InterviewScores');
  const out = {};
  rows.forEach(r => {
    if (!r.candidateKey) return;
    out[r.candidateKey] = {
      score:       r.score, scorevalue: r.score, scoretext: r.score,
      status:      r.status, difficulty: r.difficulty,
      submittedat: r.submittedAt, timestamp:  r.submittedAt,
      tabswitches: r.tabSwitches, topics: r.topics,
      raw: r.raw || null
    };
  });
  return { success: true, data: out };
}
function saveInterviewScore(candidateKey, summary) {
  if (!candidateKey) return { success: false, message: 'Missing candidateKey' };
  return upsertByKey_('InterviewScores', 'candidateKey', candidateKey, {
    candidateKey: candidateKey,
    score:       (summary && (summary.score || summary.scorevalue)) || '',
    status:      (summary && summary.status)     || '',
    difficulty:  (summary && summary.difficulty) || '',
    submittedAt: (summary && (summary.submittedat || summary.timestamp)) || new Date().toISOString(),
    tabSwitches: (summary && summary.tabswitches) || '',
    topics:      (summary && summary.topics) || '',
    raw:         summary || {}
  });
}

// ══════════════════════════════════════════════════════════════════════════
//  WHATSAPP MESSAGE  ·  optional integration via Script Properties
// ══════════════════════════════════════════════════════════════════════════
function sendWhatsAppMessage(phone, msg) {
  try {
    const props = PropertiesService.getScriptProperties();
    const url   = props.getProperty('WHATSAPP_API_URL');
    const token = props.getProperty('WHATSAPP_API_TOKEN');

    logActivity_((getCurrentUser() || {}).email || 'system', 'sendWhatsApp', 'WhatsApp', phone, { msg: msg });

    if (!url) {
      // No provider configured — record intent and succeed gracefully so the UI can proceed.
      Logger.log('[WhatsApp mock] to=%s msg=%s', phone, msg);
      return { success: true, mocked: true, message: 'Logged (no WhatsApp provider configured)' };
    }
    const payload = { phone: phone, message: msg };
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: token ? { 'Authorization': 'Bearer ' + token } : {},
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    return code >= 200 && code < 300
      ? { success: true, response: res.getContentText() }
      : { success: false, message: 'HTTP ' + code + ': ' + res.getContentText() };
  } catch (e) { return { success: false, message: e.message }; }
}

// ══════════════════════════════════════════════════════════════════════════
//  ══════════════════════════════════════════════════════════════════════
//  ↓↓↓  GENERIC HELPERS  ·  everything below is infrastructure  ↓↓↓
//  ══════════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════════

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    ensureHeaders_(sh, SCHEMAS[name] || ['id']);
  }
  return sh;
}

/** Guarantees the first row contains every header from `wanted` (appends missing). */
function ensureHeaders_(sh, wanted) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || ''));
  const present = existing.filter(Boolean);
  const missing = wanted.filter(h => present.indexOf(h) === -1);
  if (present.length === 0 && missing.length) {
    sh.getRange(1, 1, 1, missing.length).setValues([missing]);
    return;
  }
  if (missing.length) {
    sh.getRange(1, present.length + 1, 1, missing.length).setValues([missing]);
  }
}

/** Returns the current header row as string[] (auto-repairs from schema). */
function getHeaders_(sheetName) {
  const sh = getSheet_(sheetName);
  const wanted = SCHEMAS[sheetName] || [];
  if (wanted.length) ensureHeaders_(sh, wanted);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '')).filter(Boolean);
}

/** Reads all rows as an array of objects. Uses CacheService. */
function readAll_(sheetName) {
  const cache = CacheService.getScriptCache();
  const key = 'ALL::' + sheetName;
  const hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* fall through */ } }

  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) { cache.put(key, '[]', CFG.CACHE_TTL_SEC); return []; }

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(v => String(v || ''));
  const jsonCols = (JSON_FIELDS[sheetName] || []);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    // Skip fully-empty rows
    if (row.every(v => v === '' || v === null)) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;
      let v = row[j];
      if (v instanceof Date) v = v.toISOString();
      if (jsonCols.indexOf(headers[j]) !== -1 && typeof v === 'string' && v) {
        try { v = JSON.parse(v); } catch (e) { /* leave as string */ }
      }
      obj[headers[j]] = v;
    }
    rows.push(obj);
  }
  try { cache.put(key, JSON.stringify(rows), CFG.CACHE_TTL_SEC); } catch (e) { /* payload too big — skip */ }
  return rows;
}

/** Serialises a record into a row aligned with the sheet headers. */
function serializeRow_(sheetName, obj) {
  const headers = getHeaders_(sheetName);
  // Add any brand-new fields as columns
  const extra = Object.keys(obj).filter(k => headers.indexOf(k) === -1);
  if (extra.length) {
    const sh = getSheet_(sheetName);
    sh.getRange(1, headers.length + 1, 1, extra.length).setValues([extra]);
    extra.forEach(h => headers.push(h));
  }
  const jsonCols = JSON_FIELDS[sheetName] || [];
  return headers.map(h => {
    let v = obj[h];
    if (v === undefined || v === null) return '';
    if (jsonCols.indexOf(h) !== -1 && typeof v === 'object') return JSON.stringify(v);
    if (v instanceof Date) return v.toISOString();
    return v;
  });
}

function insert_(sheetName, obj) {
  const sh = getSheet_(sheetName);
  const row = serializeRow_(sheetName, obj);
  sh.appendRow(row);
  clearCacheKey_(sheetName);
  return obj;
}

function insertWithId_(sheetName, prefix, data) {
  try {
    if (!data) return { success: false, message: 'No data' };
    const lock = LockService.getScriptLock();
    lock.tryLock(CFG.LOCK_TIMEOUT_MS);
    try {
      const now = new Date().toISOString();
      if (!data.id) data.id = prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
      if (!data.createdDate && !data.createdAt) data.createdDate = now;
      insert_(sheetName, data);
      logActivity_((getCurrentUser() || {}).email || 'system', 'create', sheetName, data.id, null);
      return { success: true, id: data.id, data: data };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  } catch (e) { return { success: false, message: e.message }; }
}

function findRow_(sheetName, idField, idValue) {
  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { sheet: sh, rowIndex: -1, headers: [], row: null };
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(v => String(v || ''));
  const idx = headers.indexOf(idField);
  if (idx === -1) return { sheet: sh, rowIndex: -1, headers: headers, row: null };
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx]) === String(idValue)) {
      return { sheet: sh, rowIndex: i + 1, headers: headers, row: values[i] };
    }
  }
  return { sheet: sh, rowIndex: -1, headers: headers, row: null };
}

function updateById_(sheetName, id, updates) {
  try {
    if (!id) return { success: false, message: 'Missing id' };
    if (!updates || typeof updates !== 'object') return { success: false, message: 'Missing updates' };
    const lock = LockService.getScriptLock(); lock.tryLock(CFG.LOCK_TIMEOUT_MS);
    try {
      const loc = findRow_(sheetName, 'id', id);
      if (loc.rowIndex === -1) return { success: false, message: 'Not found' };
      // Rebuild the record object then re-serialize (auto-adds columns).
      const current = {};
      loc.headers.forEach((h, j) => { if (h) current[h] = loc.row[j]; });
      Object.assign(current, updates);
      if (!current.updatedDate && !current.updatedAt) current.updatedDate = new Date().toISOString();
      const newRow = serializeRow_(sheetName, current);
      loc.sheet.getRange(loc.rowIndex, 1, 1, newRow.length).setValues([newRow]);
      clearCacheKey_(sheetName);
      logActivity_((getCurrentUser() || {}).email || 'system', 'update', sheetName, id, Object.keys(updates));
      return { success: true, id: id };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteById_(sheetName, id) {
  try {
    if (!id) return { success: false, message: 'Missing id' };
    const lock = LockService.getScriptLock(); lock.tryLock(CFG.LOCK_TIMEOUT_MS);
    try {
      const loc = findRow_(sheetName, 'id', id);
      if (loc.rowIndex === -1) return { success: false, message: 'Not found' };
      loc.sheet.deleteRow(loc.rowIndex);
      clearCacheKey_(sheetName);
      logActivity_((getCurrentUser() || {}).email || 'system', 'delete', sheetName, id, null);
      return { success: true };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  } catch (e) { return { success: false, message: e.message }; }
}

function upsertByKey_(sheetName, keyField, keyValue, record) {
  try {
    const lock = LockService.getScriptLock(); lock.tryLock(CFG.LOCK_TIMEOUT_MS);
    try {
      const loc = findRow_(sheetName, keyField, keyValue);
      const newRow = serializeRow_(sheetName, record);
      if (loc.rowIndex === -1) {
        loc.sheet.appendRow(newRow);
      } else {
        loc.sheet.getRange(loc.rowIndex, 1, 1, newRow.length).setValues([newRow]);
      }
      clearCacheKey_(sheetName);
      return { success: true };
    } finally { try { lock.releaseLock(); } catch (e) {} }
  } catch (e) { return { success: false, message: e.message }; }
}

function okData_(rows) { return { success: true, data: rows }; }

function clearCacheKey_(sheetName) {
  try { CacheService.getScriptCache().remove('ALL::' + sheetName); } catch (e) {}
}
function clearCache_() {
  try { CacheService.getScriptCache().removeAll(Object.keys(SCHEMAS).map(s => 'ALL::' + s)); } catch (e) {}
}

function logActivity_(user, action, entity, entityId, payload) {
  try {
    const sh = getSheet_('ActivityLog');
    sh.appendRow([
      new Date().toISOString(),
      user || 'anon',
      action || '',
      entity || '',
      entityId || '',
      payload ? JSON.stringify(payload).slice(0, 5000) : ''
    ]);
  } catch (e) { /* logging is best-effort */ }
}

// ══════════════════════════════════════════════════════════════════════════
//  MAINTENANCE HELPERS (optional — run from the editor)
// ══════════════════════════════════════════════════════════════════════════
function resetCache()      { clearCache_(); return { success: true }; }
function healthCheck()     {
  const summary = {};
  Object.keys(SCHEMAS).forEach(s => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s);
    summary[s] = sh ? { rows: Math.max(0, sh.getLastRow() - 1), cols: sh.getLastColumn() } : 'missing';
  });
  return { success: true, sheets: summary, timestamp: new Date().toISOString() };
}

// ══════════════════════════════════════════════════════════════════════════
//  BULK CSV IMPORT · accepts sheet name + array of records. Batches under
//  a single Lock lease. Returns { success, ok, failed, errors:[{row, msg}] }
// ══════════════════════════════════════════════════════════════════════════
function bulkImport(sheetName, rows) {
  try {
    if (!sheetName) return { success:false, message:'Missing sheetName' };
    if (!Array.isArray(rows) || rows.length === 0) return { success:false, message:'No rows to import' };
    if (!SCHEMAS[sheetName]) return { success:false, message:'Unknown sheet: '+sheetName };
    const prefix = { Employees:'EMP', Candidates:'CAND', Requisitions:'REQ', Leaves:'LV', Holidays:'HOL', Feedback:'FB', Trainings:'TR', Problems:'PRB', Users:'USR' }[sheetName] || 'BLK';
    const lock = LockService.getScriptLock(); lock.tryLock(30000);
    const result = { success:true, ok:0, failed:0, errors:[] };
    const nowIso = new Date().toISOString();
    try {
      rows.forEach((row, i) => {
        try {
          if (!row || typeof row !== 'object') throw new Error('Not an object');
          const rec = Object.assign({}, row);
          if (!rec.id) rec.id = prefix + '-' + Utilities.getUuid().slice(0,8).toUpperCase();
          if (!rec.createdDate && !rec.createdAt) rec.createdDate = nowIso;
          insert_(sheetName, rec);
          result.ok++;
        } catch (e) {
          result.failed++;
          result.errors.push({ row: i + 1, message: String(e.message || e) });
        }
      });
      clearCacheKey_(sheetName);
      logActivity_((getCurrentUser()||{}).email||'system', 'bulkImport', sheetName, '', { ok:result.ok, failed:result.failed });
    } finally { try { lock.releaseLock(); } catch(e){} }
    return result;
  } catch (e) { return { success:false, message: e.message }; }
}

// ## 🔑 Default Login Credentials (already seeded)
