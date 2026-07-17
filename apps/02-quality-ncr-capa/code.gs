/**
 * ============================================================================
 *  QUALITY INSPECTION, NCR & CAPA MANAGEMENT SYSTEM — Google Apps Script
 * ============================================================================
 *  DEPLOYMENT
 *  ----------
 *  1. Sheets → Extensions → Apps Script → paste as code.gs; index.html as "index".
 *  2. Menu "Quality System" → "Initialize System" (idempotent).
 *  3. Deploy → New deployment → Web app.
 *
 *  SHEETS: InspectionPlans, Inspections, InspectionReadings, NCRs, CAPAs,
 *  DefectCodes, Suppliers, RepeatAlerts, Users, Config, AuditLog, JobRuns
 * ============================================================================
 */

var SHEETS = {
  InspectionPlans: ['planId','item','stage','characteristic','specMin','specMax','target','critical','samplingQty','method'],
  Inspections: ['inspId','date','stage','item','lotNo','lotQty','sampleQty','inspector','supplier','machine','result','linkedNcr'],
  InspectionReadings: ['inspId','characteristic','observed','withinSpec'],
  NCRs: ['ncrId','date','sourceInspection','item','defectCode','qtyAffected','disposition','dispositionStatus','approver1','approver2','status','costImpact','notes'],
  CAPAs: ['capaId','ncrRef','problemStatement','why1','why2','why3','why4','why5','correctiveAction','responsible','targetDate','verificationDate','verifiedBy','effectivenessResult','status','createdAt'],
  DefectCodes: ['code','description','category'],
  Suppliers: ['supplierId','name','contact'],
  RepeatAlerts: ['alertId','defectCode','item','occurrences','windowDays','firedAt','capaId','status'],
  Users: ['email','name','role','department','active','notify'],
  Config: ['key','value','description'],
  AuditLog: ['ts','user','action','entity','recordId','oldValue','newValue','screen'],
  JobRuns: ['ts','job','durationMs','outcome','detail']
};

var ROLES = ['Admin','Quality Head','Inspector','Section Head','Plant Head'];

/* ------------------------------ plumbing ----------------------------------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Quality System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Quality System')
    .addItem('Initialize System', 'setup').addItem('Reset Demo Data', 'resetDemoData')
    .addItem('Run Repeat-Defect Scan Now', 'repeatDefectScan').addToUi();
}
function ok(d){return {ok:true,data:d,error:null};}
function fail(m){return {ok:false,data:null,error:String(m)};}
function safe(fn){try{return fn();}catch(e){logJob('safeCall',0,'error',e.message);return fail(e.message);}}
function ss(){return SpreadsheetApp.getActiveSpreadsheet();}
function sheet(n){var s=ss().getSheetByName(n);if(!s)throw new Error('Sheet missing: '+n);return s;}
function headerMap(sh){var h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];var m={};h.forEach(function(c,i){m[c]=i;});return m;}
function rows(n){
  var sh=sheet(n);if(sh.getLastRow()<2)return [];
  var hm=headerMap(sh);
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().map(function(r,i){
    var o={_row:i+2};Object.keys(hm).forEach(function(k){o[k]=r[hm[k]];});return o;
  });
}
function appendObj(n,obj){
  var sh=sheet(n);var hm=headerMap(sh);
  var r=new Array(sh.getLastColumn()).fill('');
  Object.keys(obj).forEach(function(k){if(hm[k]!==undefined)r[hm[k]]=obj[k];});
  sh.appendRow(r);
}
function updateByKey(n,kc,kv,fields){
  var sh=sheet(n);var hm=headerMap(sh);
  var all=rows(n);var f=null;
  for(var i=0;i<all.length;i++)if(String(all[i][kc])===String(kv)){f=all[i];break;}
  if(!f)throw new Error(n+' not found: '+kv);
  var old={};
  Object.keys(fields).forEach(function(k){old[k]=f[k];sh.getRange(f._row,hm[k]+1).setValue(fields[k]);});
  return old;
}
function getConfig(k){var c=rows('Config').filter(function(r){return r.key===k;})[0];return c?c.value:null;}
function getConfigNum(k){return Number(getConfig(k));}
function nextId(prefix){
  var lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    var key='SEQ_'+prefix;var cur=getConfig(key);var n=cur?Number(cur)+1:1;
    var f=rows('Config').filter(function(r){return r.key===key;})[0];
    if(f)updateByKey('Config','key',key,{value:n});else appendObj('Config',{key:key,value:n,description:'ID counter'});
    return prefix+'-'+new Date().getFullYear()+'-'+('0000'+n).slice(-4);
  }finally{lock.releaseLock();}
}
function currentUser(){
  var email=Session.getActiveUser().getEmail()||'inspector@plant.local';
  var u=rows('Users').filter(function(r){return r.email===email&&r.active===true;})[0];
  return u||{email:email,name:email.split('@')[0],role:'Plant Head',department:'-',notify:false};
}
function requireRole(allowed){
  var u=currentUser();
  if(allowed.indexOf(u.role)===-1&&u.role!=='Admin')throw new Error('Permission denied for role: '+u.role);
  return u;
}
function audit(action,entity,id,oldV,newV,screen){
  appendObj('AuditLog',{ts:new Date(),user:currentUser().email,action:action,entity:entity,recordId:id,
    oldValue:JSON.stringify(oldV||{}),newValue:JSON.stringify(newV||{}),screen:screen||''});
}
function logJob(j,d,o,det){try{appendObj('JobRuns',{ts:new Date(),job:j,durationMs:d,outcome:o,detail:det||''});}catch(e){}}
function notify(to,subject,body){
  try{if(to)MailApp.sendEmail({to:to,subject:'[Quality System] '+subject,htmlBody:
    '<div style="font-family:Arial;font-size:14px;color:#1a1a18"><p>'+body+'</p><p style="color:#6b6b66;font-size:12px">Quality Inspection, NCR & CAPA System</p></div>'});}
  catch(e){logJob('notify',0,'error',e.message);}
}
function fmtD(d){if(!d)return '';if(typeof d==='string'&&d.match(/^\d{4}-\d{2}-\d{2}/))return d.slice(0,10);
  return Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'yyyy-MM-dd');}

/* --------------------------------- setup ----------------------------------- */

function setup(){
  var s=ss();
  Object.keys(SHEETS).forEach(function(n){
    var sh=s.getSheetByName(n);
    if(!sh){sh=s.insertSheet(n);
      sh.getRange(1,1,1,SHEETS[n].length).setValues([SHEETS[n]]).setFontWeight('bold').setBackground('#f1f1ef');
      sh.setFrozenRows(1);}
  });
  seedConfig();seedDemoData();installTriggers();
  return ok('Setup complete');
}

function seedConfig(){
  var defs=[
    ['APP_VERSION','v1.0','Version'],['BUILD_DATE',new Date().toDateString(),'Build date'],
    ['REPEAT_DEFECT_COUNT','3','Repeat alert: same defect+item >= this many times'],
    ['REPEAT_DEFECT_DAYS','90','...within this many days'],
    ['DISPOSITION_ESCALATE_HOURS','48','Pending dispositions older than this escalate'],
    ['EFFECTIVENESS_CHECK_DAYS','30','Days after CAPA closure before effectiveness verification task'],
    ['REJECT_RATE_PER_UNIT','120','Cost of a rejected unit (₹) for COPQ'],
    ['REWORK_RATE_PER_UNIT','45','Cost of a reworked unit (₹) for COPQ'],
    ['QUALITY_HEAD_EMAIL','qualityhead@plant.local','Escalations'],
    ['PLANT_HEAD_EMAIL','planthead@plant.local','Escalations + use-as-is approvals'],
    ['DISPOSITION_MATRIX',JSON.stringify({
      'rework':['Section Head'],'reject':['Quality Head'],
      'use-as-is':['Quality Head','Plant Head'],'return-to-supplier':['Quality Head']
    }),'Approvers required per disposition']
  ];
  var ex=rows('Config').map(function(r){return r.key;});
  defs.forEach(function(d){if(ex.indexOf(d[0])===-1)appendObj('Config',{key:d[0],value:d[1],description:d[2]});});
}

function resetDemoData(){
  ['InspectionPlans','Inspections','InspectionReadings','NCRs','CAPAs','DefectCodes','Suppliers','RepeatAlerts','Users'].forEach(function(n){
    var sh=sheet(n);if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  });
  seedDemoData();return ok('Demo data reset');
}

function seedDemoData(){
  if(rows('Suppliers').length>0)return;
  [['SUP-01','Precision Castings Ltd','pc@sup.local'],['SUP-02','Metro Forgings','mf@sup.local']]
    .forEach(function(s){appendObj('Suppliers',{supplierId:s[0],name:s[1],contact:s[2]});});
  [['DC-01','Dimensional out of tolerance','Dimensional'],['DC-02','Surface crack','Material'],
   ['DC-03','Porosity / blow hole','Casting'],['DC-04','Plating peel-off','Finish'],
   ['DC-05','Hardness below spec','Heat Treatment'],['DC-06','Burr / sharp edge','Machining']]
    .forEach(function(d){appendObj('DefectCodes',{code:d[0],description:d[1],category:d[2]});});
  [['admin@plant.local','Asha Admin','Admin','IT',true,true],
   ['qualityhead@plant.local','Qadir QHead','Quality Head','Quality',true,true],
   ['inspector@plant.local','Indu Inspector','Inspector','Quality',true,false],
   ['sectionhead@plant.local','Suresh Section','Section Head','Machining',true,true],
   ['planthead@plant.local','Priya PlantHead','Plant Head','Management',true,true]]
    .forEach(function(u){appendObj('Users',{email:u[0],name:u[1],role:u[2],department:u[3],active:u[4],notify:u[5]});});

  // inspection plans
  var plans=[
    ['IP-01','ITEM-100','incoming','Diameter (mm)',24.95,25.05,25,true,5,'Micrometer'],
    ['IP-02','ITEM-100','incoming','Hardness (HRC)',40,45,42,false,3,'Rockwell'],
    ['IP-03','ITEM-100','in-process','Length (mm)',99.9,100.1,100,true,5,'Vernier'],
    ['IP-04','ITEM-100','final','Surface finish (Ra)',0,1.6,0.8,false,3,'Roughness tester'],
    ['IP-05','ITEM-200','incoming','Thickness (mm)',3.95,4.05,4,true,5,'Micrometer'],
    ['IP-06','ITEM-200','final','Coating (µm)',18,25,20,true,3,'Coating gauge']
  ];
  plans.forEach(function(p){appendObj('InspectionPlans',{planId:p[0],item:p[1],stage:p[2],characteristic:p[3],specMin:p[4],specMax:p[5],target:p[6],critical:p[7],samplingQty:p[8],method:p[9]});});

  // Inspections across 90 days: pass/fail all stages, supplier variance
  var today=new Date();var n=0;
  var stages=['incoming','in-process','final'];
  for(var d=90;d>=0;d-=3){
    var date=new Date(today.getTime()-d*86400000);
    stages.forEach(function(st,si){
      n++;
      var item=(n%2===0)?'ITEM-100':'ITEM-200';
      var sup=(st==='incoming')?((n%3===0)?'SUP-02':'SUP-01'):'';
      // SUP-02 fails more often → visibly different scorecards
      var failIt=(st==='incoming'&&sup==='SUP-02'&&n%2===0)||(n%11===0);
      var inspId='INS-'+fmtD(date).replace(/-/g,'')+'-'+si+n;
      appendObj('Inspections',{inspId:inspId,date:fmtD(date),stage:st,item:item,lotNo:'LOT-'+(1000+n),
        lotQty:500,sampleQty:5,inspector:'inspector@plant.local',supplier:sup,machine:st==='in-process'?'M-0'+((n%4)+1):'',
        result:failIt?'fail':'pass',linkedNcr:''});
      appendObj('InspectionReadings',{inspId:inspId,characteristic:'Primary characteristic',
        observed:failIt?'out of spec':'within spec',withinSpec:!failIt});
      if(failIt){
        var ncrId='NCR-SEED-'+n;
        var dispositions=['rework','reject','use-as-is','return-to-supplier'];
        var disp=dispositions[n%4];
        appendObj('NCRs',{ncrId:ncrId,date:fmtD(date),sourceInspection:inspId,item:item,
          defectCode:'DC-0'+((n%6)+1),qtyAffected:40+(n%60),disposition:disp,
          dispositionStatus:(d<4&&n%4===2)?'pending':'approved',
          approver1:'qualityhead@plant.local',approver2:disp==='use-as-is'?'planthead@plant.local':'',
          status:(d<10)?'open':'closed',costImpact:(40+(n%60))*(disp==='reject'?120:45),notes:'Seeded demo NCR'});
        updateByKey('Inspections','inspId',inspId,{linkedNcr:ncrId});
      }
    });
  }
  // CAPAs: one overdue, one closed pending effectiveness, one complete
  appendObj('CAPAs',{capaId:'CAPA-2026-0001',ncrRef:'NCR-SEED-6',problemStatement:'Recurring porosity on ITEM-100 castings from SUP-02',
    why1:'Porosity found at incoming','why2':'Gas entrapment during pour',why3:'Mould venting inadequate',
    why4:'Vent design not updated after pattern change',why5:'No design-review checkpoint on pattern changes',
    correctiveAction:'Add vent-design review to supplier pattern-change checklist; audit SUP-02 process',
    responsible:'qualityhead@plant.local',targetDate:fmtD(new Date(today.getTime()-10*86400000)),
    verificationDate:'',verifiedBy:'',effectivenessResult:'',status:'action-overdue',createdAt:fmtD(new Date(today.getTime()-40*86400000))});
  appendObj('CAPAs',{capaId:'CAPA-2026-0002',ncrRef:'NCR-SEED-9',problemStatement:'Plating peel-off on ITEM-200 final inspection',
    why1:'Peel-off at final QC',why2:'Poor adhesion',why3:'Surface prep bath contaminated',why4:'Bath change frequency not defined',
    why5:'No preventive schedule for plating line consumables',
    correctiveAction:'Define bath-change SOP with 50-lot frequency; add to PM calendar',
    responsible:'sectionhead@plant.local',targetDate:fmtD(new Date(today.getTime()-20*86400000)),
    verificationDate:fmtD(new Date(today.getTime()+10*86400000)),verifiedBy:'',effectivenessResult:'',
    status:'closed-pending-verification',createdAt:fmtD(new Date(today.getTime()-55*86400000))});
  appendObj('RepeatAlerts',{alertId:'RA-2026-0001',defectCode:'DC-03',item:'ITEM-100',occurrences:4,windowDays:90,
    firedAt:fmtD(new Date(today.getTime()-12*86400000)),capaId:'CAPA-2026-0001',status:'capa-linked'});
}

function installTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('repeatDefectScan').timeBased().everyDays(1).atHour(2).create();
  ScriptApp.newTrigger('escalationScan').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('sendDailyDigest').timeBased().everyDays(1).atHour(7).create();
}

/* ------------------------------- bootstrap --------------------------------- */

function bootstrap(){
  return safe(function(){
    var cfg={};rows('Config').forEach(function(r){cfg[r.key]=r.value;});
    return ok({me:currentUser(),config:cfg,
      defectCodes:rows('DefectCodes'),suppliers:rows('Suppliers'),
      items:uniq(rows('InspectionPlans').map(function(p){return p.item;})),
      dispositionMatrix:JSON.parse(getConfig('DISPOSITION_MATRIX')||'{}')});
  });
}
function uniq(a){return a.filter(function(v,i){return a.indexOf(v)===i;});}

/* ---------------------------- inspections ----------------------------------- */

function getInspectionPlan(item,stage){
  return safe(function(){
    return ok(rows('InspectionPlans').filter(function(p){return p.item===item&&p.stage===stage;}));
  });
}

function submitInspection(payload){
  return safe(function(){
    requireRole(['Inspector','Quality Head']);
    var lock=LockService.getScriptLock();lock.waitLock(10000);
    try{
      var inspId=nextId('INS');
      var anyCriticalFail=false,anyFail=false;
      payload.readings.forEach(function(r){
        var within=Number(r.observed)>=Number(r.specMin)&&Number(r.observed)<=Number(r.specMax);
        if(!within){anyFail=true;if(r.critical)anyCriticalFail=true;}
        appendObj('InspectionReadings',{inspId:inspId,characteristic:r.characteristic,observed:r.observed,withinSpec:within});
      });
      var result=anyCriticalFail?'fail':(anyFail?'conditional':'pass');
      appendObj('Inspections',{inspId:inspId,date:fmtD(new Date()),stage:payload.stage,item:payload.item,
        lotNo:payload.lotNo,lotQty:Number(payload.lotQty),sampleQty:payload.readings.length,
        inspector:currentUser().email,supplier:payload.supplier||'',machine:payload.machine||'',
        result:result,linkedNcr:''});
      audit('CREATE','Inspections',inspId,null,{result:result},'Inspection');
      return ok({inspId:inspId,result:result,requiresNcr:result==='fail'});
    }finally{lock.releaseLock();}
  });
}

function createNcr(p){
  return safe(function(){
    requireRole(['Inspector','Quality Head']);
    var lock=LockService.getScriptLock();lock.waitLock(10000);
    try{
      var ncrId=nextId('NCR');
      var rejectRate=getConfigNum('REJECT_RATE_PER_UNIT'),reworkRate=getConfigNum('REWORK_RATE_PER_UNIT');
      var cost=Number(p.qtyAffected)*(p.disposition==='reject'?rejectRate:reworkRate);
      appendObj('NCRs',{ncrId:ncrId,date:fmtD(new Date()),sourceInspection:p.sourceInspection,item:p.item,
        defectCode:p.defectCode,qtyAffected:Number(p.qtyAffected),disposition:p.disposition,
        dispositionStatus:'pending',approver1:'',approver2:'',status:'open',costImpact:cost,notes:p.notes||''});
      if(p.sourceInspection)updateByKey('Inspections','inspId',p.sourceInspection,{linkedNcr:ncrId});
      audit('CREATE','NCRs',ncrId,null,p,'NCR');
      notify(getConfig('QUALITY_HEAD_EMAIL'),'NCR '+ncrId+' raised — lot on Quality Hold',
        'Item '+p.item+', defect '+p.defectCode+', qty '+p.qtyAffected+'. Disposition "'+p.disposition+'" awaits approval.');
      return ok({ncrId:ncrId});
    }finally{lock.releaseLock();}
  });
}

function approveDisposition(ncrId){
  return safe(function(){
    var u=currentUser();
    var ncr=rows('NCRs').filter(function(r){return r.ncrId===ncrId;})[0];
    if(!ncr)return fail('NCR not found');
    if(ncr.dispositionStatus==='approved')return fail('Already fully approved');
    var matrix=JSON.parse(getConfig('DISPOSITION_MATRIX')||'{}');
    var needed=matrix[ncr.disposition]||['Quality Head'];
    if(needed.indexOf(u.role)===-1&&u.role!=='Admin')return fail('Role '+u.role+' cannot approve "'+ncr.disposition+'". Needs: '+needed.join(' + '));
    var lock=LockService.getScriptLock();lock.waitLock(10000);
    try{
      var fields={};
      if(!ncr.approver1)fields.approver1=u.email;
      else if(!ncr.approver2&&ncr.approver1!==u.email)fields.approver2=u.email;
      else if(ncr.approver1===u.email)return fail('You have already approved this NCR');
      var approversAfter=(ncr.approver1?1:0)+(fields.approver1?1:0)+(ncr.approver2?1:0)+(fields.approver2?1:0);
      if(approversAfter>=needed.length)fields.dispositionStatus='approved';
      var old=updateByKey('NCRs','ncrId',ncrId,fields);
      audit('APPROVE','NCRs',ncrId,old,fields,'Disposition');
      return ok({fullyApproved:fields.dispositionStatus==='approved',pendingApprovals:Math.max(0,needed.length-approversAfter)});
    }finally{lock.releaseLock();}
  });
}

function closeNcr(ncrId){
  return safe(function(){
    requireRole(['Quality Head']);
    var ncr=rows('NCRs').filter(function(r){return r.ncrId===ncrId;})[0];
    if(!ncr)return fail('NCR not found');
    if(ncr.dispositionStatus!=='approved')return fail('Disposition must be approved before closure');
    var old=updateByKey('NCRs','ncrId',ncrId,{status:'closed'});
    audit('CLOSE','NCRs',ncrId,old,{status:'closed'},'NCR');
    return ok(true);
  });
}

function getRegister(entity,filters){
  return safe(function(){
    var data=rows(entity);
    if(filters){
      if(filters.q){var q=filters.q.toLowerCase();
        data=data.filter(function(r){return JSON.stringify(r).toLowerCase().indexOf(q)>-1;});}
      if(filters.status)data=data.filter(function(r){return String(r.status)===filters.status||String(r.result)===filters.status;});
    }
    return ok(data.slice(-300).reverse());
  });
}

/* -------------------------------- CAPA -------------------------------------- */

function saveCapa(p){
  return safe(function(){
    requireRole(['Quality Head','Section Head']);
    var lock=LockService.getScriptLock();lock.waitLock(10000);
    try{
      var whys=[p.why1,p.why2,p.why3,p.why4,p.why5];
      var allFilled=whys.every(function(w){return w&&String(w).trim().length>0;});
      var status=p.status||'draft';
      if(status==='action-defined'&&!allFilled)return fail('CAPA cannot move to Action Defined until all five Why fields are filled (or marked "N/A — <reason>").');
      var capaId=p.capaId||nextId('CAPA');
      if(p.capaId){
        var old=updateByKey('CAPAs','capaId',capaId,{problemStatement:p.problemStatement,why1:p.why1,why2:p.why2,why3:p.why3,why4:p.why4,why5:p.why5,
          correctiveAction:p.correctiveAction,responsible:p.responsible,targetDate:p.targetDate,status:status});
        audit('UPDATE','CAPAs',capaId,old,p,'CAPA');
      } else {
        appendObj('CAPAs',{capaId:capaId,ncrRef:p.ncrRef,problemStatement:p.problemStatement,
          why1:p.why1||'',why2:p.why2||'',why3:p.why3||'',why4:p.why4||'',why5:p.why5||'',
          correctiveAction:p.correctiveAction||'',responsible:p.responsible||'',targetDate:p.targetDate||'',
          verificationDate:'',verifiedBy:'',effectivenessResult:'',status:status,createdAt:fmtD(new Date())});
        audit('CREATE','CAPAs',capaId,null,p,'CAPA');
      }
      return ok({capaId:capaId});
    }finally{lock.releaseLock();}
  });
}

function closeCapa(capaId){
  return safe(function(){
    requireRole(['Quality Head']);
    var days=getConfigNum('EFFECTIVENESS_CHECK_DAYS');
    var vDate=fmtD(new Date(Date.now()+days*86400000));
    var old=updateByKey('CAPAs','capaId',capaId,{status:'closed-pending-verification',verificationDate:vDate});
    audit('CLOSE','CAPAs',capaId,old,{verificationDate:vDate},'CAPA');
    return ok({verificationDate:vDate});
  });
}

function verifyCapa(capaId,effective){
  return safe(function(){
    requireRole(['Quality Head']);
    var old=updateByKey('CAPAs','capaId',capaId,{
      status:effective?'closed-effective':'reopened-ineffective',
      verifiedBy:currentUser().email,effectivenessResult:effective?'effective':'not effective'});
    audit('VERIFY','CAPAs',capaId,old,{effective:effective},'CAPA');
    return ok(true);
  });
}

/* --------------------------- engines & analytics ----------------------------- */

function repeatDefectScan(){
  var t0=Date.now();
  try{
    var X=getConfigNum('REPEAT_DEFECT_COUNT'),Y=getConfigNum('REPEAT_DEFECT_DAYS');
    var cutoff=Date.now()-Y*86400000;
    var ncrs=rows('NCRs').filter(function(n){return new Date(n.date).getTime()>=cutoff;});
    var combos={};
    ncrs.forEach(function(n){var k=n.defectCode+'|'+n.item;combos[k]=(combos[k]||0)+1;});
    var existing=rows('RepeatAlerts').map(function(a){return a.defectCode+'|'+a.item;});
    var fired=0;
    Object.keys(combos).forEach(function(k){
      if(combos[k]>=X&&existing.indexOf(k)===-1){
        var parts=k.split('|');
        appendObj('RepeatAlerts',{alertId:nextId('RA'),defectCode:parts[0],item:parts[1],
          occurrences:combos[k],windowDays:Y,firedAt:fmtD(new Date()),capaId:'',status:'open-capa-required'});
        notify(getConfig('QUALITY_HEAD_EMAIL'),'REPEAT DEFECT: '+parts[0]+' on '+parts[1],
          parts[0]+' occurred '+combos[k]+' times in '+Y+' days on '+parts[1]+'. A mandatory CAPA is required.');
        fired++;
      }
    });
    logJob('repeatDefectScan',Date.now()-t0,'ok',fired+' alerts fired');
  }catch(e){logJob('repeatDefectScan',Date.now()-t0,'error',e.message);}
}

function escalationScan(){
  var t0=Date.now();
  try{
    var hrs=getConfigNum('DISPOSITION_ESCALATE_HOURS');
    var pend=rows('NCRs').filter(function(n){
      return n.dispositionStatus==='pending'&&(Date.now()-new Date(n.date).getTime())>hrs*3600000;});
    pend.forEach(function(n){
      notify(getConfig('PLANT_HEAD_EMAIL'),'ESCALATION: disposition pending >'+hrs+'h on '+n.ncrId,
        'NCR '+n.ncrId+' ('+n.item+', '+n.defectCode+', qty '+n.qtyAffected+') has awaited disposition approval for over '+hrs+' hours.');
    });
    var overdue=rows('CAPAs').filter(function(c){
      return c.status!=='closed-effective'&&c.status!=='closed-pending-verification'&&c.targetDate&&new Date(c.targetDate).getTime()<Date.now();});
    overdue.forEach(function(c){
      if(c.status!=='action-overdue')updateByKey('CAPAs','capaId',c.capaId,{status:'action-overdue'});
      notify(getConfig('QUALITY_HEAD_EMAIL'),'CAPA overdue: '+c.capaId,'Corrective action past target date '+fmtD(c.targetDate)+'. Responsible: '+c.responsible);
    });
    logJob('escalationScan',Date.now()-t0,'ok',pend.length+' dispositions, '+overdue.length+' CAPAs');
  }catch(e){logJob('escalationScan',Date.now()-t0,'error',e.message);}
}

function sendDailyDigest(){
  var t0=Date.now();
  try{
    var d=getDashboard().data;
    var html='<div style="font-family:Arial;font-size:14px"><h3>Quality daily digest</h3><ul>'+
      '<li>Open NCRs: '+d.openNcrs+'</li><li>Pending dispositions: '+d.pendingDispositions+'</li>'+
      '<li>Overdue CAPAs: '+d.overdueCapas+'</li><li>Repeat-defect alerts open: '+d.openRepeatAlerts+'</li>'+
      '<li>COPQ this month: ₹'+d.copqMonth.toLocaleString()+'</li></ul></div>';
    MailApp.sendEmail({to:getConfig('QUALITY_HEAD_EMAIL'),subject:'[Quality System] Daily digest',htmlBody:html});
    logJob('sendDailyDigest',Date.now()-t0,'ok','');
  }catch(e){logJob('sendDailyDigest',Date.now()-t0,'error',e.message);}
}

function getDashboard(){
  return safe(function(){
    var ncrs=rows('NCRs'),capas=rows('CAPAs'),insps=rows('Inspections');
    var monthStart=new Date();monthStart.setDate(1);
    var copqMonth=ncrs.filter(function(n){return new Date(n.date)>=monthStart;})
      .reduce(function(s,n){return s+Number(n.costImpact||0);},0);
    var last30=insps.filter(function(i){return(Date.now()-new Date(i.date).getTime())<30*86400000;});
    var passRate=last30.length?Math.round(last30.filter(function(i){return i.result==='pass';}).length/last30.length*100):0;
    return ok({
      openNcrs:ncrs.filter(function(n){return n.status==='open';}).length,
      pendingDispositions:ncrs.filter(function(n){return n.dispositionStatus==='pending';}).length,
      overdueCapas:capas.filter(function(c){return c.status==='action-overdue';}).length,
      openRepeatAlerts:rows('RepeatAlerts').filter(function(a){return a.status==='open-capa-required';}).length,
      copqMonth:copqMonth,passRate30:passRate,
      attention:ncrs.filter(function(n){return n.dispositionStatus==='pending'||n.status==='open';}).slice(-10).reverse(),
      repeatAlerts:rows('RepeatAlerts').slice(-10).reverse()
    });
  });
}

function getSupplierScorecard(){
  return safe(function(){
    var incoming=rows('Inspections').filter(function(i){return i.stage==='incoming'&&i.supplier;});
    var ncrs=rows('NCRs');
    var bySup={};
    incoming.forEach(function(i){
      if(!bySup[i.supplier])bySup[i.supplier]={lots:0,accepted:0,sampleUnits:0,defectUnits:0,defects:{}};
      var s=bySup[i.supplier];s.lots++;
      if(i.result==='pass')s.accepted++;
      s.sampleUnits+=Number(i.sampleQty||0);
      if(i.linkedNcr){
        var ncr=ncrs.filter(function(n){return n.ncrId===i.linkedNcr;})[0];
        if(ncr){s.defectUnits+=Number(ncr.qtyAffected||0);s.defects[ncr.defectCode]=(s.defects[ncr.defectCode]||0)+1;}
      }
    });
    var sups=rows('Suppliers');
    return ok(Object.keys(bySup).map(function(id){
      var s=bySup[id];var sup=sups.filter(function(x){return x.supplierId===id;})[0];
      var totalUnits=s.lots*500; // lotQty seeded 500
      return {supplierId:id,name:sup?sup.name:id,lots:s.lots,
        acceptanceRate:Math.round(s.accepted/s.lots*100),
        ppm:totalUnits?Math.round(s.defectUnits/totalUnits*1000000):0,
        topDefects:Object.keys(s.defects).sort(function(a,b){return s.defects[b]-s.defects[a];}).slice(0,3)};
    }));
  });
}

function getCopq(){
  return safe(function(){
    var ncrs=rows('NCRs');
    var by=function(key){
      var m={};ncrs.forEach(function(n){var k=n[key]||'—';m[k]=(m[k]||0)+Number(n.costImpact||0);});
      return Object.keys(m).map(function(k){return {key:k,cost:Math.round(m[k])};})
        .sort(function(a,b){return b.cost-a.cost;}).slice(0,8);
    };
    var byMonth={};
    ncrs.forEach(function(n){var k=fmtD(n.date).slice(0,7);byMonth[k]=(byMonth[k]||0)+Number(n.costImpact||0);});
    return ok({byDefect:by('defectCode'),
      byMonth:Object.keys(byMonth).sort().map(function(k){return {key:k,cost:Math.round(byMonth[k])};}),
      total:Math.round(ncrs.reduce(function(s,n){return s+Number(n.costImpact||0);},0))});
  });
}

function getAuditBinder(fromDate,toDate){
  return safe(function(){
    requireRole(['Quality Head','Plant Head','Admin']);
    var inRange=function(d){var t=new Date(d).getTime();return t>=new Date(fromDate).getTime()&&t<=new Date(toDate).getTime()+86399000;};
    return ok({
      inspections:rows('Inspections').filter(function(i){return inRange(i.date);}),
      ncrs:rows('NCRs').filter(function(n){return inRange(n.date);}),
      capas:rows('CAPAs').filter(function(c){return inRange(c.createdAt);}),
      generatedAt:new Date(),range:{from:fromDate,to:toDate}
    });
  });
}

/* ---------------------------------- about ------------------------------------ */

function getAboutData(){
  return safe(function(){
    var cfg={};rows('Config').forEach(function(r){cfg[r.key]=r.value;});
    return ok({version:cfg.APP_VERSION,buildDate:cfg.BUILD_DATE,config:cfg,
      sheets:Object.keys(SHEETS).map(function(n){return {name:n,columns:SHEETS[n]};}),
      roles:ROLES,lastVerified:new Date()});
  });
}
