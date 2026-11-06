/** MYTASK — Google Apps Script backend (V8 runtime).
 * Script Properties: SPREADSHEET_ID, FIREBASE_API_KEY, ALLOWED_ORIGINS
 * ALLOWED_ORIGINS = comma separated fully qualified origins (e.g. https://example.vercel.app,http://localhost:5500).
 * Deployment: Web app / Execute as me / Anyone. This is safe only because EVERY
 * protected operation independently checks the Firebase token & row ownership.
 */
const MT = {
  schema: {
    Users:['id','uid','email','name','description','avatar','createdAt','updatedAt'],
    Tasks:['id','uid','title','description','status','priority','startDate','dueDate','dueTime','projectId','categoryId','tags','notes','progress','archived','deletedAt','createdAt','updatedAt','completedAt','version'],
    Projects:['id','uid','name','description','status','priority','startDate','dueDate','color','archived','deletedAt','createdAt','updatedAt','version'],
    Categories:['id','uid','name','color','createdAt','updatedAt','deletedAt','version'],
    Tags:['id','uid','name','createdAt','updatedAt','deletedAt','version'],
    TaskTags:['id','uid','taskId','tagId','createdAt'],
    Subtasks:['id','uid','taskId','title','completed','createdAt','updatedAt','deletedAt','version'],
    RecurringRules:['id','uid','taskId','frequency','interval','weekdays','startDate','endDate','nextDate','active','createdAt','updatedAt','version'],
    Reminders:['id','uid','taskId','message','remindAt','read','snoozedUntil','createdAt','updatedAt','deletedAt','version'],
    FocusSessions:['id','uid','taskId','durationSeconds','completed','startedAt','endedAt','createdAt'],
    ActivityLogs:['id','uid','action','entity','entityId','details','createdAt'],
    Preferences:['id','uid','timezone','dateFormat','defaultPriority','defaultView','focusMinutes','shortBreak','longBreak','notifyDue','createdAt','updatedAt','version'],
    APIRequests:['id','uid','operation','response','createdAt']
  },
  allowed: {
    Tasks:['title','description','status','priority','startDate','dueDate','dueTime','projectId','categoryId','tags','notes','progress','archived'],
    Projects:['name','description','status','priority','startDate','dueDate','color','archived'],
    Categories:['name','color'], Tags:['name'],
    Subtasks:['taskId','title','completed'],
    RecurringRules:['taskId','frequency','interval','weekdays','startDate','endDate','nextDate','active'],
    Reminders:['taskId','message','remindAt','read','snoozedUntil'],
    Preferences:['timezone','dateFormat','defaultPriority','defaultView','focusMinutes','shortBreak','longBreak','notifyDue'],
    Users:['name','description','avatar']
  }
};
// Execution-scoped spreadsheet handle removes repeated openById() calls.
let mtSpreadsheetHandle_ = null;
let mtSheetHandles_ = {};
function mtProps_(){return PropertiesService.getScriptProperties();}
function mtCache_(){return CacheService.getScriptCache();}
function mtSnapKey_(uid){return 'mt-snap-v2-'+uid;}
function mtInvalidate_(uid){try{mtCache_().remove(mtSnapKey_(uid));}catch(e){console.warn('Cache invalidation:',e.message)}}
function mtSnapCached_(uid){
 try{const hit=mtCache_().get(mtSnapKey_(uid));if(hit){return JSON.parse(hit)}}catch(e){console.warn('Cache read:',e.message)}
 const data=mtSnapshot_(uid);
 try{const serialized=JSON.stringify(data);if(serialized.length<85000)mtCache_().put(mtSnapKey_(uid),serialized,25)}catch(e){console.warn('Cache write:',e.message)}
 return data;
}
function mtNow_(){return new Date().toISOString();}
function mtId_(){return Utilities.getUuid();}
function mtJson_(v){return JSON.stringify(v).replace(/</g,'\\u003c').replace(/>/g,'\\u003e').replace(/&/g,'\\u0026').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');}
function mtErr_(code,message){const e=new Error(message);e.code=code;throw e;}
function mtSpreadsheet_(){const id=mtProps_().getProperty('SPREADSHEET_ID');if(!id) mtErr_('CONFIG','Set SPREADSHEET_ID in Script Properties.');return mtSpreadsheetHandle_ || (mtSpreadsheetHandle_=SpreadsheetApp.openById(id));}
function setupMyTaskDatabase(){
  const ss=mtSpreadsheet_();
  Object.keys(MT.schema).forEach(name=>{
    let sheet=ss.getSheetByName(name);if(!sheet)sheet=ss.insertSheet(name);
    const headers=MT.schema[name];
    if(sheet.getLastRow()===0){sheet.getRange(1,1,1,headers.length).setValues([headers]);sheet.setFrozenRows(1);sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#FFE4F0');}
    else {const found=sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getDisplayValues()[0];headers.forEach((h,i)=>{if(found[i]!==h)mtErr_('SCHEMA','Sheet '+name+' has unexpected column '+(i+1)+'. Back up and reconcile schema without deleting data.');});}
  });
  return 'MyTask database initialized: '+Object.keys(MT.schema).length+' sheets.';
}
function mtSheet_(name){if(!MT.schema[name])mtErr_('BAD_TABLE','Invalid table.');const s=mtSheetHandles_[name]||(mtSheetHandles_[name]=mtSpreadsheet_().getSheetByName(name));if(!s)mtErr_('SCHEMA','Missing sheet '+name+'. Run setupMyTaskDatabase.');return s;}
function mtList_(name,uid,all){
  const s=mtSheet_(name), n=s.getLastRow(); if(n<2)return [];
  const heads=MT.schema[name], rows=s.getRange(2,1,n-1,heads.length).getValues();
  return rows.map((r,i)=>{
    const obj={_row:i+2};heads.forEach((k,j)=>obj[k]=r[j] instanceof Date?Utilities.formatDate(r[j],'GMT',"yyyy-MM-dd'T'HH:mm:ss'Z'"):r[j]);return obj;
  }).filter(x=>x.uid===uid && (all || !x.deletedAt));
}
function mtWrite_(name,obj,row){
  const s=mtSheet_(name), keys=MT.schema[name], values=keys.map(k=>obj[k]===undefined?'':obj[k]);
  const target=row||s.getLastRow()+1;s.getRange(target,1,1,values.length).setValues([values]);return obj;
}
function mtFind_(name,uid,id){const row=mtList_(name,uid,true).find(x=>x.id===id);if(!row)mtErr_('NOT_FOUND','Item tidak ditemukan atau tidak memiliki izin akses.');return row;}
function mtClean_(obj){const ret={...obj};delete ret._row;return ret;}
function mtValidate_(name,data,partial){
  if(!data||typeof data!=='object'||Array.isArray(data))mtErr_('INPUT','Payload harus berupa object.');
  const allowed=MT.allowed[name];if(!allowed)mtErr_('INPUT','Unknown entity.');
  const out={};Object.keys(data).forEach(k=>{if(!allowed.includes(k))mtErr_('INPUT','Field tidak diizinkan: '+k);if(typeof data[k]==='object'&&data[k]!==null)mtErr_('INPUT','Nested value not allowed: '+k);const val=String(data[k]??'');if(val.length>5000)mtErr_('INPUT','Value too long: '+k);out[k]=data[k]===null?'':data[k];});
  if(!partial){const title=name==='Tasks'?'title':name==='Projects'||name==='Categories'||name==='Tags'?'name':name==='Subtasks'?'title':null;if(title&&!String(out[title]||'').trim())mtErr_('INPUT',title+' wajib diisi.');}
  if(out.status&&name==='Tasks'&&!['To Do','In Progress','On Hold','Completed','Cancelled'].includes(out.status))mtErr_('INPUT','Invalid task status.');
  if(out.priority&&!['Critical','High','Medium','Low'].includes(out.priority))mtErr_('INPUT','Invalid priority.');
  ['dueDate','startDate','endDate','nextDate'].forEach(k=>{if(out[k]&&!/^\d{4}-\d{2}-\d{2}$/.test(out[k]))mtErr_('INPUT',k+' must be YYYY-MM-DD.');});
  return out;
}
function mtVerify_(token){
  if(!token||token.length>8000)mtErr_('AUTH','Firebase ID token missing.');
  const key=mtProps_().getProperty('FIREBASE_API_KEY');if(!key)mtErr_('CONFIG','Missing FIREBASE_API_KEY.');
  const url='https://identitytoolkit.googleapis.com/v1/accounts:lookup?key='+encodeURIComponent(key);
  const response=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',payload:JSON.stringify({idToken:token}),muteHttpExceptions:true});
  if(response.getResponseCode()!==200)mtErr_('AUTH','Session invalid or expired. Please sign in again.');
  const parsed=JSON.parse(response.getContentText());const u=parsed.users&&parsed.users[0];
  if(!u||!u.localId||u.disabled||!u.emailVerified)mtErr_('UNVERIFIED','Email verification required.');
  return {uid:u.localId,email:u.email||'',name:u.displayName||''};
}
function mtActivity_(uid,action,entity,entityId){mtWrite_('ActivityLogs',{id:mtId_(),uid,action,entity,entityId,details:'',createdAt:mtNow_()});}
function mtWithLock_(fn){const lock=LockService.getScriptLock();if(!lock.tryLock(10000))mtErr_('BUSY','Server sibuk, silakan coba lagi.');try{return fn();}finally{lock.releaseLock();}}
function mtCreate_(table,uid,data){
  const fields=mtValidate_(table,data,false),now=mtNow_();
  if(fields.projectId)mtFind_('Projects',uid,fields.projectId);
  if(fields.taskId)mtFind_('Tasks',uid,fields.taskId);
  const row={id:mtId_(),uid,...fields,createdAt:now,updatedAt:now,version:1,deletedAt:'',archived:false};
  if(table==='Tasks'){row.status=row.status||'To Do';row.priority=row.priority||'Medium';row.progress=row.status==='Completed'?100:Number(row.progress)||0;row.completedAt=row.status==='Completed'?now:'';}
  if(table==='Projects'){row.status=row.status||'Active';row.priority=row.priority||'Medium';row.color=row.color||'#FF4F9A';}
  if(table==='Reminders'){row.read=false;row.snoozedUntil='';}
  if(table==='RecurringRules'){row.active=true;row.interval=Math.max(1,Number(row.interval)||1);row.nextDate=row.startDate||new Date().toISOString().slice(0,10);}
  mtWrite_(table,row);mtActivity_(uid,'created',table,row.id);mtInvalidate_(uid);return mtClean_(row);
}
function mtUpdate_(table,uid,id,data){
  const old=mtFind_(table,uid,id),fields=mtValidate_(table,data,true);
  if(fields.projectId)mtFind_('Projects',uid,fields.projectId);
  if(fields.taskId)mtFind_('Tasks',uid,fields.taskId);
  if(fields.version!==undefined)delete fields.version;
  const updated={...old,...fields,updatedAt:mtNow_(),version:Number(old.version||0)+1};
  if(table==='Tasks'&&fields.status!==undefined){updated.completedAt=fields.status==='Completed'?(old.completedAt||mtNow_()):'';updated.progress=fields.status==='Completed'?100:Number(updated.progress)||0;}
  mtWrite_(table,updated,old._row);mtActivity_(uid,'updated',table,id);mtInvalidate_(uid);return mtClean_(updated);
}
function mtDelete_(table,uid,id,mode){
  const old=mtFind_(table,uid,id);
  if(mode==='permanent'){
    mtSheet_(table).deleteRow(old._row);mtActivity_(uid,'permanently deleted',table,id);mtInvalidate_(uid);return{id};
  }
  const next={...old,updatedAt:mtNow_(),version:Number(old.version||0)+1};
  if(mode==='archive')next.archived=true;
  else next.deletedAt=mtNow_();
  mtWrite_(table,next,old._row);
  mtActivity_(uid,mode==='archive'?'archived':'trashed',table,id);mtInvalidate_(uid);
  return mtClean_(next);
}
function mtSnapshot_(uid){
  const names=['Tasks','Projects','Categories','Tags','Subtasks','RecurringRules','Reminders','FocusSessions','Preferences','Users'];const data={};names.forEach(n=>data[n]=mtList_(n,uid,true).map(mtClean_));
  data.ActivityLogs=mtList_('ActivityLogs',uid,true).slice(-150).reverse().map(mtClean_);
  return data;
}
function mtNextDate_(date,rule){const d=new Date(date+'T12:00:00Z'),step=Math.max(1,Number(rule.interval)||1);
  if(rule.frequency==='weekly'){d.setUTCDate(d.getUTCDate()+7*step);}else if(rule.frequency==='monthly'){const day=d.getUTCDate();d.setUTCDate(1);d.setUTCMonth(d.getUTCMonth()+step);const last=new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth()+1,0)).getUTCDate();d.setUTCDate(Math.min(day,last));}else d.setUTCDate(d.getUTCDate()+step);
  return d.toISOString().slice(0,10);
}
function mtGenerateRecurring_(uid){
  const today=new Date().toISOString().slice(0,10),rules=mtList_('RecurringRules',uid,false);let created=0;
  rules.filter(r=>r.active===true||r.active==='true').forEach(rule=>{
    const template=mtList_('Tasks',uid,false).find(x=>x.id===rule.taskId);if(!template)return;
    let next=rule.nextDate||rule.startDate;let iterations=0;
    while(next&&next<=today&&iterations++<60&&(!rule.endDate||next<=rule.endDate)){
      const signature=rule.id+'@'+next;
      // Deterministic marker in the notes field prevents duplicate generation on retries.
      if(!mtList_('Tasks',uid,true).some(t=>String(t.notes||'').includes('[recurrence:'+signature+']'))){
        const n=mtCreate_('Tasks',uid,{title:template.title,description:template.description||'',status:'To Do',priority:template.priority||'Medium',dueDate:next,projectId:template.projectId||'',categoryId:template.categoryId||'',tags:template.tags||'',notes:'[recurrence:'+signature+']'});created++;
      }
      next=mtNextDate_(next,rule);
    }
    mtWrite_('RecurringRules',{...rule,nextDate:next,updatedAt:mtNow_()},rule._row);
  });return{created};
}
function mtRoute_(user,operation,p){const uid=user.uid;
  switch(operation){
    case 'bootstrap':{
      let profile=mtList_('Users',uid,true)[0];
      if(!profile){return mtWithLock_(()=>{const existing=mtList_('Users',uid,true)[0];if(!existing){profile={id:mtId_(),uid,email:user.email,name:user.name||user.email.split('@')[0],description:'',avatar:'',createdAt:mtNow_(),updatedAt:mtNow_()};mtWrite_('Users',profile)}if(!mtList_('Preferences',uid,true).length)mtWrite_('Preferences',{id:mtId_(),uid,timezone:'Asia/Jakarta',dateFormat:'id-ID',defaultPriority:'Medium',defaultView:'overview',focusMinutes:25,shortBreak:5,longBreak:15,notifyDue:true,createdAt:mtNow_(),updatedAt:mtNow_(),version:1});mtInvalidate_(uid);return mtSnapCached_(uid)})}
      if(!mtList_('Preferences',uid,true).length){return mtWithLock_(()=>{if(!mtList_('Preferences',uid,true).length)mtWrite_('Preferences',{id:mtId_(),uid,timezone:'Asia/Jakarta',dateFormat:'id-ID',defaultPriority:'Medium',defaultView:'overview',focusMinutes:25,shortBreak:5,longBreak:15,notifyDue:true,createdAt:mtNow_(),updatedAt:mtNow_(),version:1});mtInvalidate_(uid);return mtSnapCached_(uid)})}
      return mtSnapCached_(uid);
    }
    case 'snapshot':return mtSnapCached_(uid);
    case 'list':if(!MT.schema[p.table]||['APIRequests','Users'].includes(p.table))mtErr_('INPUT','Invalid table.');return mtList_(p.table,uid,p.includeDeleted===true).map(mtClean_);
    case 'create':return mtWithLock_(()=>{if(!MT.allowed[p.table]||['Users','Preferences'].includes(p.table))mtErr_('INPUT','Invalid entity.');return mtCreate_(p.table,uid,p.data);});
    case 'update':return mtWithLock_(()=>{if(!MT.allowed[p.table]||['Users','Preferences'].includes(p.table))mtErr_('INPUT','Invalid entity.');return mtUpdate_(p.table,uid,p.id,p.data);});
    case 'delete':return mtWithLock_(()=>{if(!['Tasks','Projects','Categories','Tags','Subtasks','Reminders','RecurringRules'].includes(p.table))mtErr_('INPUT','Invalid entity.');return mtDelete_(p.table,uid,p.id,p.mode==='archive'?'archive':'trash');});
    case 'restore':return mtWithLock_(()=>{const x=mtFind_(p.table,uid,p.id);x.deletedAt='';x.archived=false;x.updatedAt=mtNow_();mtWrite_(p.table,x,x._row);mtActivity_(uid,'restored',p.table,p.id);mtInvalidate_(uid);return mtClean_(x);});
    case 'permanentDelete':return mtWithLock_(()=>{if(!['Tasks','Projects','Categories','Tags'].includes(p.table))mtErr_('INPUT','Invalid entity.');return mtDelete_(p.table,uid,p.id,'permanent');});
    case 'profile':return mtList_('Users',uid,true).map(mtClean_)[0]||null;
    case 'profileUpdate':return mtWithLock_(()=>{const x=mtList_('Users',uid,true)[0];if(!x)mtErr_('NOT_FOUND','Profile missing.');const fields=mtValidate_('Users',p.data,true),result={...x,...fields,updatedAt:mtNow_()};mtWrite_('Users',result,x._row);mtInvalidate_(uid);return mtClean_(result);});
    case 'preferencesUpdate':return mtWithLock_(()=>{const x=mtList_('Preferences',uid,true)[0];if(!x)mtErr_('NOT_FOUND','Preferences missing.');const fields=mtValidate_('Preferences',p.data,true),result={...x,...fields,updatedAt:mtNow_(),version:Number(x.version||0)+1};mtWrite_('Preferences',result,x._row);mtInvalidate_(uid);return mtClean_(result);});
    case 'focusSave':return mtWithLock_(()=>{const sec=Number(p.durationSeconds);if(!Number.isFinite(sec)||sec<1||sec>86400)mtErr_('INPUT','Invalid focus duration.');if(p.taskId)mtFind_('Tasks',uid,p.taskId);const row={id:mtId_(),uid,taskId:p.taskId||'',durationSeconds:sec,completed:!!p.completed,startedAt:p.startedAt||mtNow_(),endedAt:mtNow_(),createdAt:mtNow_()};mtWrite_('FocusSessions',row);mtActivity_(uid,'focus session','FocusSessions',row.id);mtInvalidate_(uid);return row;});
    case 'recurrenceRun':return mtWithLock_(()=>{const result=mtGenerateRecurring_(uid);mtInvalidate_(uid);return result;});
    case 'bulkTasks':return mtWithLock_(()=>{if(!Array.isArray(p.ids)||p.ids.length>100)mtErr_('INPUT','Invalid task IDs.');const fields=mtValidate_('Tasks',p.data,true);return p.ids.map(id=>mtUpdate_('Tasks',uid,String(id),fields));});
    default:mtErr_('OPERATION','Unknown operation.');
  }
}
function doGet(){return HtmlService.createHtmlOutput('<h1>MyTask API online</h1><p>Use authenticated form POST.</p>');}
function doPost(e){
  const p=e&&e.parameter||{},requestId=String(p.requestId||'').slice(0,100),replyOrigin=String(p.replyOrigin||''),nonce=String(p.nonce||'').slice(0,100);
  const permitted=(mtProps_().getProperty('ALLOWED_ORIGINS')||'').split(',').map(x=>x.trim()).filter(Boolean);
  let response;const timerStart=Date.now();
  try {
    const serverMode=p.protocol==='MYTASK_SERVER_V2';
    if(!serverMode&&p.protocol!=='MYTASK_BRIDGE_V1')mtErr_('PROTOCOL','Invalid protocol.');
    // Vercel sends requests server-to-server; identity/ownership are always
    // verified independently using the Firebase ID token below.
    if(!serverMode&&!permitted.includes(replyOrigin))mtErr_('ORIGIN','Origin is not allowed.');
    if(String(p.payload||'').length>120000)mtErr_('INPUT','Request too large.');
    const user=mtVerify_(String(p.token||''));
    const payload=JSON.parse(String(p.payload||'{}'));
    response={success:true,data:mtRoute_(user,String(p.operation||''),payload),error:null,meta:{requestId}};
  }catch(err){response={success:false,data:null,error:{code:err.code||'INTERNAL',message:err.code?err.message:'Server error. Please retry.'},meta:{requestId}};}
  console.log('[MyTask Backend Speed]',String(p.operation||'invalid'),'elapsedMs=',Date.now()-timerStart,'success=',!!response.success);
  // JSON is consumed ONLY by the server-to-server Vercel function, eliminating
  // cross-origin browser fetch and fragile iframe relay in production.
  if(p.protocol==='MYTASK_SERVER_V2'){
    return ContentService.createTextOutput(mtJson_(response)).setMimeType(ContentService.MimeType.JSON);
  }
  // Legacy five-file bridge: HtmlService response via iframe postMessage.
  // Never use client-supplied origin without explicit allowlist matching.
  const packet=mtJson_({protocol:'MYTASK_BRIDGE_V1',requestId,nonce,response});
  const target=permitted.includes(replyOrigin)?replyOrigin:'https://invalid.example';
  const html='<!doctype html><html><head><meta charset="utf-8"></head><body><script>window.top.postMessage('+packet+','+mtJson_(target)+');<\/script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
