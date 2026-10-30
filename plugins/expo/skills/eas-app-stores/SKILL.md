/*******************************************************
 * GROUP 6 — Google Apps Script
 * Admin: alkhwadethar@gmail.com
 *******************************************************/

const CFG = {
  APP_NAME: 'GROUP 6',
  ADMIN_EMAIL: 'alkhwadethar@gmail.com',
  DB_PROPERTY: 'GROUP6_DB_ID',
  OTP_TTL: 600,
  SESSION_TTL: 21600,
  SHEETS: {
    Members: ['Email','Name','College','Role','Status','JoinedAt'],
    Attendance: ['ID','Date','Email','Name','Status','Note','RecordedBy','CreatedAt'],
    Tasks: ['ID','Title','Description','AssignedTo','Deadline','Status','CreatedBy','CreatedAt'],
    Surveys: ['ID','Title','Description','QuestionsJSON','Active','CreatedBy','CreatedAt'],
    Responses: ['ID','SurveyID','SurveyTitle','Email','Name','AnswersJSON','SubmittedAt'],
    Evaluations: ['ID','Email','Name','Attendance','Tasks','Quality','Teamwork','Total','Comments','CreatedBy','CreatedAt']
  }
};

function setup() {
  let id = PropertiesService.getScriptProperties().getProperty(CFG.DB_PROPERTY);
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('GROUP 6 Database');
    PropertiesService.getScriptProperties().setProperty(CFG.DB_PROPERTY, ss.getId());
  }
  Object.keys(CFG.SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = CFG.SHEETS[name];
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    else {
      const first = sh.getRange(1,1,1,headers.length).getValues()[0];
      if (first.join('|') !== headers.join('|')) {
        sh.clear();
        sh.appendRow(headers);
      }
    }
  });
  let members = sheet_('Members');
  if (!findRow_('Members','Email',CFG.ADMIN_EMAIL)) {
    members.appendRow([CFG.ADMIN_EMAIL,'إيثار الخواض - Admin','الطب', 'admin','approved',new Date()]);
  } else {
    const row = findRow_('Members','Email',CFG.ADMIN_EMAIL);
    members.getRange(row,4,1,2).setValues([['admin','approved']]);
  }
  Logger.log('GROUP 6 is ready. Database: ' + ss.getUrl());
  return 'GROUP 6 setup complete';
}

function doGet(e) {
  if (e && e.parameter && e.parameter.survey) {
    return HtmlService.createHtmlOutput(publicSurveyHtml_(e.parameter.survey))
      .setTitle('GROUP 6 — Survey')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutput(appHtml_())
    .setTitle(CFG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function db_() {
  const id = PropertiesService.getScriptProperties().getProperty(CFG.DB_PROPERTY);
  if (!id) throw new Error('يرجى تشغيل الدالة setup() أولاً من محرر الأكواد.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) { return db_().getSheetByName(name); }

function rows_(name) {
  const sh = sheet_(name);
  // استخدام getDisplayValues لحل مشاكل تحويل التواريخ بين السيرفر والواجهة
  const v = sh.getDataRange().getDisplayValues(); 
  if (v.length < 2) return [];
  const h = v[0];
  return v.slice(1).filter(r => r.some(x => x !== '')).map(r => {
    const o = {}; h.forEach((k,i) => o[k] = r[i]); return o;
  });
}

function findRow_(name, col, value) {
  const sh = sheet_(name), data = sh.getDataRange().getDisplayValues();
  const ix = data[0].indexOf(col);
  if (ix < 0) return 0;
  for (let i=1;i<data.length;i++)
    if (String(data[i][ix]).toLowerCase() === String(value).toLowerCase()) return i+1;
  return 0;
}

function id_() { return Utilities.getUuid(); }
function email_(s) { return String(s || '').trim().toLowerCase(); }
function clean_(s) { return String(s == null ? '' : s).trim(); }

function requireUser_(token) {
  const email = CacheService.getScriptCache().get('sess_' + token);
  if (!email) throw new Error('انتهت الجلسة. سجّلي الدخول مرة أخرى.');
  const row = findRow_('Members','Email',email);
  if (!row) throw new Error('الحساب غير موجود.');
  const m = rows_('Members').find(x => email_(x.Email) === email);
  if (!m || m.Status !== 'approved') throw new Error('الحساب غير معتمد بعد.');
  return m;
}

function requireAdmin_(token) {
  const m = requireUser_(token);
  if (email_(m.Email) !== CFG.ADMIN_EMAIL || m.Role !== 'admin')
    throw new Error('هذه الصلاحية للمشرفة فقط.');
  return m;
}

/* Authentication */
function requestCode(email, name, college) {
  email = email_(email);
  name = clean_(name);
  college = clean_(college);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))
    throw new Error('أدخلي بريداً إلكترونياً صحيحاً.');
  if (!name || !college) throw new Error('الاسم والكلية مطلوبان.');

  let row = findRow_('Members','Email',email);
  if (!row) {
    sheet_('Members').appendRow([email,name,college,'member','pending',new Date()]);
  } else {
    const m = rows_('Members').find(x => email_(x.Email) === email);
    if (m && m.Status !== 'approved')
      sheet_('Members').getRange(row,2,1,2).setValues([[name,college]]);
  }

  const code = String(Math.floor(100000 + Math.random()*900000));
  CacheService.getScriptCache().put('otp_' + email, code, CFG.OTP_TTL);
  MailApp.sendEmail({
    to: email,
    subject: 'GROUP 6 — Verification code',
    body: 'رمز الدخول لمنصة GROUP 6 هو: ' + code + '\nصالح لمدة 10 دقائق.'
  });
  return 'تم إرسال رمز التحقق إلى بريدك. إذا كان الحساب جديداً فسيحتاج موافقة المشرفة بعد التحقق.';
}

function verifyCode(email, code) {
  email = email_(email);
  const cached = CacheService.getScriptCache().get('otp_' + email);
  if (!cached || cached !== clean_(code))
    throw new Error('رمز التحقق غير صحيح أو انتهت صلاحيته.');
  CacheService.getScriptCache().remove('otp_' + email);

  const member = rows_('Members').find(x => email_(x.Email) === email);
  if (!member) throw new Error('الحساب غير موجود.');
  if (member.Status !== 'approved')
    return {pending:true, message:'تم التحقق من البريد. الحساب بانتظار موافقة المشرفة.'};

  const token = id_();
  CacheService.getScriptCache().put('sess_' + token, email, CFG.SESSION_TTL);
  return {
    pending:false,
    token:token,
    user:{email:member.Email,name:member.Name,role:member.Role}
  };
}

function logout(token) {
  CacheService.getScriptCache().remove('sess_' + token);
  return true;
}

/* Main data */
function getData(token) {
  const me = requireUser_(token);
  const admin = me.Role === 'admin';
  const members = rows_('Members');
  const tasks = rows_('Tasks').filter(t =>
    admin || !t.AssignedTo || email_(t.AssignedTo) === email_(me.Email));
  const attendance = rows_('Attendance').filter(a =>
    admin || email_(a.Email) === email_(me.Email));
  const evaluations = rows_('Evaluations').filter(x =>
    admin || email_(x.Email) === email_(me.Email));
  const surveys = rows_('Surveys').filter(s =>
    String(s.Active).toLowerCase() === 'true' || s.Active === true);
    
  let appUrl = '';
  try { appUrl = ScriptApp.getService().getUrl(); } catch(e) {}

  return {
    appUrl: appUrl,
    me: {email:me.Email,name:me.Name,role:me.Role},
    members: admin ? members : members.filter(x => email_(x.Email) === email_(me.Email)),
    pending: admin ? members.filter(x => x.Status !== 'approved') : [],
    tasks, attendance, evaluations,
    surveys: surveys.map(s => ({
      ID:s.ID,Title:s.Title,Description:s.Description,
      Active:s.Active,CreatedAt:s.CreatedAt
    })),
    stats: {
      members: members.filter(x=>x.Status==='approved').length,
      tasks: rows_('Tasks').length,
      surveys: surveys.length,
      attendance: attendance.length
    }
  };
}

/* Members */
function approveMember(token,email) {
  requireAdmin_(token);
  const row = findRow_('Members','Email',email_(email));
  if (!row) throw new Error('العضو غير موجود.');
  sheet_('Members').getRange(row,5).setValue('approved');
  return true;
}

function rejectMember(token,email) {
  requireAdmin_(token);
  const row = findRow_('Members','Email',email_(email));
  if (row) sheet_('Members').deleteRow(row);
  return true;
}

/* Attendance */
function recordAttendance(token,email,status,note) {
  const me = requireUser_(token);
  email = email_(email || me.Email);
  if (me.Role !== 'admin' && email !== email_(me.Email))
    throw new Error('يمكنك تسجيل حضورك فقط.');
  const target = rows_('Members').find(x =>
    email_(x.Email)===email && x.Status==='approved');
  if (!target) throw new Error('العضو غير معتمد.');
  sheet_('Attendance').appendRow([
    id_(),new Date(),email,target.Name,clean_(status),
    clean_(note),me.Email,new Date()
  ]);
  return true;
}

/* Tasks */
function createTask(token,title,description,assignedTo,deadline) {
  const me = requireAdmin_(token);
  if (!clean_(title)) throw new Error('عنوان المهمة مطلوب.');
  sheet_('Tasks').appendRow([
    id_(),clean_(title),clean_(description),email_(assignedTo),
    clean_(deadline),'Pending',me.Email,new Date()
  ]);
  return true;
}

function updateTaskStatus(token,taskId,status) {
  const me = requireUser_(token);
  const sh=sheet_('Tasks'), data=sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++) {
    if(String(data[i][0])===String(taskId)) {
      if(me.Role!=='admin' && email_(data[i][3])!==email_(me.Email))
        throw new Error('لا يمكنك تعديل هذه المهمة.');
      sh.getRange(i+1,6).setValue(clean_(status));
      return true;
    }
  }
  throw new Error('المهمة غير موجودة.');
}

/* Surveys */
function createSurvey(token,title,description,questions) {
  const me = requireAdmin_(token);
  if(!clean_(title)) throw new Error('عنوان الاستبيان مطلوب.');
  if(!Array.isArray(questions) || !questions.length)
    throw new Error('أضيفي سؤالاً واحداً على الأقل.');
  const sid = id_();
  sheet_('Surveys').appendRow([
    sid,clean_(title),clean_(description),
    JSON.stringify(questions),true,me.Email,new Date()
  ]);
  return (ScriptApp.getService().getUrl() || '') + '?survey=' + encodeURIComponent(sid);
}

function setSurveyActive(token,sid,active) {
  requireAdmin_(token);
  const row = findRow_('Surveys','ID',sid);
  if(!row) throw new Error('الاستبيان غير موجود.');
  sheet_('Surveys').getRange(row,5).setValue(!!active);
  return true;
}

function getPublicSurvey(sid) {
  const s = rows_('Surveys').find(x=>String(x.ID)===String(sid));
  if(!s || !(s.Active===true || String(s.Active).toLowerCase()==='true'))
    throw new Error('الاستبيان غير متاح.');
  return {
    id:s.ID,
    title:s.Title,
    description:s.Description,
    questions:JSON.parse(s.QuestionsJSON||'[]')
  };
}

function submitPublicResponse(sid,name,email,answers) {
  const s = getPublicSurvey(sid);
  email = email_(email);
  if(!clean_(name) || !email)
    throw new Error('الاسم والبريد مطلوبان.');
  sheet_('Responses').appendRow([
    id_(),sid,s.title,email,clean_(name),
    JSON.stringify(answers||{}),new Date()
  ]);
  return 'تم إرسال إجابتك بنجاح. شكراً لك.';
}

/* Evaluations */
function createEvaluation(token,email,attendance,tasks,quality,teamwork,comments) {
  const me = requireAdmin_(token);
  email = email_(email);
  const m = rows_('Members').find(x=>email_(x.Email)===email);
  if(!m) throw new Error('العضو غير موجود.');
  const vals = [
    Number(attendance)||0,
    Number(tasks)||0,
    Number(quality)||0,
    Number(teamwork)||0
  ];
  const total = vals.reduce((a,b)=>a+b,0)/4;
  sheet_('Evaluations').appendRow([
    id_(),email,m.Name,vals[0],vals[1],vals[2],vals[3],
    total,clean_(comments),me.Email,new Date()
  ]);
  return true;
}

function getAnalytics(token) {
  requireAdmin_(token);
  const members=rows_('Members');
  const attendance=rows_('Attendance');
  const tasks=rows_('Tasks');
  const responses=rows_('Responses');
  const evals=rows_('Evaluations');

  const byStatus={};
  attendance.forEach(x=>byStatus[x.Status]=(byStatus[x.Status]||0)+1);

  const taskStatus={};
  tasks.forEach(x=>taskStatus[x.Status]=(taskStatus[x.Status]||0)+1);

  return {
    members: members.length,
    approved: members.filter(x=>x.Status==='approved').length,
    attendanceByStatus: byStatus,
    tasksByStatus: taskStatus,
    responses: responses.length,
    evaluations: evals.length,
    averageEvaluation: evals.length
      ? evals.reduce((a,x)=>a+(Number(x.Total)||0),0)/evals.length
      : 0
  };
}

/* Public survey page */
function publicSurveyHtml_(sid) {
  const safeSid = JSON.stringify(String(sid));
  return `<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GROUP 6 Survey</title>
<style>
body{font-family:Arial,sans-serif;background:#f4f7fb;margin:0;padding:20px;color:#243047;direction:rtl}
.card{max-width:650px;margin:auto;background:white;padding:24px;border-radius:18px;box-shadow:0 4px 22px #0001}
input,textarea,select{width:100%;box-sizing:border-box;padding:12px;margin:7px 0 15px;border:1px solid #ccd5e2;border-radius:9px;font-family:inherit}
button{background:#3867e8;color:white;border:0;padding:13px 22px;border-radius:9px;font-size:16px;cursor:pointer}
button:hover{background:#2a52be}
</style>
</head>
<body>
<div class="card">
<h2 id="title">جاري التحميل...</h2>
<p id="desc"></p>
<div id="form"></div>
</div>
<script>
const sid=${safeSid};
let survey;

google.script.run.withSuccessHandler(s=>{
  survey=s;
  document.getElementById('title').textContent=s.title;
  document.getElementById('desc').textContent=s.description||'';

  let h='<label>الاسم الكامل</label><input id="name" required><label>البريد الإلكتروني</label><input id="email" type="email" required>';

  s.questions.forEach((q,i)=>{
    h+='<label>'+esc(q.text||q)+'</label>';
    if(q.type==='textarea')
      h+='<textarea id="q'+i+'"></textarea>';
    else if(q.type==='select')
      h+='<select id="q'+i+'">'+(q.options||[]).map(o=>'<option>'+esc(o)+'</option>').join('')+'</select>';
    else
      h+='<input id="q'+i+'" '+(q.type==='number'?'type="number"':'type="text"')+'>';
  });

  h+='<button onclick="send()">إرسال الإجابات</button><p id="msg"></p>';
  document.getElementById('form').innerHTML=h;

}).withFailureHandler(e=>{
  document.getElementById('form').textContent='الاستبيان غير متاح أو الرابط غير صحيح.';
}).getPublicSurvey(sid);

function esc(x){
  return String(x).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
}

function send(){
  let a={};
  survey.questions.forEach((q,i)=>{
    a[q.text||q]=document.getElementById('q'+i).value;
  });

  google.script.run
  .withSuccessHandler(m=>{
    document.getElementById('msg').innerHTML='<strong style="color:green">'+m+'</strong>';
    document.getElementById('form').innerHTML='';
  })
  .withFailureHandler(e=>document.getElementById('msg').innerHTML='<span style="color:red">'+e.message+'</span>')
  .submitPublicResponse(
    sid,
    document.getElementById('name').value,
    document.getElementById('email').value,
    a
  );
}
</script>
</body>
</html>`;
}

/* Main application interface */
function appHtml_() {
return `<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<base target="_top">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GROUP 6 Platform</title>
<style>
:root{--blue:#4169e1;--bg:#f4f7fc;--ink:#243047;--muted:#718096}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);font-family:Arial,sans-serif;color:var(--ink);direction:rtl}
header{background:linear-gradient(120deg,#4169e1,#7c55d9);color:white;padding:22px}
header h1{margin:0;font-size:25px}
header p{margin:6px 0 0;opacity:.9}
main{max-width:1050px;margin:20px auto;padding:0 14px}
.card{background:white;border-radius:16px;padding:18px;margin-bottom:15px;box-shadow:0 3px 15px #263b5a0b}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}
.stat{padding:16px;border-radius:13px;background:#edf2ff;text-align:center}
.stat b{font-size:25px;display:block;color:var(--blue)}
input,textarea,select{width:100%;padding:11px;border:1px solid #d8deea;border-radius:9px;margin:6px 0 12px;font:inherit}
button{border:0;border-radius:9px;background:var(--blue);color:white;padding:10px 15px;margin:4px;cursor:pointer;font:inherit;transition:0.2s}
button:hover{opacity:0.9}
button.alt{background:#e9edfa;color:#34466e}
button.danger{background:#d94c64;color:white}
nav{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:14px}
.hidden{display:none}
.muted{color:var(--muted);font-size:13px}
.item{padding:12px;border:1px solid #e5eaf2;border-radius:10px;margin:8px 0;background:#fafbfc}
table{border-collapse:collapse;width:100%;font-size:14px;margin-top:10px}
td,th{padding:9px;border:1px solid #e7ebf2;text-align:right}
th{background:#f4f7fc}
#auth{max-width:440px;margin:30px auto}
.tag{display:inline-block;padding:4px 8px;background:#eef2ff;border-radius:20px;font-size:12px;color:#3867e8;font-weight:bold}
.admin-badge{background:#fff3cd;color:#856404;padding:3px 8px;border-radius:12px;font-size:11px}
</style>
</head>

<body>
<header>
<h1>GROUP 6</h1>
<p>منصة إدارة المهام الأكاديمية والتعاون</p>
</header>

<main>
<section id="auth" class="card">
<h2>تسجيل الدخول</h2>

<div id="authStep1">
<label>البريد الإلكتروني (الجامعي أو الشخصي)</label>
<input id="email" type="email" placeholder="name@example.com">

<label>الاسم الكامل</label>
<input id="name" placeholder="الاسم الكامل كما في السجلات">

<label>الكلية / الجامعة</label>
<input id="college" placeholder="اسم الكلية (مثال: كلية الطب)">

<button onclick="sendCode()" style="width:100%;margin-top:10px">إرسال رمز التحقق</button>
</div>

<div id="authStep2" class="hidden">
<label>رمز التحقق (تم إرساله للبريد)</label>
<input id="otp" inputmode="numeric" placeholder="أدخل الرمز المكون من 6 أرقام">
<button onclick="verify()" style="width:100%;margin-top:10px">تأكيد الدخول</button>
</div>

<p id="authMsg" class="muted" style="margin-top:15px;text-align:center"></p>
</section>

<section id="app" class="hidden">
<div class="card" style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap">
<b id="welcome" style="font-size:18px"></b>
<div>
<button class="alt" onclick="refresh()">تحديث البيانات ⟳</button>
<button class="danger" onclick="signout()">تسجيل الخروج</button>
</div>
</div>
<nav id="nav"></nav>
<div id="view"></div>
</section>
</main>

<script>
let token='', data=null, tab='home';

function el(id){ return document.getElementById(id); }

function call(fn,args,ok){
  google.script.run
  .withSuccessHandler(ok||(()=>{}))
  .withFailureHandler(e=>alert(e.message))[fn](...args);
}

function sendCode(){
  const e=el('email').value, n=el('name').value, c=el('college').value;
  el('authMsg').textContent='جاري الإرسال...';
  call('requestCode',[e,n,c],m=>{
    el('authMsg').textContent=m;
    el('authStep1').classList.add('hidden');
    el('authStep2').classList.remove('hidden');
  });
}

function verify(){
  el('authMsg').textContent='جاري التحقق...';
  call('verifyCode',[el('email').value,el('otp').value],r=>{
    if(r.pending){
      el('authMsg').textContent=r.message;
      return;
    }
    token=r.token;
    el('auth').classList.add('hidden');
    el('app').classList.remove('hidden');
    refresh();
  });
}

function signout(){
  call('logout',[token],()=>location.reload());
}

function refresh(){
  el('view').innerHTML='<div style="text-align:center;padding:40px">جاري جلب البيانات...</div>';
  call('getData',[token],d=>{
    data=d;
    el('welcome').innerHTML='مرحباً، '+d.me.name+
      (d.me.role==='admin'?' <span class="admin-badge">مشرفة النظام</span>':'');
    renderNav();
    show(tab);
  });
}

function renderNav(){
  let tabs=[
    ['home','الرئيسية'],
    ['attendance','الحضور'],
    ['tasks','المهام'],
    ['surveys','الاستبيانات'],
    ['members','الأعضاء'],
    ['evaluations','تقييمي']
  ];

  if(data.me.role==='admin') tabs.push(['analytics','التحليلات والإحصاءات']);

  el('nav').innerHTML=tabs.map(t=>
    '<button class="alt" onclick="show(\\''+t[0]+'\\')" style="'+(tab===t[0]?'background:#4169e1;color:white':'')+'">'+t[1]+'</button>'
  ).join('');
}

function show(t){
  tab=t;
  renderNav();
  let h='';

  if(t==='home'){
    h='<div class="grid">'+
    '<div class="stat"><b>'+data.stats.members+'</b>الأعضاء المعتمدون</div>'+
    '<div class="stat"><b>'+data.stats.tasks+'</b>إجمالي المهام</div>'+
    '<div class="stat"><b>'+data.stats.surveys+'</b>الاستبيانات النشطة</div>'+
    '<div class="stat"><b>'+data.stats.attendance+'</b>سجلات حضورك</div>'+
    '</div><div class="card"><h3>منصة GROUP 6</h3><p>مرحباً بك في لوحة التحكم الخاصة بالمجموعة. يمكنك من هنا متابعة مهامك، تسجيل الحضور، والوصول للاستبيانات.</p></div>';

    if(data.pending.length) {
      h+='<div class="card"><h3>طلبات الانضمام المعلقة</h3>'+
      data.pending.map(m=>
        '<div class="item"><b>'+m.Name+'</b> <small>('+m.College+')</small><br><span class="muted">'+m.Email+'</span><br>'+
        '<button onclick="approve(\\''+m.Email+'\\')">قبول</button>'+
        '<button class="danger" onclick="reject(\\''+m.Email+'\\')">رفض</button>'+
        '</div>'
      ).join('')+'</div>';
    }
  }

  else if(t==='members'){
    h='<div class="card"><h3>قائمة الأعضاء</h3>'+
    data.members.map(m=>
      '<div class="item"><b>'+m.Name+'</b> <small class="muted">'+m.College+'</small><br>'+
      '<span class="muted">'+m.Email+'</span> '+
      '<span class="tag">'+(m.Role==='admin'?'مشرف':m.Status)+'</span></div>'
    ).join('')+'</div>';
  }

  else if(t==='attendance'){
    h='<div class="card"><h3>تسجيل الحضور</h3>';
    if(data.me.role==='admin')
      h+='<label>البريد الإلكتروني للعضو</label><input id="attEmail" placeholder="email@example.com (اتركيه فارغاً لتسجيل حضورك)">';

    h+='<label>حالة الحضور</label><select id="attStatus">'+
      '<option value="Present">حاضر (Present)</option><option value="Absent">غائب (Absent)</option>'+
      '<option value="Late">متأخر (Late)</option><option value="Excused">مستأذن (Excused)</option></select>'+
      '<label>ملاحظات (اختياري)</label><input id="attNote" placeholder="اكتب ملاحظة إن وجدت...">'+
      '<button onclick="attendance()">تسجيل السجل</button></div>'+
      '<div class="card"><h3>سجلات الحضور الخاصة بك</h3>'+
      (data.attendance.length ? data.attendance.reverse().map(a=>
        '<div class="item"><b>'+a.Date.split(' ')[0]+'</b> — '+a.Name+' <span class="tag">'+a.Status+'</span><br>'+
        '<small class="muted">'+(a.Note||'لا توجد ملاحظات')+'</small></div>'
      ).join('') : '<p class="muted">لا توجد سجلات بعد.</p>')+'</div>';
  }

  else if(t==='tasks'){
    if(data.me.role==='admin'){
      h+='<div class="card"><h3>إضافة مهمة جديدة</h3>'+
      '<label>عنوان المهمة</label><input id="taskTitle">'+
      '<label>التفاصيل</label><textarea id="taskDesc" rows="3"></textarea>'+
      '<label>البريد المكلّف (اتركيه فارغاً لتكون مهمة عامة)</label><input id="taskEmail">'+
      '<label>تاريخ التسليم النهائي</label><input id="taskDeadline" type="date">'+
      '<button onclick="addTask()">إصدار المهمة</button></div>';
    }
    h+='<div class="card"><h3>قائمة المهام</h3>'+
      (data.tasks.length ? data.tasks.map(x=>
        '<div class="item"><b>'+x.Title+'</b><p style="margin:5px 0">'+x.Description+'</p>'+
        '<small class="muted">التسليم: '+x.Deadline+' | الحالة: '+x.Status+'</small><br>'+
        '<div style="margin-top:10px">'+
        '<select id="status_'+x.ID+'" style="width:auto;display:inline-block;padding:7px;margin:0 0 0 10px">'+
        '<option value="Pending" '+(x.Status==='Pending'?'selected':'')+'>قيد الانتظار (Pending)</option>'+
        '<option value="In Progress" '+(x.Status==='In Progress'?'selected':'')+'>قيد التنفيذ (In Progress)</option>'+
        '<option value="Completed" '+(x.Status==='Completed'?'selected':'')+'>مكتملة (Completed)</option></select>'+
        '<button class="alt" onclick="taskStatus(\\''+x.ID+'\\')">تحديث الحالة</button>'+
        '</div></div>'
      ).join('') : '<p class="muted">لا توجد مهام مسندة.</p>')+'</div>';
  }

  else if(t==='surveys'){
    if(data.me.role==='admin'){
      h+='<div class="card"><h3>إنشاء استبيان جديد</h3>'+
      '<label>عنوان الاستبيان</label><input id="surveyTitle">'+
      '<label>الوصف</label><textarea id="surveyDesc"></textarea>'+
      '<label>الأسئلة (اكتبي كل سؤال في سطر منفصل)</label>'+
      '<textarea id="surveyQs" rows="4" placeholder="السؤال الأول؟&#10;السؤال الثاني؟"></textarea>'+
      '<button onclick="addSurvey()">حفظ وإنشاء رابط</button>'+
      '<p id="surveyResult" style="margin-top:10px"></p></div>';
    }
    h+='<div class="card"><h3>الاستبيانات المتاحة</h3>'+
      (data.surveys.length ? data.surveys.map(s=>
        '<div class="item"><b>'+s.Title+'</b><p style="margin:5px 0">'+(s.Description||'')+'</p>'+
        '<button class="alt" onclick="copyLink(\\''+s.ID+'\\')">نسخ رابط المشاركة 🔗</button></div>'
      ).join('') : '<p class="muted">لا توجد استبيانات نشطة.</p>')+'</div>';
  }

  else if(t==='evaluations'){
    h='<div class="card"><h3>سجل تقييماتك</h3>'+
      (data.evaluations.length ? data.evaluations.map(x=>
        '<div class="item"><div style="display:flex;gap:15px;margin-bottom:10px">'+
        '<span>الالتزام بالحضور: <b>'+x.Attendance+'</b></span>'+
        '<span>إنجاز المهام: <b>'+x.Tasks+'</b></span>'+
        '<span>جودة العمل: <b>'+x.Quality+'</b></span>'+
        '<span>العمل الجماعي: <b>'+x.Teamwork+'</b></span></div>'+
        '<b>التقييم الإجمالي: '+Number(x.Total).toFixed(1)+' / 100</b>'+
        '<p class="muted" style="margin:5px 0">الملاحظات: '+(x.Comments||'لا يوجد')+'</p></div>'
      ).join('') : '<p class="muted">لا توجد تقييمات مسجلة لك حتى الآن.</p>')+'</div>';
  }

  else if(t==='analytics' && data.me.role==='admin'){
    h='<div class="card">جاري جلب التحليلات...</div>';
    call('getAnalytics',[token],a=>{
      
      let attTbl = '<table><tr><th>حالة الحضور</th><th>مرات التكرار</th></tr>';
      for(let k in a.attendanceByStatus) attTbl+= '<tr><td>'+k+'</td><td>'+a.attendanceByStatus[k]+'</td></tr>';
      attTbl+='</table>';

      let tskTbl = '<table><tr><th>حالة المهام</th><th>العدد</th></tr>';
      for(let k in a.tasksByStatus) tskTbl+= '<tr><td>'+k+'</td><td>'+a.tasksByStatus[k]+'</td></tr>';
      tskTbl+='</table>';

      el('view').innerHTML=
        '<div class="grid" style="margin-bottom:15px">'+
        '<div class="stat"><b>'+a.approved+'</b>عضو معتمد</div>'+
        '<div class="stat"><b>'+a.responses+'</b>رد على الاستبيانات</div>'+
        '<div class="stat"><b>'+Number(a.averageEvaluation).toFixed(1)+'%</b>متوسط التقييم العام</div>'+
        '</div>'+
        '<div class="card"><h3>إحصاءات الحضور</h3>'+attTbl+'</div>'+
        '<div class="card"><h3>مؤشرات المهام</h3>'+tskTbl+'</div>';
    });
  }

  if(h) el('view').innerHTML=h;
}

function approve(e){ call('approveMember',[token,e],refresh); }
function reject(e){ if(confirm('تأكيد الرفض؟')) call('rejectMember',[token,e],refresh); }

function attendance(){
  let em = el('attEmail') ? el('attEmail').value : '';
  call('recordAttendance',[token, em, el('attStatus').value, el('attNote').value], refresh);
}

function addTask(){
  call('createTask',[token, el('taskTitle').value, el('taskDesc').value, el('taskEmail').value, el('taskDeadline').value], refresh);
}

function taskStatus(id){ call('updateTaskStatus',[token, id, el('status_'+id).value], refresh); }

function addSurvey(){
  let qs=el('surveyQs').value.split('\\n').map(x=>x.trim()).filter(Boolean).map(x=>({text:x,type:'text'}));
  el('surveyResult').innerHTML = 'جاري الإنشاء...';
  call('createSurvey',[token, el('surveyTitle').value, el('surveyDesc').value, qs], url=>{
    el('surveyResult').innerHTML = '<span style="color:green">تم الإنشاء!</span> <a target="_blank" href="'+url+'">تجرية الرابط</a>';
    refresh();
  });
}

function copyLink(id){
  const base = data.appUrl || location.href.split('?')[0];
  const u = base + '?survey=' + encodeURIComponent(id);
  if(navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(u).then(()=>alert('تم نسخ الرابط بنجاح!')).catch(()=>prompt('انسخي الرابط من هنا:',u));
  } else {
    prompt('انسخي الرابط من هنا:',u);
  }
}
</script>
</body>
</html>`;
}
