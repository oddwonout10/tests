/**
 * Aptitude Test Backend – Google Apps Script
 * One spreadsheet. Web App endpoints: action=start, submit, config.
 * Images are embedded at import as base64 from Drive file IDs.
 */

const TZ = 'Asia/Kolkata';
const RESUME_MINUTES_DEFAULT = 15; // overridden per test via code_ttl_min

// ===== Utilities =====
const S = () => SpreadsheetApp.getActiveSpreadsheet();
const SH = name => S().getSheetByName(name) || S().insertSheet(name);

function fmtIST(date, pattern) {
  return Utilities.formatDate(date, TZ, pattern);
}
function now() { return new Date(); }

function getConfigRow(testId) {
  const sh = SH('Config_Tests');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idx = {}
  headers.forEach((h,i)=> idx[h]=i);
  for (let r=1; r<data.length; r++) {
    if ((data[r][idx['test_id']] || '').toString().trim() === testId) {
      // Return as object
      return {
        row: r+1,
        test_id: data[r][idx['test_id']],
        title: data[r][idx['title']],
        doc_url: data[r][idx['doc_url']],
        time_limit_min: Number(data[r][idx['time_limit_min']])||0,
        status: data[r][idx['status']],
        admin_email: data[r][idx['admin_email']],
        allow_resume: (data[r][idx['allow_resume']]||'Y').toString().toUpperCase()==='Y',
        code_ttl_min: Number(data[r][idx['code_ttl_min']])||RESUME_MINUTES_DEFAULT,
        rate_limit_per_10min: Number(data[r][idx['rate_limit_per_10min']])||5,
      };
    }
  }
  throw new Error('Unknown test_id '+testId);
}

function codesSheet(testId){ return SH('Codes_'+testId); }
function responsesSheet(testId){ return SH('Responses_'+testId); }

function getTestJson(testId) {
  const text = PropertiesService.getDocumentProperties().getProperty('TEST_JSON_'+testId);
  if (!text) throw new Error('No published JSON for test '+testId+'. Run "Import From Google Doc" then "Publish Test JSON".');
  return JSON.parse(text);
}

function setTestJson(testId, obj) {
  PropertiesService.getDocumentProperties().setProperty('TEST_JSON_'+testId, JSON.stringify(obj));
}

// ===== Admin menu =====
function onOpen(){
  SpreadsheetApp.getUi().createMenu('Aptitude Admin')
    .addItem('Import From Google Doc…','menu_import')
    .addItem('Publish Test JSON','menu_publish')
    .addItem('Generate Codes…','menu_generateCodes')
    .addItem('Expire Locks','menu_expireLocks')
    .addItem('Rebuild Student Index','menu_rebuildStudentIndex')
    .addToUi();
}

function menu_import(){
  const testId = SpreadsheetApp.getUi().prompt('Enter test_id to import').getResponseText().trim();
  importFromGoogleDoc(testId);
  SpreadsheetApp.getUi().alert('Imported from Doc for '+testId);
}
function menu_publish(){
  const testId = SpreadsheetApp.getUi().prompt('Enter test_id to publish').getResponseText().trim();
  publishTestJson(testId);
  SpreadsheetApp.getUi().alert('Published JSON for '+testId);
}
function menu_generateCodes(){
  const testId = SpreadsheetApp.getUi().prompt('Enter test_id').getResponseText().trim();
  const count = Number(SpreadsheetApp.getUi().prompt('How many codes?').getResponseText().trim());
  generateCodes(testId, count);
  SpreadsheetApp.getUi().alert('Generated '+count+' codes for '+testId);
}
function menu_expireLocks(){
  expireStaleInUse();
  SpreadsheetApp.getUi().alert('Expired stale IN_USE codes');
}
function menu_rebuildStudentIndex(){
  rebuildStudentIndex();
  SpreadsheetApp.getUi().alert('Student_Index rebuilt');
}

// ===== Code generation =====
function generateCodes(testId, count){
  const sh = codesSheet(testId);
  if (sh.getLastRow()===0) sh.appendRow(['code','state','issued_at_IST','started_at_IST','submitted_at_IST','used_by_name','used_by_email','resume_until_IST']);
  const existing = new Set(sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),1).getValues().flat().filter(Boolean).map(String));
  const rows=[];
  while(rows.length<count){
    const code = ('000000'+Math.floor(Math.random()*1e6)).slice(-6);
    if (!existing.has(code)){
      rows.push([code,'AVAILABLE',fmtIST(now(),'dd/MM/YYYY HH:mm:ss'),'', '', '', '', '']);
      existing.add(code);
    }
  }
  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
}

// ===== Import & publish =====
function importFromGoogleDoc(testId){
  const cfg = getConfigRow(testId);
  const docId = (cfg.doc_url.match(/\/d\/([A-Za-z0-9_-]+)/)||[])[1] || cfg.doc_url;
  const body = DocumentApp.openById(docId).getBody().getText();
  const lines = body.split(/\r?\n/).map(s=>s.trim()).filter(s=>s.length>0);
  const titleLine = lines[0].startsWith('#') ? lines.shift().replace(/^#\s*/,'') : cfg.title;

  const questions=[];
  let cursor=0, qno=0;
  while(cursor < lines.length){
    if(!/^Q\d+\./.test(lines[cursor])){ cursor++; continue; }
    const qHeader = lines[cursor++];
    qno++;
    const qText = qHeader.replace(/^Q\d+\.\s*/,'').trim();
    let imgData = '';
    if (cursor<lines.length && /^IMG:/.test(lines[cursor])){
      const fileId = lines[cursor].replace(/^IMG:\s*/,'').trim();
      try { imgData = driveFileToDataUrl(fileId); } catch(err){ imgData=''; }
      cursor++;
    }
    const opts=[];
    for (let k=0; k<5 && cursor<lines.length; k++){
      const m = lines[cursor].match(/^[A-E]\)\s*(.*)$/);
      if(!m) break;
      opts.push(m[1]);
      cursor++;
    }
    if (opts.length!==5) throw new Error('Question '+qno+' does not have 5 options.');
    questions.push({ qno, text:qText, image:imgData, options:opts });
  }

  PropertiesService.getDocumentProperties().setProperty('IMPORTED_'+testId, JSON.stringify({ title:titleLine, questions }));
}

function publishTestJson(testId){
  const raw = PropertiesService.getDocumentProperties().getProperty('IMPORTED_'+testId);
  if (!raw) throw new Error('Nothing imported for '+testId);
  setTestJson(testId, JSON.parse(raw));
}

function driveFileToDataUrl(fileId){
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const mime = blob.getContentType();
  const b64 = Utilities.base64Encode(blob.getBytes());
  return 'data:'+mime+';base64,'+b64;
}

// ===== Rate limiting (best-effort using client IP from frontend) =====
function checkRateLimit(ip, maxPer10){
  if (!ip) return; // skip if missing
  const cache = CacheService.getScriptCache();
  const bucket = Math.floor(now().getTime()/ (10*60*1000));
  const key = 'rl:'+ip+':'+bucket;
  const curr = Number(cache.get(key)||'0');
  if (curr >= maxPer10) throw new Error('Too many attempts. Try again later.');
  cache.put(key, String(curr+1), 600);
}

// ===== Web app endpoints =====
function doPost(e){
  const params = e.parameter || {};
  const action = params.action || '';
  try {
    if (action==='start') return json(startHandler(params));
    if (action==='submit') return json(submitHandler(params));
    if (action==='config') return json({ ok:true, tz:TZ });
    throw new Error('Unknown action');
  } catch(err){
    return json({ ok:false, error: String(err) });
  }
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function startHandler(p){
  const name = (p.name||'').trim();
  const email = (p.email||'').trim();
  const testId = (p.test_id||'').trim();
  const code = (p.code||'').trim();
  const ip = (p.ip||'').trim();
  const ua = (p.ua||'').trim();

  if (!name || !email || !testId || !code) throw new Error('Missing fields');

  const cfg = getConfigRow(testId);
  if (cfg.status!=='ACTIVE') throw new Error('Test not active');

  checkRateLimit(ip, cfg.rate_limit_per_10min);

  const sh = codesSheet(testId);
  const data = sh.getDataRange().getValues();
  if (data.length<2) throw new Error('No codes exist');
  const headers = data[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);

  let row = -1; let rec = null;
  for (let r=1; r<data.length; r++){
    if (String(data[r][idx['code']])===code){ row=r+1; rec=data[r]; break; }
  }
  if (row<0) throw new Error('Invalid code');

  const state = String(rec[idx['state']]||'');
  const nowD = now();
  const startStr = fmtIST(nowD,'dd/MM/YYYY HH:mm:ss');

  let serverStart = null; let endAtMs = null; let remainingSec = null;

  if (state==='AVAILABLE'){
    const resumeUntil = new Date(nowD.getTime()+ cfg.code_ttl_min*60*1000);
    sh.getRange(row, idx['state']+1, 1, 7).setValues([[
      'IN_USE', fmtIST(nowD,'dd/MM/YYYY HH:mm:ss'), startStr, '', name, email, fmtIST(resumeUntil,'dd/MM/YYYY HH:mm:ss')
    ]]);
    serverStart = nowD;
  } else if (state==='IN_USE'){
    // read original start
    const startedAtStr = rec[idx['started_at_IST']]||startStr;
    const startedAt = parseIST(startedAtStr);
    const resumeUntilStr = rec[idx['resume_until_IST']]||'';
    const resumeUntil = resumeUntilStr ? parseIST(resumeUntilStr) : new Date(startedAt.getTime()+cfg.code_ttl_min*60*1000);
    if (nowD > resumeUntil && !cfg.allow_resume) throw new Error('Resume window over');
    serverStart = startedAt;
  } else if (state==='USED'){
    throw new Error('Code already used');
  } else if (state==='EXPIRED'){
    throw new Error('Code expired');
  } else {
    throw new Error('Code state invalid');
  }

  endAtMs = serverStart.getTime() + cfg.time_limit_min*60*1000;
  remainingSec = Math.max(0, Math.floor((endAtMs - nowD.getTime())/1000));

  const test = getTestJson(testId);
  return { ok:true, 
    test_id:testId, 
    title:test.title, 
    time_limit_min: cfg.time_limit_min, 
    remainingSec, 
    serverNow: fmtIST(nowD,'dd/MM/YYYY HH:mm:ss'), 
    endAt: fmtIST(new Date(endAtMs),'dd/MM/YYYY HH:mm:ss'), 
    questions:test.questions };
}

function submitHandler(p){
  const name = (p.name||'').trim();
  const email = (p.email||'').trim();
  const testId = (p.test_id||'').trim();
  const code = (p.code||'').trim();
  const answers = (p.answers_json||'[]');
  const ip = (p.ip||'').trim();
  const ua = (p.ua||'').trim();
  const clientEndedAt = (p.client_ended_at||'');

  const cfg = getConfigRow(testId);
  const shC = codesSheet(testId);
  const data = shC.getDataRange().getValues();
  const headers = data[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i);

  let row = -1; let rec=null;
  for (let r=1; r<data.length; r++){
    if (String(data[r][idx['code']])===code){ row=r+1; rec=data[r]; break; }
  }
  if (row<0) throw new Error('Invalid code');

  const state = String(rec[idx['state']]||'');
  if (state==='USED') throw new Error('Already submitted');

  const startedAt = parseIST(rec[idx['started_at_IST']]||fmtIST(now(),'dd/MM/YYYY HH:mm:ss'));
  const endDeadline = new Date(startedAt.getTime() + cfg.time_limit_min*60*1000);
  const endNow = now();
  const late = endNow > endDeadline ? 'Y' : 'N';
  const durationSec = Math.max(0, Math.floor((endNow.getTime()-startedAt.getTime())/1000));

  // Write response row
  const shR = responsesSheet(testId);
  if (shR.getLastRow()===0) shR.appendRow(['timestamp_IST','student_name','student_email','code','client_ip','user_agent','started_IST','ended_IST','duration_sec','late_flag','answers_json']);
  shR.appendRow([
    fmtIST(endNow,'dd/MM/YYYY HH:mm:ss'), name, email, code, ip, ua,
    fmtIST(startedAt,'dd/MM/YYYY HH:mm:ss'), fmtIST(endNow,'dd/MM/YYYY HH:mm:ss'), durationSec, late, answers
  ]);

  // Master log
  const shM = SH('Master_Log');
  if (shM.getLastRow()===0) shM.appendRow(['date_IST (DD/MM/YYYY)','time_IST (HH:MM:SS)','test_id','test_title','student_name','student_email','code','duration_sec','late_flag']);
  shM.appendRow([
    fmtIST(endNow,'dd/MM/YYYY'), fmtIST(endNow,'HH:mm:ss'), testId, cfg.title, name, email, code, durationSec, late
  ]);

  // Mark code used
  shC.getRange(row, 2, 1, 6).setValues([[
    'USED', rec[idx['issued_at_IST']], rec[idx['started_at_IST']], fmtIST(endNow,'dd/MM/YYYY HH:mm:ss'), name, email
  ]]);

  // Notify admin
  if (cfg.admin_email){
    MailApp.sendEmail({
      to: cfg.admin_email,
      subject: `[Aptitude] Submission ${testId} by ${name}`,
      htmlBody: `Test: ${cfg.title}<br>Student: ${name} (${email})<br>Code: ${code}<br>Duration: ${durationSec}s<br>Late: ${late}`
    });
  }

  return { ok:true, message:'Submission recorded', late_flag: late, duration_sec: durationSec };
}

function parseIST(str){
  // str in dd/MM/YYYY HH:mm:ss
  const [d,m,y,hh,mm,ss] = str.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/).slice(1).map(Number);
  return new Date(y, m-1, d, hh, mm, ss);
}

// Expire IN_USE older than resume window
function expireStaleInUse(){
  const shCfg = SH('Config_Tests');
  const cfgRows = shCfg.getDataRange().getValues();
  const headers = cfgRows[0];
  const idx={}; headers.forEach((h,i)=>idx[h]=i);
  for (let r=1;r<cfgRows.length;r++){
    const testId = String(cfgRows[r][idx['test_id']]||'');
    if (!testId) continue;
    const ttl = Number(cfgRows[r][idx['code_ttl_min']]||RESUME_MINUTES_DEFAULT);
    const sh = codesSheet(testId);
    const data = sh.getDataRange().getValues();
    if (data.length<2) continue;
    const h2 = {}; data[0].forEach((h,i)=>h2[h]=i);
    for (let i=1;i<data.length;i++){
      if (String(data[i][h2['state']])==='IN_USE'){
        const started = data[i][h2['started_at_IST']];
        if (!started) continue;
        const startedAt = parseIST(started);
        const expireAt = new Date(startedAt.getTime()+ ttl*60*1000);
        if (now()>expireAt){
          sh.getRange(i+1, h2['state']+1).setValue('EXPIRED');
        }
      }
    }
  }
}

// Build Student_Index summary
function rebuildStudentIndex(){
  const sh = SH('Student_Index');
  sh.clear();
  const cfgs = SH('Config_Tests').getDataRange().getValues();
  const headers = cfgs[0];
  const idx={}; headers.forEach((h,i)=>idx[h]=i);
  const tests = [];
  for (let r=1;r<cfgs.length;r++){
    if ((cfgs[r][idx['test_id']]||'').toString().trim()) tests.push({ id: cfgs[r][idx['test_id']], title: cfgs[r][idx['title']] });
  }
  const map = new Map(); // email -> { name, perTest }
  tests.forEach(t=>{
    const shR = responsesSheet(t.id);
    const vals = shR.getDataRange().getValues();
    if (vals.length<2) return;
    const h2 = {}; vals[0].forEach((h,i)=>h2[h]=i);
    for (let i=1;i<vals.length;i++){
      const email = String(vals[i][h2['student_email']]||'').trim();
      if (!email) continue;
      const name = String(vals[i][h2['student_name']]||'').trim();
      const ts = String(vals[i][h2['timestamp_IST']]||'');
      const obj = map.get(email) || { name, per:{} };
      obj.per[t.id] = 'Taken ('+ts+')';
      obj.name = obj.name || name;
      map.set(email, obj);
    }
  });
  // headers
  const headerRow = ['student_email','student_name', ...tests.map(t=>t.id+' – '+t.title)];
  const rows=[headerRow];
  for (const [email,info] of map.entries()){
    const row=[email, info.name];
    tests.forEach(t=> row.push(info.per[t.id]||'Not taken'));
    rows.push(row);
  }
  if (rows.length===1) rows.push(['','','']);
  sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
}
