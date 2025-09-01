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

function _inspectPublishedIndex() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Published_Index');
  const vals = sh.getDataRange().getValues(); // [test_id, file_id, ...]
  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const testId = String(vals[r][0] || '').trim();
    const fid = String(vals[r][1] || '').trim();
    if (!testId || !fid) continue;
    try {
      const f = DriveApp.getFileById(fid);
      const mime = f.getMimeType();
      const head = f.getBlob().getDataAsString().slice(0, 20);
      out.push({ testId, fileId: fid, name: f.getName(), mime, head });
    } catch (err) {
      out.push({ testId, fileId: fid, error: String(err) });
    }
  }
  Logger.log(JSON.stringify(out, null, 2));
}

function _repairPublishedIndexKeepLatest() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Published_Index');
  const vals = sh.getDataRange().getValues(); // header + rows
  if (vals.length < 2) return;

  const header = vals[0];
  // Build map from bottom -> top so the latest row per test_id is kept
  const keep = new Map();
  for (let r = vals.length - 1; r >= 1; r--) {
    const row = vals[r];
    const testId = String(row[0] || '').trim();
    if (!testId) continue;
    if (!keep.has(testId)) keep.set(testId, row);
  }

  // Rebuild rows: header + kept latest rows (sorted by test_id for neatness)
  const rows = [header, ...Array.from(keep.entries())
    .sort((a,b)=> a[0].localeCompare(b[0]))
    .map(([,row]) => row)];

  sh.clearContents();
  sh.getRange(1, 1, rows.length, header.length).setValues(rows);
  SpreadsheetApp.flush();
}

function appVersion_() {
  return 'img-b64-v1'; // bump this string each time you deploy
}

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
  const sh = SH('Published_Index');
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) throw new Error('No published index yet');

  for (let r = vals.length - 1; r >= 1; r--) { // newest first
    const idInRow = (vals[r][0] || '').toString().trim();
    if (idInRow !== testId) continue;

    const fileId = (vals[r][1] || '').toString().trim();
    if (!fileId) continue;

    try {
      // Try to parse as JSON
      const obj = readJsonFromDrive_(fileId);
      // minimal shape check
      if (obj && typeof obj === 'object' && obj.questions) return obj;
    } catch (e) {
      // If this row was a Google Doc (PDF blob) or corrupt, skip and try the next match above
      continue;
    }
  }
  throw new Error('Not found usable JSON in Published_Index for ' + testId);
}

function setTestJson(testId, obj) {
  const fileId = writeJsonToDrive_('Aptitude_Published_JSON', `${testId}.json`, obj);
  const sh = SH('Published_Index');
  if (sh.getLastRow() === 0) sh.appendRow(['test_id','file_id','size_bytes','updated_at_IST']);

  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let r=1; r<vals.length; r++) {
    if ((vals[r][0]||'').toString().trim() === testId) { row = r+1; break; }
  }

  const size = JSON.stringify(obj).length;
  const nowStr = fmtIST(new Date(), 'dd/MM/yyyy HH:mm:ss');

  if (row < 0) {
    sh.appendRow([testId, fileId, size, nowStr]);
  } else {
    sh.getRange(row, 2, 1, 3).setValues([[fileId, size, nowStr]]);
  }
}

function jsonSheet(name) {
  return S().getSheetByName(name) || S().insertSheet(name);
}

function writeJsonRow(sheetName, testId, obj) {
  const sh = jsonSheet(sheetName);
  if (sh.getLastRow() === 0) sh.appendRow(['test_id','json']);
  const data = sh.getDataRange().getValues();
  const idx = { test_id: 0, json: 1 };
  let row = -1;
  for (let r=1; r<data.length; r++) {
    if ((data[r][idx.test_id]||'').toString().trim() === testId) { row = r+1; break; }
  }
  const payload = JSON.stringify(obj);
  if (row < 0) {
    sh.appendRow([testId, payload]);
  } else {
    sh.getRange(row, idx.json+1).setValue(payload);
  }
}

function readJsonRow(sheetName, testId) {
  const sh = jsonSheet(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error('No data in '+sheetName);
  const idx = { test_id: 0, json: 1 };
  for (let r=1; r<data.length; r++) {
    if ((data[r][idx.test_id]||'').toString().trim() === testId) {
      const raw = data[r][idx.json];
      if (!raw) throw new Error('Empty JSON for '+testId+' in '+sheetName);
      return JSON.parse(raw);
    }
  }
  throw new Error('Not found: '+testId+' in '+sheetName);
}

function ensureFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function writeJsonToDrive_(folderName, filename, obj) {
  const folder = ensureFolder_(folderName);
  const files = folder.getFilesByName(filename);
  const content = JSON.stringify(obj);
  if (files.hasNext()) {
    const file = files.next();
    file.setTrashed(false);
    file.setContent(content);
    return file.getId();
  } else {
    return folder.createFile(filename, content, MimeType.PLAIN_TEXT).getId();
  }
}
function readJsonFromDrive_(fileId) {
  const file = DriveApp.getFileById(fileId);
  return JSON.parse(file.getBlob().getDataAsString());
}

// ===== Admin menu =====
function onOpen(){
  SpreadsheetApp.getUi().createMenu('Aptitude Admin')
    .addItem('Import From Google Doc…','menu_import')
    .addItem('Publish Test JSON','menu_publish')
    .addItem('Generate Codes…','menu_generateCodes')
    .addItem('Expire Locks','menu_expireLocks')
    .addItem('Rebuild Student Index','menu_rebuildStudentIndex')
    .addItem('Validate Published JSON','menu_validatePublished')
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
  if (sh.getLastRow()===0){
    sh.appendRow(['code','state','issued_at_IST','started_at_IST','submitted_at_IST','used_by_name','used_by_email','resume_until_IST']);
  }
  // Force text format for code column
  sh.getRange(1,1,sh.getMaxRows(),1).setNumberFormat('@');

  count = Number(count);
  if (!Number.isFinite(count) || count <= 0) throw new Error('Please enter a positive number for count');

  // Collect existing codes normalized to 6 digits
  const last = sh.getLastRow();
  let existing = new Set();
  if (last > 1){
    const vals = sh.getRange(2,1,last-1,1).getValues().flat().filter(Boolean);
    existing = new Set(vals.map(v => String(v).replace(/[^0-9]/g,'').padStart(6,'0')));
  }

  const rows=[];
  while(rows.length < count){
    const code = ('000000'+Math.floor(Math.random()*1e6)).slice(-6);
    if (!existing.has(code)){
      rows.push(["'"+code, 'AVAILABLE', fmtIST(now(),'dd/MM/yyyy HH:mm:ss'), '', '', '', '', '']); // prefix ' to force text
      existing.add(code);
    }
  }
  const start = sh.getLastRow()+1;
  sh.getRange(start,1,rows.length,rows[0].length).setValues(rows);
  sh.getRange(start,1,rows.length,1).setNumberFormat('@'); // keep text
}

// ===== Import & publish =====
function importFromGoogleDoc(testId){
  const cfg = getConfigRow(testId);
  const docId = (cfg.doc_url.match(/\/d\/([A-Za-z0-9_-]+)/)||[])[1] || cfg.doc_url;

  // Read the Doc as plain text
  const body = DocumentApp.openById(docId).getBody().getText();

  // Normalize lines: trim, drop empties
  const lines = body.split(/\r?\n/).map(s => s.trim()).filter(Boolean);

  // Title: first line starting with "#", else fall back to config title
  const titleLine = lines[0] && lines[0].startsWith('#')
    ? lines.shift().replace(/^#\s*/, '')
    : (cfg.title || ('Test ' + testId));

  const questions = [];
  let cursor = 0, qno = 0;

  while (cursor < lines.length) {
    // Look for a question header like "Q1. ..." / "Q12. ..."
    if (!/^Q\d+\./.test(lines[cursor])) { cursor++; continue; }

    const qHeader = lines[cursor++]; // consume the Q#. line
    qno++;
    const qText = qHeader.replace(/^Q\d+\.\s*/, '').trim();

    // Optional image line: "IMG: <DriveFileID>"
    let imageId = '';
    if (cursor < lines.length && /^IMG:/.test(lines[cursor])) {
      imageId = lines[cursor].replace(/^IMG:\s*/, '').trim(); // keep ONLY the file ID
      cursor++;
    }

    // Collect exactly 5 options labeled A) ... E)
    const opts = [];
    for (let k = 0; k < 5 && cursor < lines.length; k++) {
      const m = lines[cursor].match(/^[A-E]\)\s*(.*)$/);
      if (!m) break;
      opts.push(m[1]);
      cursor++;
    }
    if (opts.length !== 5) {
      throw new Error('Question ' + qno + ' does not have 5 options (A–E).');
    }

    questions.push({ qno, text: qText, imageId, options: opts });
  }

  // Persist a compact import payload; publisher will move this to Drive JSON
  writeJsonRow('Imported_JSON', testId, { title: titleLine, questions });
}

function publishTestJson(testId){
  // If you already have Imported_JSON sheet helpers:
  const imported = readJsonRow('Imported_JSON', testId);
  setTestJson(testId, imported);
}

function clearOldProperties() {
  const p = PropertiesService.getDocumentProperties();
  const all = p.getProperties();
  let n=0;
  Object.keys(all).forEach(k => {
    if (k.startsWith('IMPORTED_') || k.startsWith('TEST_JSON_')) {
      p.deleteProperty(k);
      n++;
    }
  });
  Logger.log('Deleted '+n+' old properties.');
}

// 1) Route GET requests
// DEBUG: returns metadata so we can see what the ID points to
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'image_b64')   return imageBase64_(e);
  if (action === 'debug_image') return debugImageSimple_(e);
  if (action === 'version')     return ContentService
                                   .createTextOutput(appVersion_())
                                   .setMimeType(ContentService.MimeType.TEXT);
  // (Optional) keep the old image endpoint off or return a hint:
  if (action === 'image') return ContentService
                             .createTextOutput('Use action=image_b64')
                             .setMimeType(ContentService.MimeType.TEXT);

  return ContentService.createTextOutput('OK default');
}

function selfTest_(){
  const report = { ok:true, steps:[] };

  // 1) config
  try {
    const tests = listActiveTests();
    report.steps.push({ step:'config', ok:true, n:tests.length });
    if (!tests.length) throw new Error('No ACTIVE tests');
    // pick first test
    const t = tests[0].test_id;
    const obj = getTestJson(t);
    if (!obj || !obj.title || !Array.isArray(obj.questions)) throw new Error('Bad JSON shape');
    report.steps.push({ step:'json', ok:true, test_id:t, title:obj.title, q:obj.questions.length });

    // optional: if first q has imageId, try base64
    const imgId = (obj.questions.find(q=>q.imageId) || {}).imageId;
    if (imgId){
      const f = DriveApp.getFileById(imgId);
      const mime = f.getMimeType();
      const b64 = Utilities.base64Encode(f.getBlob().getBytes()).slice(0,24);
      report.steps.push({ step:'image', ok:true, mime, sample:b64+'...' });
    } else {
      report.steps.push({ step:'image', ok:true, note:'no imageId found' });
    }
  } catch(e){
    report.ok = false;
    report.steps.push({ step:'error', ok:false, error:String(e) });
  }

  return ContentService.createTextOutput(JSON.stringify(report, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

// Streams an image from Drive by ID
// Handles: plain files, Drive shortcuts, Google files export (Drawings, Docs images not supported)
// Serve ONLY real image files (PNG/JPG/GIF) by ID using DriveApp


function imageHandlerSimple_(e) {
  const id = (e.parameter && e.parameter.id || '').trim();
  if (!id) {
    return ContentService.createTextOutput('Missing id')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  try {
    const file = DriveApp.getFileById(id);
    const mime = file.getMimeType();

    // Only serve real image files
    if (!mime || !mime.startsWith('image/')) {
      return ContentService.createTextOutput('Not an image file')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    // Convert to a fresh Blob explicitly (most robust)
    const src = file.getBlob();
    const out = Utilities.newBlob(src.getBytes(), mime, file.getName());
    return out; // <-- returning a Blob is a supported return type for doGet
  } catch (err) {
    return ContentService.createTextOutput('Not found')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}


// Optional: debug without using Drive.Files.* (no Advanced Drive needed)
function debugImageSimple_(e) {
  const id = (e.parameter && e.parameter.id || '').trim();
  const result = { ok: true, id, whoAmI: null, steps: [] };
  try { result.whoAmI = Session.getEffectiveUser().getEmail(); } catch (err) {}
  try {
    const f = DriveApp.getFileById(id);
    result.steps.push({
      step: 'DriveApp.getFileById',
      name: f.getName(),
      mimeType: f.getMimeType(),
      sizeBytes: f.getSize ? f.getSize() : null,
      isImage: f.getMimeType().startsWith('image/')
    });
  } catch (err) {
    result.steps.push({ step: 'DriveApp.getFileById', error: String(err) });
  }
  return ContentService.createTextOutput(JSON.stringify(result, null, 2))
                       .setMimeType(ContentService.MimeType.JSON);
}

function imageBase64_(e) {
  const id = (e.parameter && e.parameter.id || '').trim();
  if (!id) return json({ ok:false, error:'Missing id' });
  try {
    const f = DriveApp.getFileById(id);
    const mime = f.getMimeType();
    if (!mime || !mime.startsWith('image/')) {
      return json({ ok:false, error:'Not an image file' });
    }
    const b64 = Utilities.base64Encode(f.getBlob().getBytes());
    return json({ ok:true, mime, b64 });
  } catch (err) {
    return json({ ok:false, error:'Not found' });
  }
}


// ===== Rate limiting (best-effort using client IP from frontend) =====
// Soften & scope the limiter
function checkRateLimit(p, maxPer10){
  // p should include: ip, test_id, code, email
  const ip = (p.ip || 'noip').trim();
  const test = (p.test_id || 'notest').trim();
  const code = (p.code || '').trim();
  const mail = (p.email || '').trim();

  // Build a key that spreads users sharing an IP
  // Prefer code if provided; else email; else just IP+test
  const who = code ? ('code:' + code)
            : (mail ? ('mail:' + mail.toLowerCase()) : ('ip:' + ip));

  const bucket = Math.floor(Date.now()/(10*60*1000));         // 10-min window
  const key = ['rl', test, who, bucket].join(':');            // rl:T3:code:123456:1234567

  const cache = CacheService.getScriptCache();
  const curr = Number(cache.get(key) || '0');

  if (curr >= maxPer10) {
    throw new Error('Too many attempts. Please wait a few minutes and try again.');
  }
  cache.put(key, String(curr+1), 600); // 600s = 10min
}

// ===== Web app endpoints =====
function doPost(e){
  const params = e.parameter || {};
  const action = params.action || '';
  try {
    if (action==='start') return json(startHandler(params));
    if (action==='submit') return json(submitHandler(params));
    if (action==='config') return json({ ok:true, tz:TZ, tests: listActiveTests() }); // <-- add tests
    throw new Error('Unknown action');
  } catch(err){
    return json({ ok:false, error: String(err) });
  }
}

function listActiveTests(){
  const sh = SH('Config_Tests');
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const h = {}; vals[0].forEach((k,i)=>h[k]=i);
  const out = [];
  for (let r=1; r<vals.length; r++){
    const row = vals[r];
    if ((row[h['status']]||'').toString().toUpperCase() !== 'ACTIVE') continue;
    const test_id = (row[h['test_id']]||'').toString().trim();
    if (!test_id) continue;
    out.push({
      test_id,
      title: (row[h['title']]||'').toString(),
      time_limit_min: Number(row[h['time_limit_min']]||0)
    });
  }
  return out;
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function startHandler(p){
  const name   = (p.name   || '').trim();
  const email  = (p.email  || '').trim();
  const testId = (p.test_id|| '').trim();
  const code   = (p.code   || '').trim();
  const ip     = (p.ip     || '').trim();
  const ua     = (p.ua     || '').trim();

  if (!name || !email || !testId || !code) throw new Error('Missing fields');

  // Config (time limit, resume window, status, rate limit, etc.)
  const cfg = getConfigRow(testId); // expects: time_limit_min, status, code_ttl_min, rate_limit_per_10min, allow_resume
  if ((cfg.status || '').toString().toUpperCase() !== 'ACTIVE') throw new Error('Test not active');

  // Best-effort rate limit by IP
  checkRateLimit({ ip, test_id: testId, code, email }, cfg.rate_limit_per_10min);

  // Locate the code row
  const sh = codesSheet(testId);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error('No codes exist for this test');

  const headers = data[0];
  const H = {};
  headers.forEach((h,i)=> H[h]=i);

  let row = -1, rec = null;
  for (let r=1; r<data.length; r++){
    if (String(data[r][H['code']]||'') === code){
      row = r+1; rec = data[r]; break;
    }
  }
 // If code not found → count as a failed attempt (anti-bruteforce)
  if (row < 0) {
    // Only failed lookups increment the limiter
    checkRateLimit({ ip, test_id: testId, code, email }, cfg.rate_limit_per_10min);
    throw new Error('Invalid code');
  }


  // Compute or set the official start time
  let serverStart = null;

  if (state === 'AVAILABLE') {
    // Move to IN_USE and stamp start time + resume window
    const issuedAt = rec[H['issued_at_IST']] || nowStr;
    const resumeUntilDate = new Date(nowD.getTime() + (Number(cfg.code_ttl_min)||15)*60*1000);
    const resumeUntilStr  = fmtIST(resumeUntilDate, 'dd/MM/yyyy HH:mm:ss');

    // Update columns: state, issued_at_IST, started_at_IST, submitted_at_IST, used_by_name, used_by_email, resume_until_IST
    sh.getRange(row, H['state']+1, 1, 7).setValues([[
      'IN_USE',
      issuedAt,
      nowStr,
      '',              // submitted_at_IST stays empty at start
      name,
      email,
      resumeUntilStr
    ]]);

    serverStart = nowD;

  } else if (state === 'IN_USE') {
    // Resume only within the allowed window
    const startedAtStr   = rec[H['started_at_IST']] || nowStr;
    const resumeUntilStr = rec[H['resume_until_IST']] || '';
    const startedAt      = parseIST(startedAtStr);
    const resumeUntil    = resumeUntilStr ? parseIST(resumeUntilStr)
                                          : new Date(startedAt.getTime() + (Number(cfg.code_ttl_min)||15)*60*1000);

    if (nowD > resumeUntil) {
      // Strict 15-min resume window enforcement
      throw new Error('Resume window over for this code');
    }
    // (Optionally refresh name/email on resume so logs have latest)
    if (name)  sh.getRange(row, H['used_by_name'] +1).setValue(name);
    if (email) sh.getRange(row, H['used_by_email']+1).setValue(email);

    serverStart = startedAt;

  } else if (state === 'USED') {
    throw new Error('Code already used');
  } else if (state === 'EXPIRED') {
    throw new Error('Code expired');
  } else {
    throw new Error('Invalid code state');
  }

  // Compute remaining time based on serverStart
  const endAtMs      = serverStart.getTime() + (Number(cfg.time_limit_min)||0)*60*1000;
  const remainingSec = Math.max(0, Math.floor((endAtMs - nowD.getTime()) / 1000));

  // Load test JSON (title + questions). With Option B this reads from Drive.
  const test = getTestJson(testId); // expects { title, questions }

  return {
    ok: true,
    test_id: testId,
    title: test.title,
    time_limit_min: Number(cfg.time_limit_min)||0, // <-- important for heading
    remainingSec,
    serverNow: nowStr,
    endAt: fmtIST(new Date(endAtMs), 'dd/MM/yyyy HH:mm:ss'),
    questions: test.questions
  };
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

  const startedAt = parseIST(rec[idx['started_at_IST']]||fmtIST(now(),'dd/MM/yyyy HH:mm:ss'));
  const endDeadline = new Date(startedAt.getTime() + cfg.time_limit_min*60*1000);
  const endNow = now();
  const late = endNow > endDeadline ? 'Y' : 'N';
  const durationSec = Math.max(0, Math.floor((endNow.getTime()-startedAt.getTime())/1000));

  // Write response row
  const shR = responsesSheet(testId);
  if (shR.getLastRow()===0) shR.appendRow(['timestamp_IST','student_name','student_email','code','client_ip','user_agent','started_IST','ended_IST','duration_sec','late_flag','answers_json']);
  shR.appendRow([
    fmtIST(endNow,'dd/MM/yyyy HH:mm:ss'), name, email, code, ip, ua,
    fmtIST(startedAt,'dd/MM/yyyy HH:mm:ss'), fmtIST(endNow,'dd/MM/yyyy HH:mm:ss'), durationSec, late, answers
  ]);

  // Master log
  const shM = SH('Master_Log');
  if (shM.getLastRow()===0) shM.appendRow(['date_IST (DD/MM/yyyy)','time_IST (HH:MM:SS)','test_id','test_title','student_name','student_email','code','duration_sec','late_flag']);
  shM.appendRow([
    fmtIST(endNow,'dd/MM/yyyy'), fmtIST(endNow,'HH:mm:ss'), testId, cfg.title, name, email, code, durationSec, late
  ]);

  // Mark code used
  shC.getRange(row, 2, 1, 6).setValues([[
    'USED', rec[idx['issued_at_IST']], rec[idx['started_at_IST']], fmtIST(endNow,'dd/MM/yyyy HH:mm:ss'), name, email
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

function parseIST(val){
  // Already a Date?
  if (val instanceof Date) return val;

  // Numeric timestamps or Excel/Sheets serials (rare in Apps Script, but safe)
  if (typeof val === 'number') {
    const d = new Date(val);
    if (!isNaN(d)) return d;
  }

  // Strings: try DD/MM/yyyy HH:mm[:ss]
  if (typeof val === 'string') {
    const s = val.trim();
    const m = s.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?/);
    if (m) {
      const [ , dd, mm, yyyy, HH, MM, SS ] = m;
      return new Date(Number(yyyy), Number(mm)-1, Number(dd), Number(HH), Number(MM), Number(SS||0));
    }
    // Fallback: let Date parse it (e.g., ISO)
    const d = new Date(s);
    if (!isNaN(d)) return d;
  }

  // Last resort: now()
  return new Date();
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

function normalizeCodes(testId){
  const sh = codesSheet(testId);
  if (sh.getLastRow() <= 1) return;
  // Force text for the whole column
  sh.getRange(1,1,sh.getMaxRows(),1).setNumberFormat('@');

  const last = sh.getLastRow();
  const rng = sh.getRange(2,1,last-1,1);
  const vals = rng.getValues();

  for (let i=0; i<vals.length; i++){
    const raw = vals[i][0];
    if (!raw) continue;
    const s = String(raw).replace(/[^0-9]/g,'');
    vals[i][0] = "'"+s.padStart(6, '0');  // write as text with leading zeros
  }
  rng.setValues(vals);
}

function whoAmI(){ Logger.log(Session.getEffectiveUser().getEmail()); }


// --- Helpers to ensure plain-text files, not Google Docs ---
function _trashNonPlainWithSameName_(folder, filename) {
  // Only trash files with the exact same name in that folder that are NOT plain text
  let trashed = 0;
  const it = folder.getFilesByName(filename);
  while (it.hasNext()) {
    const f = it.next();
    if (f.getMimeType() !== MimeType.PLAIN_TEXT) {
      f.setTrashed(true);
      trashed++;
    }
  }
  return trashed;
}

function writeJsonToDrivePlain_(folderName, filename, obj) {
  const folder = ensureFolder_(folderName);
  const content = JSON.stringify(obj);

  // Prefer an existing *plain text* file with that name
  let plain = null;
  let it = folder.getFilesByName(filename);
  while (it.hasNext()) {
    const f = it.next();
    if (f.getMimeType() === MimeType.PLAIN_TEXT) { plain = f; break; }
  }

  if (!plain) {
    // If only Google Docs exist with same name, move them to Trash to avoid collisions
    _trashNonPlainWithSameName_(folder, filename);
    // Create a fresh plain text file
    plain = folder.createFile(filename, content, MimeType.PLAIN_TEXT);
  } else {
    plain.setTrashed(false);
    plain.setContent(content);
  }
  return plain.getId();
}

// Re-publish one test, forcing a plain-text JSON file and fixing Published_Index
function _forcePlainRepublish_(testId) {
  const imported = readJsonRow('Imported_JSON', testId); // {title, questions}
  const fileId = writeJsonToDrivePlain_('Aptitude_Published_JSON', `${testId}.json`, imported);

  const sh = SH('Published_Index');
  if (sh.getLastRow() === 0) sh.appendRow(['test_id','file_id','size_bytes','updated_at_IST']);

  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let r=1; r<vals.length; r++) {
    if ((vals[r][0]||'').toString().trim() === testId) { row = r+1; break; }
  }
  const size = JSON.stringify(imported).length;
  const nowStr = fmtIST(new Date(), 'dd/MM/yyyy HH:mm:ss');

  if (row < 0) {
    sh.appendRow([testId, fileId, size, nowStr]);
  } else {
    sh.getRange(row, 2, 1, 3).setValues([[fileId, size, nowStr]]);
  }
  Logger.log('Re-published %s to plain text file: %s', testId, fileId);
}

// Convenience: fix all tests listed in Published_Index
function _forcePlainRepublishAll_() {
  const sh = SH('Published_Index');
  const vals = sh.getDataRange().getValues();
  const seen = new Set();
  for (let r=1; r<vals.length; r++) {
    const id = (vals[r][0]||'').toString().trim();
    if (!id || seen.has(id)) continue;
    seen.add(id);
    try { _forcePlainRepublish_(id); } catch (e) {
      Logger.log('Failed republish for %s: %s', id, e);
    }
  }
}

// (Optional) Your existing inspector; keep using it
function _inspectPublishedIndex() {
  const sh = SH('Published_Index');
  const vals = sh.getDataRange().getValues();
  const out = [];
  for (let r=1; r<vals.length; r++) {
    const testId = String(vals[r][0]||'').trim();
    const fileId = String(vals[r][1]||'').trim();
    if (!testId || !fileId) continue;
    try {
      const f = DriveApp.getFileById(fileId);
      const blob = f.getBlob();
      const bytes = blob.getBytes();
      const head = Utilities.newBlob(bytes.slice(0, 20)).getDataAsString(); // first few bytes
      out.push({
        testId, fileId, name: f.getName(), mime: f.getMimeType(),
        head: head
      });
    } catch (e) {
      out.push({ testId, fileId, error: String(e) });
    }
  }
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

function republishT1() {
  _forcePlainRepublish_('T1');
}

function republishT3() {
  _forcePlainRepublish_('T3');
}


/** Menu item */
function menu_validatePublished(){ 
  const res = validateAndFixPublishedIndex_();
  SpreadsheetApp.getUi().alert(res.msg);
}

/** Validate all rows; fix non-plain files by recreating as text/plain */
function validateAndFixPublishedIndex_(){
  const sh = SH('Published_Index');
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { ok:true, fixed:0, msg:'Nothing to validate (no rows).' };

  const headers = vals[0]; const H = {};
  headers.forEach((h,i)=>H[h]=i);

  let fixed = 0, checked = 0;
  for (let r=1; r<vals.length; r++){
    const testId = String(vals[r][H.test_id]||'').trim();
    const fileId = String(vals[r][H.file_id]||'').trim();
    if (!testId || !fileId) continue;

    checked++;
    let mime = '';
    try { mime = DriveApp.getFileById(fileId).getMimeType() || ''; } catch(e){ mime=''; }

    if (mime !== MimeType.PLAIN_TEXT){
      // read the JSON object from wherever we store imports
      let obj;
      try {
        obj = readJsonRow('Imported_JSON', testId);
      } catch(e) {
        return { ok:false, fixed, msg:`Cannot read Imported_JSON for ${testId}: ${e}` };
      }
      // recreate as a fresh plain text .json
      const newId = writeJsonToDrive_('Aptitude_Published_JSON', `${testId}.json`, obj);
      const size = JSON.stringify(obj).length;
      const nowStr = fmtIST(new Date(), 'dd/MM/yyyy HH:mm:ss');
      sh.getRange(r+1, H.file_id+1, 1, 3).setValues([[newId, size, nowStr]]);
      fixed++;
    }
  }
  return { ok:true, fixed, msg:`Checked ${checked} test(s). Fixed ${fixed}.` };
}

function setTestJson(testId, obj) {
  // Always (re)create a plain text file
  const fileId = writeJsonToDrive_('Aptitude_Published_JSON', `${testId}.json`, obj);

  const sh = SH('Published_Index');
  if (sh.getLastRow() === 0) sh.appendRow(['test_id','file_id','size_bytes','updated_at_IST']);

  const vals = sh.getDataRange().getValues();
  let row = -1;
  for (let r=1; r<vals.length; r++) {
    if (String(vals[r][0]).trim() === testId) { row = r+1; break; }
  }

  const size = JSON.stringify(obj).length;
  const nowStr = fmtIST(new Date(), 'dd/MM/yyyy HH:mm:ss');

  if (row < 0) {
    sh.appendRow([testId, fileId, size, nowStr]);
  } else {
    sh.getRange(row, 2, 1, 3).setValues([[fileId, size, nowStr]]);
  }
}