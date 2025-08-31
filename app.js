const qs = s => document.querySelector(s);
const startForm = qs('#start-form');
const startSection = qs('#start-section');
const testSection = qs('#test-section');
const qWrap = qs('#questions');
const timerEl = qs('#timer');
const titleEl = qs('#test-heading');
const submitBtn = qs('#submit-btn');
const statusEl = qs('#status');

// Populate the test dropdown from backend
const testSelect = qs('#test-select');
(async function initTests(){
  try {
    const r = await fetch(window.BACKEND_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ action: 'config' })
    });
    const j = await r.json();
    if (!j.ok) throw new Error(j.error || 'Config failed');

    testSelect.innerHTML = '';
    j.tests.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t.test_id;
      opt.textContent = `${t.title} (${t.test_id})`;
      testSelect.appendChild(opt);
    });
  } catch (e) {
    console.error(e);
    alert('Could not load tests. Please try again later.');
  }
})();

// Success screen on reload
const params = new URLSearchParams(window.location.search);
if (params.get("submitted") === "1") {
  startSection.classList.add("hidden");
  testSection.classList.remove("hidden");

  const late = params.get("late") === "Y";
  const name = params.get("name") || "Student";

  testSection.innerHTML = `
    <h2>✅ Test submitted</h2>
    <p>Thanks, ${name}. Your responses have been recorded.</p>
    ${late ? `<p><strong>Note:</strong> Marked as late on the server.</p>` : ""}
  `;

  // Clean up the URL so refresh doesn’t keep showing success
  history.replaceState({}, "", window.location.pathname);
}

let remaining = 0;
let tickHandle = null;
let testCtx = { test_id: '', code: '', name: '', email: '', questions: [] };
let isSubmitting = false;

// Light deterrents
window.addEventListener('contextmenu', e => e.preventDefault());
window.addEventListener('copy', e => e.preventDefault());
window.addEventListener('cut', e => e.preventDefault());
window.addEventListener('paste', e => e.preventDefault());

startForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const fd = new FormData(startForm);
  testCtx.name = fd.get('name').trim();
  testCtx.email = fd.get('email').trim();
  testCtx.test_id = fd.get('test_id');
  testCtx.code = fd.get('code').trim();

  const ip = await getIPSafe();
  const res = await post('start', {
    name: testCtx.name,
    email: testCtx.email,
    test_id: testCtx.test_id,
    code: testCtx.code,
    ip,
    ua: navigator.userAgent
  });
  if (!res.ok){ alert(res.error || 'Failed to start'); return; }

  startSection.classList.add('hidden');
  testSection.classList.remove('hidden');

  titleEl.textContent = `${res.title} (${res.test_id}, ${res.time_limit_min} min)`;
  remaining = res.remainingSec;
  testCtx.questions = Array.isArray(res.questions) ? res.questions : [];
  renderQuestions(testCtx.questions);
  startTimer();
});

function renderQuestions(questions){
  qWrap.innerHTML = '';

  questions.forEach(q => {
    const box = document.createElement('div');
    box.className = 'question';

    const h = document.createElement('div');
    h.className = 'qtext';
    h.textContent = `Q${q.qno}. ${q.text}`;
    box.appendChild(h);

    // Show image via Apps Script endpoint
    if (q.imageId) {
      const img = document.createElement('img');
      img.className = 'qimg';
      // cache-buster helps when you just redeployed the web app
      const t = Date.now();
      img.src = `${window.BACKEND_URL}?action=image&id=${encodeURIComponent(q.imageId)}&t=${t}`;
      img.alt = `Image for Q${q.qno}`;
      box.appendChild(img);

      // (optional) clickable debug link below the image
      const dbg = document.createElement('a');
      dbg.href = img.src;
      dbg.target = "_blank";
      dbg.textContent = "Open image";
      dbg.style.display = "inline-block";
      dbg.style.fontSize = "12px";
      dbg.style.color = "#555";
      box.appendChild(dbg);
    }

    // Five options A–E
    (q.options || []).forEach((opt, idx) => {
      const label = ['A','B','C','D','E'][idx];
      const row = document.createElement('label');
      row.className = 'option';

      const input = document.createElement('input');
      input.type = 'radio';
      input.name = 'q_' + q.qno;
      input.value = label;

      const span = document.createElement('span');
      span.textContent = `${label}) ${opt}`;

      row.appendChild(input);
      row.appendChild(span);
      box.appendChild(row);
    });

    qWrap.appendChild(box);
  });
}

function startTimer(){
  updateTimerUI();
  tickHandle = setInterval(()=>{
    remaining = Math.max(0, remaining-1);
    updateTimerUI();
    if (remaining === 0){
      clearInterval(tickHandle);
      forceSubmit('Time is up. Your answers are being submitted.');
    }
  }, 1000);
}

function updateTimerUI(){
  const m = Math.floor(remaining/60).toString().padStart(2,'0');
  const s = (remaining%60).toString().padStart(2,'0');
  timerEl.textContent = `${m}:${s}`;
  if (remaining <= 30) timerEl.classList.add('warn');
}

submitBtn.addEventListener('click', ()=> forceSubmit('Submitting your answers…'));

async function forceSubmit(msg){
  if (isSubmitting) return;
  isSubmitting = true;
  if (tickHandle) { clearInterval(tickHandle); tickHandle = null; }
  statusEl.textContent = msg;
  submitBtn.disabled = true;
  const answers = collectAnswers();
  const ip = await getIPSafe();
  const res = await post('submit', {
    name: testCtx.name,
    email: testCtx.email,
    test_id: testCtx.test_id,
    code: testCtx.code,
    answers_json: JSON.stringify(answers),
    client_ended_at: new Date().toISOString(),
    ip,
    ua: navigator.userAgent
  });
  if (!res.ok) {
  statusEl.textContent = res.error || "Submission failed";
  submitBtn.disabled = false;
  isSubmitting = false;
  return;
}

// Redirect to a clean success page on the same path
const params = new URLSearchParams();
params.set("submitted", "1");
params.set("late", res.late_flag === "Y" ? "Y" : "N");
params.set("name", testCtx.name || "Student");
window.location.href = `${window.location.pathname}?${params.toString()}`;
  return;
}

function collectAnswers(){
  const out = [];
  testCtx.questions.forEach(q =>{
    const picked = document.querySelector(`input[name="q_${q.qno}"]:checked`);
    out.push({ qno: q.qno, choice: picked ? picked.value : '' });
  });
  return out;
}

async function post(action, body){
  const form = new URLSearchParams({ action, ...body });
  try {
    const r = await fetch(window.BACKEND_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: form.toString()
    });
    return await r.json();
  } catch(e){
    return { ok:false, error: e.message };
  }
}

async function getIPSafe(){
  try {
    const r = await fetch('https://api.ipify.org?format=json');
    const j = await r.json();
    return j.ip || '';
  } catch(e){
    return '';
  }
}
