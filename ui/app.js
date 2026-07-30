'use strict';

let _running = false;
let _es      = null;
let _toastTimer = null;
let _autoscroll = true;
let _lineCount = 0;
let _elapsedTimer = null;
let _startedAt = 0;

// ── File/folder browser ───────────────────────────────────────────────────────

async function browse(type, inputId) {
  const ep = type === 'file' ? '/api/browse/file' : '/api/browse/folder';
  try {
    const r = await fetch(ep, { method: 'POST' });
    const d = await r.json();
    if (d.path) {
      const el = document.getElementById(inputId);
      el.value = d.path;
      el.classList.add('filled');
      updateStepStates();
    }
  } catch (e) {
    showToast('Could not open dialog: ' + e.message);
  }
}

// ── Run agent ─────────────────────────────────────────────────────────────────

async function runAgent() {
  if (_running) return;

  const docs    = document.getElementById('in-docs').value.trim();
  const prompt  = document.getElementById('in-prompt').value.trim();
  const gallery = document.getElementById('in-gallery').value.trim();

  setRunning(true);
  resetLog();
  hideResult();

  try {
    const res  = await fetch('/api/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ docs, prompt, gallery }),
    });
    const data = await res.json();

    if (!res.ok) {
      showToast(data.error || 'Unknown error');
      setRunning(false);
      return;
    }

    setStatusChip('running', 'Job ' + data.job_id.slice(-6));

    if (_es) { _es.close(); _es = null; }
    _es = new EventSource(`/api/stream/${data.job_id}`);

    _es.onmessage = (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.type === 'log') {
        appendLog(msg.msg);
      } else if (msg.type === 'view') {
        handleView(msg);
      } else if (msg.type === 'done') {
        _es.close(); _es = null;
        setRunning(false);
        showResult(msg.ok, msg.msg);
        setStatusChip(msg.ok ? 'ready' : 'error', msg.ok ? 'Ready' : 'Failed');
      }
      // 'ping' — ignore
    };

    _es.onerror = () => {
      if (_es) { _es.close(); _es = null; }
      setRunning(false);
      appendLog('\n[stream error — connection closed]');
      setStatusChip('error', 'Connection lost');
    };

  } catch (e) {
    showToast('Network error: ' + e.message);
    setRunning(false);
  }
}

// ── Live document view ──────────────────────────────────────────────────────

let _docviewParaEls = [];

function resetDocView() {
  document.getElementById('docview-card').classList.add('hidden');
  document.getElementById('docview-paras').innerHTML = '';
  document.getElementById('docview-image-wrap').classList.add('hidden');
  document.getElementById('docview-image').src = '';
  document.getElementById('docview-filename').textContent = '';
  document.getElementById('docview-badge').innerHTML = '<span class="live-dot"></span>editing';
  _docviewParaEls = [];
}

function handleView(evt) {
  const card = document.getElementById('docview-card');
  if (evt.kind === 'open') {
    card.classList.remove('hidden');
    document.getElementById('docview-badge').innerHTML = '<span class="live-dot"></span>editing';
    document.getElementById('docview-filename').textContent = evt.doc || '';

    const container = document.getElementById('docview-paras');
    container.innerHTML = '';
    _docviewParaEls = (evt.paragraphs || []).map((text) => {
      const p = document.createElement('div');
      p.className = 'docview-para';
      p.textContent = text;
      container.appendChild(p);
      return p;
    });
    container.scrollTop = 0;

    const imgWrap = document.getElementById('docview-image-wrap');
    const img = document.getElementById('docview-image');
    if (evt.image) {
      img.src = evt.image;
      imgWrap.classList.remove('hidden');
    } else {
      imgWrap.classList.add('hidden');
    }
  } else if (evt.kind === 'replace_preview') {
    for (const c of (evt.changes || [])) {
      const el = _docviewParaEls[c.index];
      if (!el) continue;
      el.innerHTML = c.html;
      el.classList.add('flash');
      el.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  } else if (evt.kind === 'replace_settle') {
    for (const c of (evt.changes || [])) {
      const el = _docviewParaEls[c.index];
      if (!el) continue;
      el.textContent = c.text;
      el.classList.remove('flash');
    }
  } else if (evt.kind === 'image') {
    const img = document.getElementById('docview-image');
    const imgWrap = document.getElementById('docview-image-wrap');
    imgWrap.classList.remove('hidden');
    img.classList.add('swapping');
    setTimeout(() => {
      img.src = evt.image;
      img.classList.remove('swapping');
    }, 450);
  } else if (evt.kind === 'close') {
    document.getElementById('docview-badge').innerHTML = '<span class="live-dot" style="background:var(--green)"></span>saved';
  }
}

// ── Log rendering ─────────────────────────────────────────────────────────────

function appendLog(text) {
  const pre  = document.getElementById('log-pre');
  const wrap = document.getElementById('log-wrap');
  const lines = text.split('\n');
  const wasAtBottom = isNearBottom(wrap);

  for (const line of lines) {
    _lineCount++;
    const row = document.createElement('div');
    row.className = 'log-line';

    const ln = document.createElement('span');
    ln.className = 'log-ln';
    ln.textContent = _lineCount;

    const txt = document.createElement('span');
    txt.className = 'log-txt ' + classifyLine(line);
    txt.textContent = line;

    row.appendChild(ln);
    row.appendChild(txt);
    pre.appendChild(row);
  }

  document.getElementById('log-line-count').textContent = _lineCount + ' ln';

  if (_autoscroll && wasAtBottom) {
    wrap.scrollTop = wrap.scrollHeight;
  }
  updateScrollJump(wrap);
}

function classifyLine(line) {
  const t = line.trimStart();
  if (/^STEP\s+\d/i.test(t) || /^All finished!/i.test(t)) return 'l-step';
  if (/something went wrong|unexpected happened|had a problem/i.test(t)) return 'l-err';
  if (/^\(/.test(t) || /couldn'?t|isn'?t available/i.test(t)) return 'l-warn';
  if (/^✓/.test(t) || /^Done —/i.test(t) || /^All finished!/i.test(t) || /→/.test(t)) return 'l-ok';
  if (/^·/.test(t)) return 'l-muted';
  if (/^(Reading|Asking|Understood|Found|Looking|Opening|Changing|Swapping|Now working|The document now|Saving)/i.test(t)) return 'l-info';
  return 'l-default';
}

function resetLog() {
  const pre = document.getElementById('log-pre');
  const ph  = document.getElementById('log-placeholder');
  pre.innerHTML    = '';
  pre.style.display = 'block';
  ph.style.display  = 'none';
  _lineCount = 0;
  document.getElementById('log-line-count').textContent = '';
  updateScrollJump(document.getElementById('log-wrap'));
  resetDocView();
}

function clearLog() {
  const pre = document.getElementById('log-pre');
  const ph  = document.getElementById('log-placeholder');
  pre.innerHTML    = '';
  pre.style.display = 'none';
  ph.style.display  = 'flex';
  hideResult();
  document.getElementById('topbar-status').classList.add('hidden');
  _lineCount = 0;
  document.getElementById('log-line-count').textContent = '';
  setStatusChip('ready', 'Ready');
  document.getElementById('status-job').textContent = '';
  updateScrollJump(document.getElementById('log-wrap'));
  resetDocView();
}

// ── Console toolbar: autoscroll / copy / jump-to-bottom ────────────────────────

function isNearBottom(wrap) {
  return wrap.scrollHeight - wrap.scrollTop - wrap.clientHeight < 48;
}

function updateScrollJump(wrap) {
  const btn = document.getElementById('scroll-jump');
  const show = document.getElementById('log-pre').style.display === 'block' && !isNearBottom(wrap);
  btn.classList.toggle('hidden', !show);
}

function jumpToBottom() {
  const wrap = document.getElementById('log-wrap');
  wrap.scrollTop = wrap.scrollHeight;
  updateScrollJump(wrap);
}

function toggleAutoscroll() {
  _autoscroll = !_autoscroll;
  document.getElementById('autoscroll-btn').classList.toggle('active', _autoscroll);
  if (_autoscroll) jumpToBottom();
}

async function copyLog() {
  const rows = document.querySelectorAll('#log-pre .log-txt');
  const text = Array.from(rows).map(r => r.textContent).join('\n');
  if (!text) { showToast('Nothing to copy'); return; }
  try {
    await navigator.clipboard.writeText(text);
    showToast('Log copied to clipboard', true);
  } catch (e) {
    showToast('Could not copy: ' + e.message);
  }
}

// ── UI state ──────────────────────────────────────────────────────────────────

function setRunning(running) {
  _running = running;
  const btn    = document.getElementById('run-btn');
  const icon   = document.getElementById('run-icon');
  const label  = document.getElementById('run-label');
  const badge  = document.getElementById('live-badge');
  const status = document.getElementById('topbar-status');

  if (running) {
    btn.disabled = true;
    icon.style.animation = 'spin .8s linear infinite';
    icon.setAttribute('viewBox', '0 0 24 24');
    icon.innerHTML = '<path d="M21 12a9 9 0 1 1-6.219-8.56" stroke-linecap="round"/>';
    label.textContent = 'Running…';
    badge.classList.remove('hidden');
    status.classList.remove('hidden');
    _startedAt = Date.now();
    updateElapsed();
    _elapsedTimer = setInterval(updateElapsed, 1000);
  } else {
    btn.disabled = false;
    icon.style.animation = '';
    icon.innerHTML = '<polygon points="5 3 19 12 5 21 5 3"/>';
    label.textContent = 'Run Agent';
    badge.classList.add('hidden');
    status.classList.add('hidden');
    clearInterval(_elapsedTimer);
    _elapsedTimer = null;
  }
}

function updateElapsed() {
  const s = Math.floor((Date.now() - _startedAt) / 1000);
  const mm = String(Math.floor(s / 60)).padStart(2, '0');
  const ss = String(s % 60).padStart(2, '0');
  document.getElementById('topbar-elapsed').textContent = `Processing — ${mm}:${ss}`;
}

function setStatusChip(state, label) {
  const chip = document.getElementById('status-chip');
  chip.className = 'status-chip ' + state;
  document.getElementById('status-chip-label').textContent = label;
}

function showResult(ok, msg) {
  const el = document.getElementById('result');
  el.className = 'result-pill ' + (ok ? 'ok' : 'fail');
  const check = '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>';
  const cross  = '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>';
  el.innerHTML = (ok ? check : cross) + ' ' + escHtml(msg);
  el.classList.remove('hidden');
}

function hideResult() {
  document.getElementById('result').classList.add('hidden');
}

// ── Step indicator ──────────────────────────────────────────────────────────

function updateStepStates() {
  const map = [['step-tag-1', 'in-docs', '01'], ['step-tag-2', 'in-prompt', '02'], ['step-tag-3', 'in-gallery', '03']];
  for (const [tagId, inputId, num] of map) {
    const tag = document.getElementById(tagId);
    const filled = !!document.getElementById(inputId).value.trim();
    if (filled) {
      tag.classList.add('done');
      tag.innerHTML = '<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>';
    } else {
      tag.classList.remove('done');
      tag.textContent = num;
    }
  }
}

// ── Toast ─────────────────────────────────────────────────────────────────────

function showToast(msg, ok) {
  const el = document.getElementById('toast');
  document.getElementById('toast-text').textContent = msg;
  el.classList.toggle('info', !!ok);
  el.classList.remove('hidden');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.add('hidden'), 4500);
}

// ── Clock ────────────────────────────────────────────────────────────────────

function tickClock() {
  const now = new Date();
  document.getElementById('sidebar-date').textContent = now.toLocaleDateString(undefined, { month: 'short', day: '2-digit' });
  document.getElementById('sidebar-time').textContent = now.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ── Init ─────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  ['in-docs', 'in-prompt', 'in-gallery'].forEach(id => {
    document.getElementById(id).addEventListener('input', function () {
      this.classList.toggle('filled', !!this.value.trim());
      updateStepStates();
    });
  });

  document.getElementById('log-wrap').addEventListener('scroll', function () {
    updateScrollJump(this);
  });

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      runAgent();
    }
  });

  document.getElementById('autoscroll-btn').classList.add('active');

  tickClock();
  setInterval(tickClock, 1000 * 30);

  applyUrlParams();
});

// ── URL-param autofill/autorun (for scripted demos) ─────────────────────────
// ?docs=<path>&prompt=<path>&gallery=<path>&autorun=1 pre-fills the three
// fields and, with autorun=1, clicks Run Agent automatically — lets a demo
// be driven just by opening a URL, no manual interaction needed.
function applyUrlParams() {
  const params = new URLSearchParams(window.location.search);
  const map = { docs: 'in-docs', prompt: 'in-prompt', gallery: 'in-gallery' };
  let any = false;
  for (const [key, inputId] of Object.entries(map)) {
    const val = params.get(key);
    if (val) {
      const el = document.getElementById(inputId);
      el.value = val;
      el.classList.add('filled');
      any = true;
    }
  }
  if (any) updateStepStates();
  if (params.get('autorun') === '1') {
    setTimeout(runAgent, 600);
  }
}
