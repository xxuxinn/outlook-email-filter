// app.js — Frontend logic for the Outlook Email Agent Web UI

// ---------------------------------------------------------------------------
// Auth — the server substitutes the real token into the meta tag when it
// serves index.html. Every API call goes through apiFetch().
// ---------------------------------------------------------------------------
const AUTH_TOKEN =
  document.querySelector('meta[name="auth-token"]')?.content || '';

async function apiFetch(url, opts = {}) {
  const headers = { ...(opts.headers || {}), 'X-Auth-Token': AUTH_TOKEN };
  const r = await fetch(url, { ...opts, headers });
  if (r.status === 401) {
    showBanner('Authentication failed — reload the page to get a fresh token.');
    throw new Error('unauthorized');
  }
  return r;
}

const API = {
  async get(url) {
    const r = await apiFetch(url);
    return r.json();
  },
  async post(url, body) {
    const r = await apiFetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body ?? {}),
    });
    return r.json();
  },
};

// ---------------------------------------------------------------------------
// Warning banner
// ---------------------------------------------------------------------------
function showBanner(msg) {
  const banner = document.getElementById('warning-banner');
  document.getElementById('warning-banner-text').textContent = msg;
  banner.classList.remove('hidden');
}

function hideBanner() {
  document.getElementById('warning-banner').classList.add('hidden');
}

async function bridgeIsHealthy() {
  try {
    const h = await API.get('/api/bridge/health');
    return h.poller_responsive !== false;
  } catch (e) {
    return true; // can't tell — don't block the user
  }
}

// ---------------------------------------------------------------------------
// Settings options — keys with a small set of valid values get dropdowns
// ---------------------------------------------------------------------------
const SETTINGS_OPTIONS = {
  'General.EnableLogging':       ['True','False'],
  'General.LogLevel':            ['DEBUG','INFO','WARN','ERROR'],
  'General.EnableSelfImproving': ['True','False'],
  'General.DebugMode':           ['True','False'],
  'LLM.UseLLMAPI':               ['True','False'],
  'LLM.Provider':                ['local','azure','claude','openai'],
  'LLM.APIKeyMethod':            ['ENV','HARDCODED'],
  'Agent.EnableAutoReply':        ['True','False'],
  'Agent.AutoReplyOnArrival':     ['True','False'],
  'Agent.ScanSentItems':          ['True','False'],
  'Sync.EnableCloudSync':         ['True','False'],
};

// ---------------------------------------------------------------------------
// Tab navigation
// ---------------------------------------------------------------------------
function initTabs() {
  document.querySelectorAll('nav.tabs button').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.tab;
      document.querySelectorAll('nav.tabs button').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(`tab-${target}`).classList.add('active');
    });
  });
}

// ---------------------------------------------------------------------------
// Status header
// ---------------------------------------------------------------------------
async function loadStatus() {
  try {
    const s = await API.get('/api/status');
    const digest = s.latest_digest ? ` | digest:${s.latest_digest}` : '';
    document.getElementById('version-info').textContent =
      `v${s.version} | ${s.llm_provider} | senders:${s.learned_senders} subjects:${s.learned_subjects}${digest}`;
    const badge = document.getElementById('status-badge');
    badge.textContent = s.llm_enabled === 'True' ? 'LLM on' : 'LLM off';
    badge.className = 'status-badge ' + (s.llm_enabled === 'True' ? 'ok' : 'err');
    const bridgeBadge = document.getElementById('bridge-badge');
    bridgeBadge.textContent = s.bridge_ok ? 'bridge ok' : 'bridge stale';
    bridgeBadge.className = 'status-badge ' + (s.bridge_ok ? 'ok' : 'err');
    if (!s.bridge_ok) {
      showBanner('Outlook poller not responding — command files are sitting unconsumed. Is Outlook running?');
    }
  } catch (e) {
    document.getElementById('status-badge').textContent = 'offline';
    document.getElementById('status-badge').className = 'status-badge err';
  }
}

// ---------------------------------------------------------------------------
// Settings tab
// ---------------------------------------------------------------------------
async function loadSettings() {
  const container = document.getElementById('settings-sections');
  let settings;
  try {
    settings = await API.get('/api/settings');
  } catch (e) {
    container.innerHTML = '<div class="text-dim">Could not load settings. Is the server running?</div>';
    return;
  }
  if (!settings || Object.keys(settings).length === 0) {
    container.innerHTML = '<div class="text-dim">No settings found. Restart Outlook to auto-create settings.ini, or run ShowVersionInfo.</div>';
    return;
  }
  container.innerHTML = '';

  for (const [section, keys] of Object.entries(settings)) {
    const card = document.createElement('div');
    card.className = 'card';
    card.innerHTML = `<h3>${escHtml(section)}</h3><div class="settings-grid" id="section-${escHtml(section)}"></div>
      <div class="flex-row mt-8">
        <button class="btn success sm" onclick="saveSection('${escHtml(section)}')">Save ${escHtml(section)}</button>
      </div>`;
    container.appendChild(card);

    const grid = card.querySelector('.settings-grid');
    for (const [key, value] of Object.entries(keys)) {
      const optKey = `${section}.${key}`;
      const options = SETTINGS_OPTIONS[optKey];
      const isLong = !options && value.length > 60;

      let inputEl;
      if (options) {
        const isCustom = !options.includes(value);
        const opts = options.map(o => `<option value="${escHtml(o)}"${o === value ? ' selected' : ''}>${escHtml(o)}</option>`).join('');
        inputEl = `<div class="combo-wrap">` +
          `<select id="s_${section}_${key}" onchange="comboChange(this)">` +
          opts +
          `<option value="__custom__"${isCustom ? ' selected' : ''}>Custom…</option>` +
          `</select>` +
          `<input type="text" class="combo-custom" id="sc_${section}_${key}" value="${escHtml(value)}" style="display:${isCustom ? 'block' : 'none'}">` +
          `</div>`;
      } else if (isLong) {
        inputEl = `<textarea id="s_${section}_${key}" rows="${Math.min(4, Math.ceil(value.length/80))}">${escHtml(value)}</textarea>`;
      } else {
        inputEl = `<input type="text" id="s_${section}_${key}" value="${escHtml(value)}">`;
      }

      const field = document.createElement('div');
      field.className = 'field';
      field.style.gridColumn = isLong ? '1 / -1' : '';
      field.innerHTML = `<label>${escHtml(key)}</label>${inputEl}`;
      grid.appendChild(field);
    }
  }
}

function comboChange(sel) {
  const customInput = sel.parentElement.querySelector('.combo-custom');
  if (sel.value === '__custom__') {
    customInput.style.display = 'block';
    customInput.focus();
  } else {
    customInput.style.display = 'none';
  }
}

async function saveSection(section) {
  const values = {};
  document.querySelectorAll(`[id^="s_${section}_"]`).forEach(el => {
    // Skip companion custom inputs (sc_ prefix) — handled via the select
    if (el.classList.contains('combo-custom')) return;
    const key = el.id.replace(`s_${section}_`, '');
    if (el.tagName === 'SELECT' && el.value === '__custom__') {
      const customEl = document.getElementById(`sc_${section}_${key}`);
      values[key] = customEl ? customEl.value : '';
    } else {
      values[key] = el.value;
    }
  });
  const r = await API.post('/api/settings/section', { section, values });
  if (r.ok === false) {
    showToast(`Save failed: ${r.error || 'unknown error'}`, true);
    return;
  }
  showToast(`${section} saved`);
  loadStatus();
}

async function reloadSettings() {
  const r = await API.post('/api/settings/reload', {});
  if (r.ok === false) {
    showToast(`Reload failed: ${r.error || 'unknown error'}`, true);
    return;
  }
  showToast('Reload command sent to Outlook');
}

// ---------------------------------------------------------------------------
// Learned Rules tab
// ---------------------------------------------------------------------------
async function loadLearnedRules() {
  let senders, subjects, replies;
  try {
    [senders, subjects, replies] = await Promise.all([
      API.get('/api/learned/senders'),
      API.get('/api/learned/subjects'),
      API.get('/api/learned/replies'),
    ]);
  } catch (e) {
    document.getElementById('senders-table-body').innerHTML = '<tr><td colspan="3" class="text-dim">Could not load rules.</td></tr>';
    return;
  }

  // Senders table
  const st = document.getElementById('senders-table-body');
  st.innerHTML = senders.length === 0
    ? '<tr><td colspan="3" class="text-dim">No rules yet</td></tr>'
    : senders.map(r => `<tr>
        <td><code>${escHtml(r.email)}</code></td>
        <td><span class="badge ${r.action === 'KEEP' ? 'keep' : 'delete'}">${escHtml(r.action)}</span></td>
        <td class="text-dim">${escHtml(r.timestamp)}</td>
      </tr>`).join('');

  // Subjects table
  const subt = document.getElementById('subjects-table-body');
  subt.innerHTML = subjects.length === 0
    ? '<tr><td colspan="3" class="text-dim">No rules yet</td></tr>'
    : subjects.map(r => `<tr>
        <td><code>${escHtml(r.subject)}</code></td>
        <td><span class="badge delete">${escHtml(r.action)}</span></td>
        <td class="text-dim">${escHtml(r.timestamp)}</td>
      </tr>`).join('');

  // Replies table
  const rt = document.getElementById('replies-table-body');
  rt.innerHTML = replies.length === 0
    ? '<tr><td colspan="4" class="text-dim">No reply examples yet</td></tr>'
    : replies.map(r => `<tr>
        <td>${escHtml(r.subject)}</td>
        <td>${escHtml(r.from)}</td>
        <td class="text-dim">${escHtml(r.reply_snippet.substring(0, 80))}…</td>
        <td class="text-dim">${escHtml(r.timestamp)}</td>
      </tr>`).join('');
}

// ---------------------------------------------------------------------------
// Macros tab — the runnable list comes from the server manifest
// (GET /api/macros). Local-only hints stay client-side because these two
// subs live in ThisOutlookSession and cannot be run through the bridge.
// ---------------------------------------------------------------------------
const LOCAL_ONLY_MACROS = [
  { name: 'EnableRealTimeFilter',
    description: 'Turn on automatic filtering (run in Immediate Window: ThisOutlookSession.EnableRealTimeFilter)',
    category: 'Migration & System' },
  { name: 'DisableRealTimeFilter',
    description: 'Turn off automatic filtering (run in Immediate Window: ThisOutlookSession.DisableRealTimeFilter)',
    category: 'Migration & System' },
];

const ARG_DEFAULTS = { days: '7', pattern: '' };

let MACRO_LIST = [];

async function initMacros() {
  const grid = document.getElementById('macro-grid');
  try {
    const r = await API.get('/api/macros');
    MACRO_LIST = r.macros || [];
  } catch (e) {
    grid.innerHTML = '<span class="text-dim">Could not load macro list from the server.</span>';
    return;
  }

  let html = '';
  let currentCat = '';
  const renderCat = (cat) => {
    if (cat !== currentCat) {
      currentCat = cat;
      html += `<h4 class="macro-category">${escHtml(currentCat)}</h4>`;
    }
  };
  for (const m of MACRO_LIST) {
    renderCat(m.category);
    const destructive = m.destructive ? ' macro-destructive' : '';
    html += `<button class="macro-btn${destructive}" onclick="runMacro('${escHtml(m.name)}')">
      <div class="macro-name">${escHtml(m.name)}${m.destructive ? ' ⚠' : ''}</div>
      <div class="macro-desc">${escHtml(m.description)}</div>
    </button>`;
    if (m.name === 'ReinitializeFilter') {
      for (const lm of LOCAL_ONLY_MACROS) {
        html += `<button class="macro-btn macro-local" onclick="showLocalOnlyInfo('${escHtml(lm.name)}')">
          <div class="macro-name">${escHtml(lm.name)} ⓘ</div>
          <div class="macro-desc">${escHtml(lm.description)}</div>
        </button>`;
      }
    }
  }
  grid.innerHTML = html;
}

async function runMacro(macroName) {
  const macro = MACRO_LIST.find(m => m.name === macroName);
  if (!macro) {
    setMacroOutput(`✗ Unknown macro: ${macroName}`, 'err');
    return;
  }

  if (macro.destructive &&
      !window.confirm(`"${macroName}" moves or deletes emails / rewrites rules.\n\nRun it?`)) {
    return;
  }

  const args = {};
  for (const spec of macro.args || []) {
    const promptText = `${spec.name} (${spec.type}${spec.required ? ', required' : ''}):`;
    const val = window.prompt(promptText, ARG_DEFAULTS[spec.name] ?? '');
    if (val === null) return; // cancelled
    args[spec.name] = val;
  }

  // Fast-fail when the Outlook poller is not consuming command files.
  if (!(await bridgeIsHealthy())) {
    showBanner('Outlook poller not responding — command not sent. Is Outlook running with the VBA project loaded?');
    setMacroOutput('⚠ Outlook poller not responding (stale command files detected).\nCommand was NOT sent. Start Outlook (or run ThisOutlookSession.ReinitializeFilter) and try again.', 'warn');
    return;
  }

  setMacroOutput(`Sending command: ${macroName}...`, 'dim');
  try {
    const r = await API.post('/api/command', { macro: macroName, args });
    if (r.error) {
      setMacroOutput(`✗ Rejected: ${r.error}`, 'err');
      return;
    }
    const cmdId = r.command_id;
    setMacroOutput(`Command sent (ID: ${cmdId})\nWaiting for Outlook response...`, 'dim');
    pollResult(cmdId, (result) => {
      if (result.status === 'ok') {
        setMacroOutput(`✓ ${result.output || 'Done'}`, 'ok');
      } else if (result.status === 'timeout') {
        setMacroOutput(`⚠ Timeout — Outlook may not be running or command poller not active.\n${result.output}`, 'warn');
      } else {
        setMacroOutput(`✗ Error: ${result.output || result.status}`, 'err');
      }
    });
  } catch (e) {
    setMacroOutput(`✗ Error: ${e.message}`, 'err');
  }
}

function showLocalOnlyInfo(macroName) {
  setMacroOutput(
    `ⓘ ${macroName} must be run directly in Outlook.\n\n` +
    `Open the VBA Immediate Window (Ctrl+G in the VBA Editor) and type:\n\n` +
    `    ThisOutlookSession.${macroName}\n\n` +
    `Then press Enter.`,
    'warn'
  );
}

function setMacroOutput(text, type) {
  const el = document.getElementById('macro-output');
  el.textContent = text;
  el.className = 'output-panel';
  if (type === 'ok') el.classList.add('line-ok');
  else if (type === 'warn') el.classList.add('line-warn');
  else if (type === 'err') el.classList.add('line-err');
}

function pollResult(cmdId, callback, attempts = 0) {
  if (attempts > 240) { // 120s max (LLM calls can take time)
    // On timeout, fetch debug info about the commands directory
    API.get('/api/command/debug').then(debug => {
      const fileList = (debug.files || []).map(f => `  ${f.name} (${f.size}b)`).join('\n') || '  (empty)';
      callback({
        status: 'timeout',
        output: `Timed out waiting for Outlook (120s). Command ID: ${cmdId}\n\nCommands directory: ${debug.commands_dir}\nFiles:\n${fileList}\n\nIs Outlook running with the command poller active?`
      });
    }).catch(() => {
      callback({ status: 'timeout', output: `Timed out waiting for Outlook (120s). Command ID: ${cmdId}` });
    });
    return;
  }
  // Show elapsed time every 10 seconds
  if (attempts > 0 && attempts % 20 === 0) {
    const el = document.getElementById('macro-output');
    if (el) el.textContent = `Waiting for Outlook response... (${Math.round(attempts / 2)}s elapsed, command ID: ${cmdId})`;
  }
  setTimeout(async () => {
    try {
      const r = await API.get(`/api/command/${cmdId}/result`);
      if (r.status === 'pending') {
        pollResult(cmdId, callback, attempts + 1);
      } else {
        callback(r);
      }
    } catch (e) {
      callback({ status: 'error', output: e.message });
    }
  }, 500);
}

// ---------------------------------------------------------------------------
// Digest tab
// ---------------------------------------------------------------------------
async function loadDigest() {
  const content = document.getElementById('digest-content');
  const title = document.getElementById('digest-title');
  try {
    const r = await API.get('/api/digest');
    if (r.ok === false || !r.content) {
      title.textContent = 'Daily Digest';
      content.innerHTML = '<div class="text-dim">No digest yet. Click "Generate Digest Now" (Outlook must be running).</div>';
      return;
    }
    title.textContent = `Daily Digest — ${r.date}`;
    content.innerHTML = markdownToHtml(r.content);
  } catch (e) {
    content.innerHTML = '<div class="text-dim">Could not load digest.</div>';
  }
}

async function generateDigest() {
  const status = document.getElementById('digest-status');
  if (!(await bridgeIsHealthy())) {
    showBanner('Outlook poller not responding — digest command not sent.');
    status.textContent = 'Poller not responding.';
    return;
  }
  status.textContent = 'Generating…';
  try {
    const r = await API.post('/api/digest/generate', {});
    if (r.ok === false || r.error) {
      status.textContent = `Failed: ${r.error || 'unknown error'}`;
      return;
    }
    pollResult(r.command_id, (result) => {
      if (result.status === 'ok') {
        status.textContent = result.output || 'Digest generated.';
        loadDigest();
        loadStatus();
      } else {
        status.textContent = `${result.status}: ${result.output || ''}`;
      }
    });
  } catch (e) {
    status.textContent = `Failed: ${e.message}`;
  }
}

// ---------------------------------------------------------------------------
// Proposals tab
// ---------------------------------------------------------------------------
async function loadProposals() {
  const body = document.getElementById('proposals-table-body');
  let proposals;
  try {
    proposals = await API.get('/api/proposals');
  } catch (e) {
    body.innerHTML = '<tr><td colspan="7" class="text-dim">Could not load proposals.</td></tr>';
    return;
  }
  if (!proposals || proposals.length === 0) {
    body.innerHTML = '<tr><td colspan="7" class="text-dim">No proposals yet. Click "Generate Proposals" (requires LLM + decision history).</td></tr>';
    return;
  }
  body.innerHTML = proposals.slice().reverse().map(p => {
    const statusCls = p.status === 'PENDING' ? 'warn' : (p.status === 'APPROVED' ? 'keep' : 'delete');
    const buttons = p.status === 'PENDING'
      ? `<button class="btn success sm" onclick="approveProposal('${escHtml(p.id)}')">Approve</button>
         <button class="btn sm" onclick="rejectProposal('${escHtml(p.id)}')">Reject</button>`
      : '';
    return `<tr>
      <td>${escHtml(p.type)}</td>
      <td><code>${escHtml(p.value)}</code></td>
      <td><span class="badge ${p.action === 'KEEP' ? 'keep' : 'delete'}">${escHtml(p.action)}</span></td>
      <td class="text-dim">${escHtml(p.reason)}</td>
      <td><span class="badge ${statusCls}">${escHtml(p.status)}</span></td>
      <td class="text-dim">${escHtml(p.timestamp)}</td>
      <td class="proposal-actions">${buttons}</td>
    </tr>`;
  }).join('');
}

async function approveProposal(id) {
  try {
    const r = await API.post(`/api/proposals/${id}/approve`, {});
    if (r.ok === false) {
      showToast(`Approve failed: ${r.error || 'unknown error'}`, true);
    } else {
      const note = r.reload_error ? ` (reload not sent: ${r.reload_error})` : ` — ${r.reload_macro} sent to Outlook`;
      showToast(`Rule approved${note}`);
    }
  } catch (e) {
    showToast(`Approve failed: ${e.message}`, true);
  }
  loadProposals();
  loadStatus();
}

async function rejectProposal(id) {
  try {
    const r = await API.post(`/api/proposals/${id}/reject`, {});
    if (r.ok === false) {
      showToast(`Reject failed: ${r.error || 'unknown error'}`, true);
    } else {
      showToast('Proposal rejected');
    }
  } catch (e) {
    showToast(`Reject failed: ${e.message}`, true);
  }
  loadProposals();
}

async function generateProposals() {
  const status = document.getElementById('proposals-status');
  if (!(await bridgeIsHealthy())) {
    showBanner('Outlook poller not responding — proposals command not sent.');
    status.textContent = 'Poller not responding.';
    return;
  }
  status.textContent = 'Mining decision log…';
  try {
    const r = await API.post('/api/proposals/generate', {});
    if (r.ok === false || r.error) {
      status.textContent = `Failed: ${r.error || 'unknown error'}`;
      return;
    }
    pollResult(r.command_id, (result) => {
      if (result.status === 'ok') {
        status.textContent = result.output || 'Done.';
        loadProposals();
      } else {
        status.textContent = `${result.status}: ${result.output || ''}`;
      }
    });
  } catch (e) {
    status.textContent = `Failed: ${e.message}`;
  }
}

// ---------------------------------------------------------------------------
// Decisions tab
// ---------------------------------------------------------------------------
async function loadDecisions() {
  const body = document.getElementById('decisions-table-body');
  let decisions;
  try {
    decisions = await API.get('/api/decisions?n=100');
  } catch (e) {
    body.innerHTML = '<tr><td colspan="6" class="text-dim">Could not load decision log.</td></tr>';
    return;
  }
  if (!decisions || decisions.length === 0) {
    body.innerHTML = '<tr><td colspan="6" class="text-dim">No decisions logged yet.</td></tr>';
    return;
  }
  body.innerHTML = decisions.slice().reverse().map(d => `<tr>
    <td class="text-dim">${escHtml(d.timestamp)}</td>
    <td><code>${escHtml(d.sender)}</code></td>
    <td>${escHtml(d.subject)}</td>
    <td class="text-dim">${escHtml(d.source)}</td>
    <td><span class="badge ${d.action === 'KEEP' ? 'keep' : (d.action === 'DELETE' ? 'delete' : 'warn')}">${escHtml(d.action)}</span></td>
    <td class="text-dim">${escHtml(d.confidence)}</td>
  </tr>`).join('');
}

// ---------------------------------------------------------------------------
// Logs tab
// ---------------------------------------------------------------------------
async function loadLogs() {
  // Load error log and LLM debug log in parallel
  let errorLines, llmLines;
  try {
    [errorLines, llmLines] = await Promise.all([
      API.get('/api/errors?n=200'),
      API.get('/api/llm-debug-log?n=200'),
    ]);
  } catch (e) {
    document.getElementById('logs-container').innerHTML = '<div class="text-dim">Could not load logs.</div>';
    document.getElementById('llm-debug-container').innerHTML = '<div class="text-dim">Could not load logs.</div>';
    return;
  }

  // Error log
  const container = document.getElementById('logs-container');
  if (!errorLines || errorLines.length === 0) {
    container.innerHTML = '<div class="text-dim">No errors logged.</div>';
  } else {
    container.innerHTML = errorLines.reverse().map(line => {
      const cls = line.includes('|ERROR|') || line.toLowerCase().includes('error') ? 'error'
                : line.includes('|WARN|') ? 'warn' : '';
      return `<div class="log-line ${cls}">${escHtml(line)}</div>`;
    }).join('');
  }

  // LLM debug log
  const llmContainer = document.getElementById('llm-debug-container');
  if (!llmLines || llmLines.length === 0) {
    llmContainer.innerHTML = '<div class="text-dim">No LLM debug entries. Set LogLevel=DEBUG in settings to enable.</div>';
  } else {
    llmContainer.innerHTML = '<pre class="output-panel" style="max-height:400px;overflow:auto;font-size:12px">' +
      escHtml(llmLines.reverse().join('\n')) + '</pre>';
  }
}

async function clearErrorLog() {
  const r = await API.post('/api/errors/clear', {});
  if (r.ok === false) {
    showToast(`Could not clear: ${r.error || 'unknown error'}`, true);
    return;
  }
  document.getElementById('logs-container').innerHTML = '<div class="text-dim">No errors logged.</div>';
  showToast('Error log cleared');
}

async function clearLLMDebugLog() {
  const r = await API.post('/api/llm-debug-log/clear', {});
  if (r.ok === false) {
    showToast(`Could not clear: ${r.error || 'unknown error'}`, true);
    return;
  }
  document.getElementById('llm-debug-container').innerHTML = '<div class="text-dim">No LLM debug entries. Set LogLevel=DEBUG in settings to enable.</div>';
  showToast('LLM debug log cleared');
}

// ---------------------------------------------------------------------------
// Chat tab
// ---------------------------------------------------------------------------
function initChat() {
  const input = document.getElementById('chat-input');
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendChat();
    }
  });
  // Welcome message
  addChatBubble('agent', 'Welcome! Type a command or say **help** to see what I can do.');
}

async function sendChat() {
  const input = document.getElementById('chat-input');
  const message = input.value.trim();
  if (!message) return;
  input.value = '';

  addChatBubble('user', escHtml(message));

  let r;
  try {
    r = await API.post('/api/chat', { message });
  } catch (e) {
    addChatBubble('agent', `✗ ${escHtml(e.message)}`);
    return;
  }

  if (r.type === 'help' || r.type === 'unknown' || r.type === 'setting' || r.type === 'error') {
    addChatBubble('agent', markdownToHtml(r.output || r.label));
    return;
  }

  if (r.type === 'api') {
    addChatBubble('agent', escHtml(r.label));
    const data = await API.get(r.endpoint);
    addChatBubble('agent', `<pre>${escHtml(JSON.stringify(data, null, 2).substring(0, 2000))}</pre>`);
    return;
  }

  if (r.type === 'macro') {
    const pending = addChatBubble('agent', `<span class="spinner">↻</span> ${escHtml(r.label)}`);
    pollResult(r.command_id, (result) => {
      if (result.status === 'ok') {
        pending.innerHTML = `✓ ${escHtml(result.output || 'Done')}`;
      } else if (result.status === 'timeout') {
        pending.innerHTML = `⚠ Timeout — Outlook may not be running or command poller not active.`;
      } else {
        pending.innerHTML = `✗ ${escHtml(result.output || result.status)}`;
      }
    });
    return;
  }

  addChatBubble('agent', escHtml(r.output || r.label || 'Done.'));
}

function addChatBubble(role, html) {
  const messages = document.getElementById('chat-messages');
  const bubble = document.createElement('div');
  bubble.className = `chat-bubble ${role}`;
  bubble.innerHTML = html;
  messages.appendChild(bubble);
  messages.scrollTop = messages.scrollHeight;
  return bubble;
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------
function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

// Minimal markdown renderer. Escapes FIRST, then applies formatting on the
// escaped text — no raw HTML from the input ever reaches the DOM.
// Supports: # / ## / ### headings, "- " lists, **bold**, `code`.
function markdownToHtml(text) {
  const inline = (s) => s
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/`(.+?)`/g, '<code>$1</code>');

  const lines = escHtml(text).split('\n');
  const out = [];
  let inList = false;
  const closeList = () => {
    if (inList) { out.push('</ul>'); inList = false; }
  };

  for (const line of lines) {
    const li = line.match(/^\s*-\s+(.*)$/);
    if (li) {
      if (!inList) { out.push('<ul>'); inList = true; }
      out.push(`<li>${inline(li[1])}</li>`);
      continue;
    }
    closeList();
    const heading = line.match(/^(#{1,3})\s+(.*)$/);
    if (heading) {
      const level = heading[1].length + 1; // # → h2, ## → h3, ### → h4
      out.push(`<h${level} class="md-h">${inline(heading[2])}</h${level}>`);
      continue;
    }
    out.push(`${inline(line)}<br>`);
  }
  closeList();
  return out.join('');
}

function showToast(msg, isError = false) {
  const t = document.createElement('div');
  const color = isError ? '#ef4444' : '#64ffda';
  const bg = isError ? '#3a1a1a' : '#1a3a2e';
  t.style.cssText = `position:fixed;bottom:20px;right:20px;background:${bg};color:${color};border:1px solid ${color};padding:10px 16px;border-radius:6px;font-size:13px;z-index:9999`;
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), isError ? 5000 : 2500);
}

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initMacros();
  initChat();
  loadStatus();

  // Load the active tab's content on tab switch
  document.querySelectorAll('nav.tabs button').forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      if (tab === 'settings') loadSettings();
      else if (tab === 'rules') loadLearnedRules();
      else if (tab === 'digest') loadDigest();
      else if (tab === 'proposals') loadProposals();
      else if (tab === 'decisions') loadDecisions();
      else if (tab === 'logs') loadLogs();
    });
  });

  // Load default tab (settings)
  document.querySelector('[data-tab="settings"]')?.click();
});
