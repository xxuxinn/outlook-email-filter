// app.js — Frontend logic for the Outlook Email Agent Web UI

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
};

const API = {
  async get(url) {
    const r = await fetch(url);
    return r.json();
  },
  async post(url, body) {
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    return r.json();
  },
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
    document.getElementById('version-info').textContent =
      `v${s.version} | ${s.llm_provider} | senders:${s.learned_senders} subjects:${s.learned_subjects}`;
    const badge = document.getElementById('status-badge');
    badge.textContent = s.llm_enabled === 'True' ? 'LLM on' : 'LLM off';
    badge.className = 'status-badge ' + (s.llm_enabled === 'True' ? 'ok' : 'err');
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
    card.innerHTML = `<h3>${section}</h3><div class="settings-grid" id="section-${section}"></div>
      <div class="flex-row mt-8">
        <button class="btn success sm" onclick="saveSection('${section}')">Save ${section}</button>
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
      field.innerHTML = `<label>${key}</label>${inputEl}`;
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
  await API.post('/api/settings/section', { section, values });
  showToast(`${section} saved`);
  loadStatus();
}

async function reloadSettings() {
  await API.post('/api/settings/reload', {});
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
        <td><span class="badge ${r.action === 'KEEP' ? 'keep' : 'delete'}">${r.action}</span></td>
        <td class="text-dim">${r.timestamp}</td>
      </tr>`).join('');

  // Subjects table
  const subt = document.getElementById('subjects-table-body');
  subt.innerHTML = subjects.length === 0
    ? '<tr><td colspan="3" class="text-dim">No rules yet</td></tr>'
    : subjects.map(r => `<tr>
        <td><code>${escHtml(r.subject)}</code></td>
        <td><span class="badge delete">${r.action}</span></td>
        <td class="text-dim">${r.timestamp}</td>
      </tr>`).join('');

  // Replies table
  const rt = document.getElementById('replies-table-body');
  rt.innerHTML = replies.length === 0
    ? '<tr><td colspan="4" class="text-dim">No reply examples yet</td></tr>'
    : replies.map(r => `<tr>
        <td>${escHtml(r.subject)}</td>
        <td>${escHtml(r.from)}</td>
        <td class="text-dim">${escHtml(r.reply_snippet.substring(0, 80))}…</td>
        <td class="text-dim">${r.timestamp}</td>
      </tr>`).join('');
}

// ---------------------------------------------------------------------------
// Macros tab
// ---------------------------------------------------------------------------
const MACROS = [
  // Version
  { name: 'ShowVersionInfo', desc: 'Show version and status', cat: 'Version' },
  // Filtering
  { name: 'FilterExistingDryRun', desc: 'Preview filter decisions (no changes)', cat: 'Filtering' },
  { name: 'FilterExistingEmails', desc: 'Filter all Inbox emails', cat: 'Filtering' },
  { name: 'FilterAllFolders', desc: 'Filter Inbox + Other + PST archives', cat: 'Filtering' },
  { name: 'FilterSelectedEmail', desc: 'Test classification on one selected email', cat: 'Filtering' },
  { name: 'FilterSelectedEmails', desc: 'Filter selected email(s) with confirmation', cat: 'Filtering' },
  { name: 'FilterCurrentFolder', desc: 'Filter current folder with confirmation', cat: 'Filtering' },
  { name: 'FilterLastNDays', desc: 'Filter last N days', cat: 'Filtering',
    param: { key: 'days', prompt: 'Number of days to filter (e.g. 7):', default: '7' } },
  { name: 'GenerateClassificationReport', desc: 'Count classifications without acting', cat: 'Filtering' },
  { name: 'BulkDeleteBySender', desc: 'Delete all from matching senders', cat: 'Filtering',
    param: { key: 'pattern', prompt: 'Sender pattern to match (e.g. "noreply"):', default: '' } },
  { name: 'MoveProtectedSources', desc: 'Move protected domain emails to Protected folder', cat: 'Filtering' },
  // Agent Tools
  { name: 'GenerateAddressingPatterns', desc: 'LLM-generate name/greeting patterns', cat: 'Agent Tools' },
  { name: 'ScanSentForReplyPatterns', desc: 'Learn reply style from Sent Items', cat: 'Agent Tools' },
  { name: 'DraftRepliesForInbox', desc: 'Batch draft replies for unread KEEP emails', cat: 'Agent Tools' },
  { name: 'ShowLearnedRepliesSummary', desc: 'Show learned reply pair count', cat: 'Agent Tools' },
  // LLM Tools
  { name: 'SummarizeSelectedEmail', desc: 'Summarize selected email using LLM', cat: 'LLM Tools' },
  { name: 'DraftReplyToSelected', desc: 'Draft reply to selected email using LLM', cat: 'LLM Tools' },
  // Learned Rules
  { name: 'ShowLearnedSenders', desc: 'Show learned sender rule count', cat: 'Learned Rules' },
  { name: 'ShowLearnedSendersList', desc: 'Dump sender rules to Immediate Window', cat: 'Learned Rules' },
  { name: 'ReloadLearnedSenders', desc: 'Force reload learned rules from file', cat: 'Learned Rules' },
  { name: 'CleanLearnedSendersFile', desc: 'Remove duplicate sender entries', cat: 'Learned Rules' },
  { name: 'ImportExistingLearnedFolders', desc: 'Bulk import from LearnKeep/LearnDelete', cat: 'Learned Rules' },
  { name: 'ShowLearnedSubjectsList', desc: 'Dump subject rules to Immediate Window', cat: 'Learned Rules' },
  { name: 'CleanLearnedSubjectsFile', desc: 'Remove duplicate subject entries', cat: 'Learned Rules' },
  { name: 'ImportExistingLearnedSubjectFolder', desc: 'Bulk import from LearnSubjectDelete', cat: 'Learned Rules' },
  // Server Rules
  { name: 'ImportServerRules', desc: 'Import server rules as learned DELETE rules', cat: 'Server Rules' },
  { name: 'ExportLearnedRulesToServer', desc: 'Push DELETE rules to Exchange server', cat: 'Server Rules' },
  // Undo / Recovery
  { name: 'RestoreFromReview', desc: 'Move Review folder emails back to Inbox', cat: 'Undo / Recovery' },
  { name: 'RestoreDeletedKeepEmails', desc: 'Rescue wrongly deleted KEEP emails', cat: 'Undo / Recovery' },
  // Migration & System
  { name: 'DetectAndMigrateOldFolders', desc: 'Rename v1.x folders to v2.0 names', cat: 'Migration & System' },
  { name: 'ReinitializeFilter', desc: 'Restart event handlers', cat: 'Migration & System' },
  { name: 'EnableRealTimeFilter', desc: 'Turn on automatic filtering (run in Immediate Window: ThisOutlookSession.EnableRealTimeFilter)', cat: 'Migration & System', localOnly: true },
  { name: 'DisableRealTimeFilter', desc: 'Turn off automatic filtering (run in Immediate Window: ThisOutlookSession.DisableRealTimeFilter)', cat: 'Migration & System', localOnly: true },
];

function initMacros() {
  const grid = document.getElementById('macro-grid');
  let html = '';
  let currentCat = '';
  for (const m of MACROS) {
    if (m.cat !== currentCat) {
      currentCat = m.cat;
      html += `<h4 class="macro-category">${escHtml(currentCat)}</h4>`;
    }
    if (m.localOnly) {
      html += `<button class="macro-btn macro-local" onclick="showLocalOnlyInfo('${m.name}')">
        <div class="macro-name">${m.name} ⓘ</div>
        <div class="macro-desc">${m.desc}</div>
      </button>`;
    } else {
      html += `<button class="macro-btn" onclick="runMacro('${m.name}')">
        <div class="macro-name">${m.name}</div>
        <div class="macro-desc">${m.desc}</div>
      </button>`;
    }
  }
  grid.innerHTML = html;
}

async function runMacro(macroName) {
  // Check if macro needs a parameter
  const macro = MACROS.find(m => m.name === macroName);
  let args = {};
  if (macro && macro.param) {
    const val = window.prompt(macro.param.prompt, macro.param.default);
    if (val === null) return; // cancelled
    args[macro.param.key] = val;
  }

  setMacroOutput(`Sending command: ${macroName}...`, 'dim');
  try {
    const r = await API.post('/api/command', { macro: macroName, args });
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
    const r = await API.get(`/api/command/${cmdId}/result`);
    if (r.status === 'pending') {
      pollResult(cmdId, callback, attempts + 1);
    } else {
      callback(r);
    }
  }, 500);
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
  await API.post('/api/errors/clear', {});
  document.getElementById('logs-container').innerHTML = '<div class="text-dim">No errors logged.</div>';
  showToast('Error log cleared');
}

async function clearLLMDebugLog() {
  await API.post('/api/llm-debug-log/clear', {});
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

  const r = await API.post('/api/chat', { message });

  if (r.type === 'help' || r.type === 'unknown' || r.type === 'setting') {
    addChatBubble('agent', markdownToHtml(r.output || r.label));
    return;
  }

  if (r.type === 'api') {
    addChatBubble('agent', r.label);
    const data = await API.get(r.endpoint);
    addChatBubble('agent', `<pre>${escHtml(JSON.stringify(data, null, 2).substring(0, 2000))}</pre>`);
    return;
  }

  if (r.type === 'macro') {
    const pending = addChatBubble('agent', `<span class="spinner">↻</span> ${r.label}`);
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

  addChatBubble('agent', r.output || r.label || 'Done.');
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

function escAttr(str) {
  return String(str ?? '').replace(/'/g, "\\'").replace(/"/g, '\\"');
}

function markdownToHtml(text) {
  // Minimal: bold **text**, code `text`, newlines
  return escHtml(text)
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/`(.+?)`/g, '<code>$1</code>')
    .replace(/\n/g, '<br>');
}

function showToast(msg) {
  const t = document.createElement('div');
  t.style.cssText = 'position:fixed;bottom:20px;right:20px;background:#1a3a2e;color:#64ffda;border:1px solid #64ffda;padding:10px 16px;border-radius:6px;font-size:13px;z-index:9999';
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 2500);
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
      else if (tab === 'logs') loadLogs();
    });
  });

  // Load default tab (settings)
  document.querySelector('[data-tab="settings"]')?.click();
});
