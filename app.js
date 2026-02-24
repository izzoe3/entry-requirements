/* ═══════════════════════════════════════════════════════════════
   Entry Requirements Manager — app.js

   Screens:    setup → login → denied | app
   Steps:      1:select  2:view  3:edit  4:output  5:changelog
   Modal:      inline-edit for single qualification quick-save
   Changelog:  Sheets tab "Changelog" — written on every save,
               displayed per-programme (view) and globally (step 5)
═══════════════════════════════════════════════════════════════ */

const App = (() => {

  // ── CONFIG ─────────────────────────────────────────────────────
  const CFG = {
    get clientId()  { return localStorage.getItem('cfg_client_id')  || ''; },
    get sheetId()   { return localStorage.getItem('cfg_sheet_id')   || ''; },
    set clientId(v) { localStorage.setItem('cfg_client_id',  v); },
    set sheetId(v)  { localStorage.setItem('cfg_sheet_id',   v); },
  };

  // ── CONSTANTS ──────────────────────────────────────────────────
  const SHEET_NAME   = 'Entry Requirements';
  const ACCESS_TAB   = 'Access';
  const CHANGELOG_TAB = 'Changelog';
  // Changelog columns (0-based):
  // 0:Timestamp  1:UserEmail  2:UserName  3:Programme  4:Qualification  5:OldValue  6:NewValue
  const CL_COLS = ['Timestamp','User Email','User Name','Programme','Qualification','Old Value','New Value'];

  const STEPS = ['step-select', 'step-view', 'step-edit', 'step-output', 'step-changelog'];

  // ── STATE ──────────────────────────────────────────────────────
  let accessToken    = null;
  let userEmail      = null;
  let userName       = null;
  let userPicture    = null;
  let allHeaders     = [];
  let qualHeaders    = [];
  let allData        = [];
  let changelog      = [];   // all changelog rows as objects, newest first
  let lastUpdatedMap = {};   // progName → { time, user } for card display
  let currentProg    = null;
  let currentRowIdx  = null;
  let inlineQualKey  = null;
  let _toastTimer    = null;

  // ── INIT ───────────────────────────────────────────────────────
  function init() {
    if (!CFG.clientId || !CFG.sheetId) {
      showScreen('screen-setup');
    } else {
      showScreen('screen-login');
    }
  }

  // ── SCREEN MANAGEMENT ──────────────────────────────────────────
  function showScreen(id) {
    document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
    const el = document.getElementById(id);
    if (el) el.classList.add('active');
    window.scrollTo(0, 0);
  }

  // ── STEP MANAGEMENT ────────────────────────────────────────────
  function showStep(n) {
    document.querySelectorAll('.step-panel').forEach(s => s.classList.remove('active'));
    const el = document.getElementById(STEPS[n - 1]);
    if (el) el.classList.add('active');
    window.scrollTo(0, 0);
  }

  function goToStep(n) {
    hideToast();
    if (n === 1) {
      currentProg   = null;
      currentRowIdx = null;
      renderProgList();
      filterProgList(document.getElementById('prog-search').value);
    }
    showStep(n);
  }

  function goToView() {
    renderViewTable();
    loadProgHistory();
    showStep(2);
  }

  function goToEdit() {
    renderQualGrid(allData[currentRowIdx]);
    showStep(3);
  }

  function openChangelog() {
    renderChangelog();
    showStep(5);
  }

  // ── CONFIG SAVE ────────────────────────────────────────────────
  function saveConfig() {
    const sheetId  = document.getElementById('cfg-sheet-id').value.trim();
    const clientId = document.getElementById('cfg-client-id').value.trim();
    if (!sheetId)  { alert('Please enter your Spreadsheet ID.');  document.getElementById('cfg-sheet-id').focus();  return; }
    if (!clientId) { alert('Please enter your OAuth Client ID.'); document.getElementById('cfg-client-id').focus(); return; }
    CFG.sheetId  = sheetId;
    CFG.clientId = clientId;
    showScreen('screen-login');
  }

  // ── AUTH ───────────────────────────────────────────────────────
  function signIn() {
    if (!CFG.clientId || !CFG.sheetId) { showScreen('screen-setup'); return; }
    if (!window.google?.accounts?.oauth2) {
      alert('Google sign-in library not loaded. Check your internet connection and reload.');
      return;
    }
    let tokenClient;
    try {
      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CFG.clientId,
        scope: [
          'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/userinfo.profile',
          'https://www.googleapis.com/auth/userinfo.email',
        ].join(' '),
        callback:       onTokenResponse,
        error_callback: err => alert('Sign-in error: ' + (err.message || err.type || 'Unknown') +
          '\n\nMake sure popups are allowed for this page.'),
      });
    } catch (err) {
      alert('Could not initialise sign-in: ' + err.message); return;
    }
    tokenClient.requestAccessToken({ prompt: 'select_account' });
  }

  async function onTokenResponse(resp) {
    if (resp.error) {
      alert('Sign-in failed: ' + resp.error + (resp.error_description ? '\n' + resp.error_description : ''));
      return;
    }
    accessToken = resp.access_token;
    try {
      const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: 'Bearer ' + accessToken },
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const p = await res.json();
      userEmail = p.email; userName = p.name; userPicture = p.picture;
    } catch (err) { alert('Signed in but could not fetch profile: ' + err.message); return; }
    await checkAccess();
  }

  async function checkAccess() {
    try {
      const res = await apiFetch(apiUrl(ACCESS_TAB + '!A:A'));
      if (!res.ok) { const e = await res.json().catch(()=>({})); throw new Error(e.error?.message||'HTTP '+res.status); }
      const data    = await res.json();
      const allowed = (data.values||[]).flat().map(e=>String(e).trim().toLowerCase()).filter(e=>e.includes('@'));
      if (allowed.includes(userEmail.toLowerCase())) { onAccessGranted(); } else { onAccessDenied(); }
    } catch (err) {
      alert('Could not verify access: ' + err.message +
        '\n\nMake sure:\n• The "' + ACCESS_TAB + '" tab exists\n• Your Spreadsheet ID is correct');
      showScreen('screen-login');
    }
  }

  function onAccessGranted() {
    const avatar = document.getElementById('user-avatar');
    if (userPicture) { avatar.innerHTML = `<img src="${userPicture}" alt="" />`; }
    else { avatar.textContent = (userName||userEmail).charAt(0).toUpperCase(); }
    document.getElementById('user-name-display').textContent = userName || userEmail;
    document.getElementById('topbar-user').classList.add('visible');
    document.getElementById('topbar-changelog-btn').classList.remove('hidden');
    showScreen('screen-app');
    showStep(1);
    loadAllData();
  }

  function onAccessDenied() {
    document.getElementById('denied-email').textContent = userEmail;
    showScreen('screen-denied');
  }

  function signOut() {
    accessToken = userEmail = userName = userPicture = null;
    allData = []; changelog = []; lastUpdatedMap = {};
    currentProg = null; currentRowIdx = null;
    document.getElementById('topbar-user').classList.remove('visible');
    document.getElementById('topbar-changelog-btn').classList.add('hidden');
    showScreen('screen-login');
  }

  // ── LOAD ALL DATA ──────────────────────────────────────────────
  // Load entry requirements and changelog in parallel
  async function loadAllData() {
    renderProgListSkeleton();
    try {
      const [entryRes, changeRes] = await Promise.all([
        apiFetch(apiUrl(SHEET_NAME + '!A1:ZZ')),
        apiFetch(apiUrl(CHANGELOG_TAB + '!A1:G')),
      ]);

      // Parse entry requirements
      if (!entryRes.ok) {
        const e = await entryRes.json().catch(()=>({}));
        throw new Error(e.error?.message || 'HTTP ' + entryRes.status);
      }
      const entryData = await entryRes.json();
      const rows = entryData.values || [];
      if (rows.length < 2) {
        renderProgListMsg('No programmes found. Make sure the "' + SHEET_NAME + '" tab has a header row and at least one data row.');
        return;
      }
      allHeaders  = rows[0].map(h => String(h).trim());
      qualHeaders = allHeaders.slice(1);
      allData     = rows.slice(1)
        .filter(r => String(r[0]||'').trim() !== '')
        .map(row => {
          const obj = {};
          allHeaders.forEach((h,i) => { obj[h] = String(row[i]||'').trim(); });
          return obj;
        });

      // Parse changelog (may not exist yet — that's fine)
      if (changeRes.ok) {
        const changeData = await changeRes.json();
        const clRows = changeData.values || [];
        // Skip header row, parse newest-first
        changelog = clRows.slice(1).reverse().map(row => ({
          timestamp:     row[0] || '',
          userEmail:     row[1] || '',
          userName:      row[2] || '',
          programme:     row[3] || '',
          qualification: row[4] || '',
          oldValue:      row[5] || '',
          newValue:      row[6] || '',
        }));
        buildLastUpdatedMap();
      }

      // Ensure Changelog tab has headers (create if needed, silently)
      ensureChangelogHeaders();

      renderProgList();
    } catch (err) {
      renderProgListMsg('Failed to load data: ' + err.message);
      showToast('error', '<strong>Could not load sheet data.</strong> ' + err.message);
    }
  }

  // Build a map of programme → most recent change for card display
  function buildLastUpdatedMap() {
    lastUpdatedMap = {};
    // changelog is newest-first, so first match per programme wins
    changelog.forEach(entry => {
      if (!lastUpdatedMap[entry.programme]) {
        lastUpdatedMap[entry.programme] = {
          time: entry.timestamp,
          user: entry.userName || entry.userEmail,
        };
      }
    });
  }

  // Create the Changelog tab header row if the tab is empty/missing
  async function ensureChangelogHeaders() {
    try {
      const res = await apiFetch(apiUrl(CHANGELOG_TAB + '!A1:G1'));
      if (!res.ok) return; // tab might not exist yet — first save will create it
      const data = await res.json();
      const existing = (data.values||[])[0] || [];
      if (existing.length === 0) {
        // Write headers
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${encodeURIComponent(CHANGELOG_TAB+'!A1:G1')}?valueInputOption=RAW`;
        await apiFetch(url, {
          method: 'PUT',
          body: JSON.stringify({ range: CHANGELOG_TAB+'!A1:G1', majorDimension:'ROWS', values:[CL_COLS] }),
        });
      }
    } catch(_) { /* silent — not critical */ }
  }

  // ── STEP 1: PROGRAMME LIST ─────────────────────────────────────
  function renderProgListSkeleton() {
    document.getElementById('prog-list').innerHTML =
      [1,2,3,4].map(() => '<div class="skeleton"></div>').join('');
  }

  function renderProgListMsg(msg) {
    document.getElementById('prog-list').innerHTML = `<p class="prog-empty">${msg}</p>`;
  }

  function renderProgList() {
    const progKey = allHeaders[0];
    const total   = qualHeaders.length;

    document.getElementById('prog-list').innerHTML = allData.map((prog, i) => {
      const name   = escHtml(prog[progKey] || '(Unnamed)');
      const filled = qualHeaders.filter(h => prog[h] && prog[h].trim()).length;
      const pct    = total > 0 ? Math.round((filled/total)*100) : 0;
      const lu     = lastUpdatedMap[prog[progKey]];

      const luHtml = lu
        ? `<div class="prog-last-updated">
             <svg viewBox="0 0 14 14" fill="none" width="11" height="11">
               <circle cx="7" cy="7" r="5.5" stroke="currentColor" stroke-width="1.3"/>
               <path d="M7 4v3.2l2 1.2" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/>
             </svg>
             Updated ${relativeTime(lu.time)} by ${escHtml(lu.user)}
           </div>`
        : `<div class="prog-last-updated" style="opacity:0.45;">No changes recorded yet</div>`;

      return `
        <div class="prog-card" data-name="${prog[progKey].toLowerCase()}">
          <div class="prog-icon">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
              <rect x="3" y="2" width="14" height="16" rx="2" stroke="currentColor" stroke-width="1.5"/>
              <path d="M6 7h8M6 10h8M6 13h5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
            </svg>
          </div>
          <div class="prog-info">
            <div class="prog-name">${name}</div>
            <div class="prog-meta">${filled} of ${total} qualifications · ${pct}% complete</div>
            ${luHtml}
          </div>
          <div class="prog-actions">
            <button class="prog-btn view" onclick="App.openView(${i})">
              <svg width="12" height="12" viewBox="0 0 14 14" fill="none">
                <ellipse cx="7" cy="7" rx="6" ry="4" stroke="currentColor" stroke-width="1.5"/>
                <circle cx="7" cy="7" r="1.8" fill="currentColor"/>
              </svg>
              View
            </button>
            <button class="prog-btn edit" onclick="App.openEdit(${i})">
              <svg width="12" height="12" viewBox="0 0 14 14" fill="none">
                <path d="M9.5 1.5l3 3-7.5 7.5H2.5v-3l7-7z" stroke="white" stroke-width="1.4" stroke-linejoin="round"/>
              </svg>
              Edit
            </button>
          </div>
        </div>`;
    }).join('');

    filterProgList(document.getElementById('prog-search')?.value || '');
  }

  // ── SEARCH ─────────────────────────────────────────────────────
  function filterProgList(query) {
    const q        = query.trim().toLowerCase();
    const cards    = document.querySelectorAll('#prog-list .prog-card');
    const clearBtn = document.getElementById('search-clear');
    const emptyMsg = document.getElementById('prog-empty');
    let   visible  = 0;
    clearBtn.classList.toggle('hidden', q === '');
    cards.forEach(card => {
      const matches = !q || (card.dataset.name||'').includes(q);
      card.style.display = matches ? '' : 'none';
      if (matches) visible++;
    });
    emptyMsg.classList.toggle('hidden', visible > 0 || cards.length === 0);
  }

  function clearSearch() {
    const input = document.getElementById('prog-search');
    input.value = '';
    input.focus();
    filterProgList('');
  }

  // ── OPEN VIEW / EDIT ───────────────────────────────────────────
  function openView(idx) {
    setCurrentProg(idx);
    renderViewTable();
    loadProgHistory();
    showStep(2);
  }

  function openEdit(idx) {
    setCurrentProg(idx);
    renderQualGrid(allData[idx]);
    showStep(3);
  }

  function setCurrentProg(idx) {
    currentProg   = allData[idx][allHeaders[0]];
    currentRowIdx = idx;
  }

  // ── STEP 2: VIEW TABLE ─────────────────────────────────────────
  function renderViewTable() {
    const prog   = allData[currentRowIdx];
    const filled = qualHeaders.filter(h => prog[h] && prog[h].trim());

    document.getElementById('view-prog-name').textContent = currentProg;
    document.getElementById('view-meta').textContent =
      filled.length + ' of ' + qualHeaders.length + ' qualifications apply to this programme';

    const wrap = document.getElementById('view-table-wrap');

    if (filled.length === 0) {
      wrap.innerHTML = `
        <div class="view-empty">
          No entry requirements filled in yet.<br>
          <button class="btn btn-primary btn-sm" style="margin-top:16px;" onclick="App.goToEdit()">Fill in requirements</button>
        </div>`;
      return;
    }

    const rows = filled.map(h => `
      <tr>
        <td>
          ${escHtml(h)}
          <button class="inline-edit-btn" onclick="App.openInlineEdit('${h.replace(/'/g,"\\'")}')">
            <svg width="10" height="10" viewBox="0 0 12 12" fill="none">
              <path d="M8 1.5l2.5 2.5-6 6H2v-2.5l6-6z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/>
            </svg>
            Edit
          </button>
        </td>
        <td>${escHtml(prog[h])}</td>
      </tr>`).join('');

    wrap.innerHTML = `
      <table class="view-table">
        <thead><tr><th>Qualification</th><th>Requirements</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>`;

    document.getElementById('view-copy-confirm').classList.remove('visible');
  }

  // ── STEP 2: PER-PROGRAMME HISTORY ─────────────────────────────
  function loadProgHistory() {
    const wrap    = document.getElementById('view-history-wrap');
    const metaEl  = document.getElementById('view-history-meta');
    wrap.innerHTML = '<div class="history-loading">Loading…</div>';

    // Filter changelog for this programme (already in memory)
    const entries = changelog.filter(e => e.programme === currentProg);
    metaEl.textContent = entries.length + ' change' + (entries.length !== 1 ? 's' : '');

    if (entries.length === 0) {
      wrap.innerHTML = '<div class="history-empty">No changes recorded yet for this programme.</div>';
      return;
    }

    wrap.innerHTML = `<div class="history-list">${entries.map(renderHistoryItem).join('')}</div>`;
  }

  function renderHistoryItem(entry) {
    const initials = (entry.userName||entry.userEmail||'?').charAt(0).toUpperCase();
    const oldVal   = entry.oldValue.trim();
    const newVal   = entry.newValue.trim();

    let diffHtml;
    if (!oldVal && newVal) {
      // Added
      diffHtml = `<div class="diff-row">
        <span class="diff-label new">New</span>
        <span class="diff-value new">${escHtml(newVal)}</span>
      </div>`;
    } else if (oldVal && !newVal) {
      // Removed
      diffHtml = `<div class="diff-row">
        <span class="diff-label old">Was</span>
        <span class="diff-value old">${escHtml(oldVal)}</span>
      </div>
      <div class="diff-row">
        <span class="diff-label new">Now</span>
        <span class="diff-value empty">(removed)</span>
      </div>`;
    } else {
      // Changed
      diffHtml = `<div class="diff-row">
        <span class="diff-label old">Was</span>
        <span class="diff-value old">${escHtml(oldVal)}</span>
      </div>
      <div class="diff-row">
        <span class="diff-label new">Now</span>
        <span class="diff-value new">${escHtml(newVal)}</span>
      </div>`;
    }

    return `
      <div class="history-item">
        <div class="history-meta">
          <span class="history-time">${formatTimestamp(entry.timestamp)}</span>
          <div class="history-user">
            <div class="history-user-avatar">${initials}</div>
            ${escHtml(entry.userName || entry.userEmail)}
          </div>
        </div>
        <div class="history-body">
          <div class="history-qual">${escHtml(entry.qualification)}</div>
          <div class="history-diff">${diffHtml}</div>
        </div>
      </div>`;
  }

  // ── STEP 3: QUAL GRID ──────────────────────────────────────────
  function renderQualGrid(prog) {
    document.getElementById('edit-prog-name').textContent = currentProg;
    document.getElementById('qual-grid').innerHTML = qualHeaders.map((header, i) => {
      const val    = prog[header] || '';
      const filled = val.trim() !== '';
      return `
        <div class="qual-item ${filled?'has-value':''}" id="qi-${i}">
          <div class="qual-header" onclick="document.getElementById('qa-${i}').focus()">
            <div class="qual-name">
              ${escHtml(header)}
              <span class="qual-badge ${filled?'filled':'empty'}" id="qb-${i}">${filled?'Filled':'Empty'}</span>
            </div>
          </div>
          <div class="qual-body">
            <textarea id="qa-${i}"
              placeholder="Leave blank if not accepted for this programme…"
              oninput="App.onQualInput(${i}, this)"
            >${escHtml(val)}</textarea>
          </div>
        </div>`;
    }).join('');
    updateFilledCount();
  }

  function onQualInput(i, el) {
    const filled = el.value.trim() !== '';
    document.getElementById('qi-'+i).classList.toggle('has-value', filled);
    const badge = document.getElementById('qb-'+i);
    badge.className   = 'qual-badge ' + (filled ? 'filled' : 'empty');
    badge.textContent = filled ? 'Filled' : 'Empty';
    updateFilledCount();
  }

  function updateFilledCount() {
    const filled = document.querySelectorAll('#qual-grid .qual-item.has-value').length;
    document.getElementById('filled-count').textContent = filled + ' of ' + qualHeaders.length + ' filled';
  }

  // ── SAVE FULL FORM ─────────────────────────────────────────────
  async function saveRequirements() {
    // Snapshot old values for diff
    const oldValues = {};
    qualHeaders.forEach(h => { oldValues[h] = allData[currentRowIdx][h] || ''; });

    // Collect new values
    const fields = {};
    qualHeaders.forEach((h, i) => { fields[h] = (document.getElementById('qa-'+i)?.value||'').trim(); });

    const newRow   = allHeaders.map((h,i) => i===0 ? currentProg : (fields[h]||''));
    const sheetRow = currentRowIdx + 2;
    const range    = `${SHEET_NAME}!A${sheetRow}:${colLetter(allHeaders.length-1)}${sheetRow}`;
    const url      = `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;

    const btn = document.getElementById('save-btn');
    btn.disabled = true;
    btn.innerHTML = '<div class="spinner"></div> Saving…';

    try {
      const res = await apiFetch(url, {
        method: 'PUT',
        body:   JSON.stringify({ range, majorDimension:'ROWS', values:[newRow] }),
      });
      if (!res.ok) { const e=await res.json().catch(()=>({})); throw new Error(e.error?.message||'HTTP '+res.status); }

      // Update local cache
      allHeaders.forEach((h,i) => { allData[currentRowIdx][h] = newRow[i]; });

      // Compute diff and write changelog entries
      const changes = qualHeaders
        .filter(h => (oldValues[h]||'') !== (fields[h]||''))
        .map(h => ({ qualification: h, oldValue: oldValues[h]||'', newValue: fields[h]||'' }));

      if (changes.length > 0) {
        await writeChangelogEntries(changes);
      }

      document.getElementById('output-prog-name').textContent = currentProg;
      buildOutputTable(fields);
      showStep(4);
      showToast('success', `"${currentProg}" saved. ${changes.length} field${changes.length!==1?'s':''} changed.`);

    } catch (err) {
      showToast('error', '<strong>Save failed:</strong> ' + err.message);
    } finally {
      btn.disabled = false;
      btn.innerHTML = `
        <svg width="13" height="13" viewBox="0 0 13 13" fill="none">
          <path d="M2 7l3 3 6-6" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        Save &amp; Generate Table`;
    }
  }

  // ── INLINE EDIT MODAL ──────────────────────────────────────────
  function openInlineEdit(qualKey) {
    inlineQualKey = qualKey;
    document.getElementById('modal-prog-label').textContent = currentProg;
    document.getElementById('modal-qual-name').textContent  = qualKey;
    document.getElementById('modal-textarea').value         = allData[currentRowIdx][qualKey] || '';
    document.getElementById('modal-overlay').classList.add('open');
    document.getElementById('inline-modal').classList.add('open');
    setTimeout(() => document.getElementById('modal-textarea').focus(), 100);
  }

  function closeInlineEdit() {
    document.getElementById('modal-overlay').classList.remove('open');
    document.getElementById('inline-modal').classList.remove('open');
    inlineQualKey = null;
  }

  async function saveInlineEdit() {
    if (!inlineQualKey) return;

    const oldVal   = allData[currentRowIdx][inlineQualKey] || '';
    const newVal   = document.getElementById('modal-textarea').value.trim();
    const sheetRow = currentRowIdx + 2;
    const colIdx   = allHeaders.indexOf(inlineQualKey);
    const col      = colLetter(colIdx);
    const range    = `${SHEET_NAME}!${col}${sheetRow}`;
    const url      = `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;

    const saveBtn = document.getElementById('modal-save-btn');
    saveBtn.disabled = true;
    saveBtn.innerHTML = '<div class="spinner"></div> Saving…';

    try {
      const res = await apiFetch(url, {
        method: 'PUT',
        body:   JSON.stringify({ range, majorDimension:'ROWS', values:[[newVal]] }),
      });
      if (!res.ok) { const e=await res.json().catch(()=>({})); throw new Error(e.error?.message||'HTTP '+res.status); }

      // Update local cache
      allData[currentRowIdx][inlineQualKey] = newVal;

      // Write changelog
      if (oldVal !== newVal) {
        await writeChangelogEntries([{ qualification: inlineQualKey, oldValue: oldVal, newValue: newVal }]);
      }

      closeInlineEdit();
      renderViewTable();
      loadProgHistory();
      showToast('success', `"${inlineQualKey}" updated.`);

    } catch (err) {
      showToast('error', '<strong>Save failed:</strong> ' + err.message);
    } finally {
      saveBtn.disabled = false;
      saveBtn.innerHTML = `
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
          <path d="M1.5 6.5l3 3 6-6" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        Save Change`;
    }
  }

  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeInlineEdit(); });

  // ── WRITE CHANGELOG ENTRIES ────────────────────────────────────
  // Appends one row per changed field to the Changelog tab
  async function writeChangelogEntries(changes) {
    const now  = new Date().toISOString().replace('T',' ').slice(0,19); // "YYYY-MM-DD HH:MM:SS"
    const rows = changes.map(c => [
      now,
      userEmail,
      userName,
      currentProg,
      c.qualification,
      c.oldValue,
      c.newValue,
    ]);

    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${encodeURIComponent(CHANGELOG_TAB+'!A:G')}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;

    try {
      const res = await apiFetch(url, {
        method: 'POST',
        body:   JSON.stringify({ majorDimension:'ROWS', values:rows }),
      });
      if (!res.ok) { const e=await res.json().catch(()=>({})); throw new Error(e.error?.message||'HTTP '+res.status); }

      // Update local changelog cache (prepend, newest first)
      const newEntries = rows.map(row => ({
        timestamp:     row[0],
        userEmail:     row[1],
        userName:      row[2],
        programme:     row[3],
        qualification: row[4],
        oldValue:      row[5],
        newValue:      row[6],
      }));
      changelog = [...newEntries, ...changelog];
      buildLastUpdatedMap();
      renderProgList(); // refresh "last updated" on cards

    } catch (err) {
      // Non-fatal — data is saved, just the log failed
      showToast('warn', 'Data saved but changelog could not be written: ' + err.message);
    }
  }

  // ── STEP 5: CHANGELOG ─────────────────────────────────────────
  function renderChangelog() {
    // Populate programme filter dropdown
    const filter = document.getElementById('cl-prog-filter');
    const currentVal = filter.value;
    filter.innerHTML = '<option value="">All programmes</option>' +
      allData.map(p => {
        const name = p[allHeaders[0]];
        return `<option value="${escHtml(name)}" ${name===currentVal?'selected':''}>${escHtml(name)}</option>`;
      }).join('');

    filterChangelog();
  }

  function filterChangelog() {
    const query   = (document.getElementById('cl-search').value || '').trim().toLowerCase();
    const progFilter = document.getElementById('cl-prog-filter').value;
    const wrap    = document.getElementById('changelog-wrap');

    let filtered = changelog;

    if (progFilter) {
      filtered = filtered.filter(e => e.programme === progFilter);
    }

    if (query) {
      filtered = filtered.filter(e =>
        e.programme.toLowerCase().includes(query) ||
        e.qualification.toLowerCase().includes(query) ||
        (e.userName||'').toLowerCase().includes(query) ||
        (e.userEmail||'').toLowerCase().includes(query) ||
        e.newValue.toLowerCase().includes(query) ||
        e.oldValue.toLowerCase().includes(query)
      );
    }

    if (filtered.length === 0) {
      wrap.innerHTML = '<div class="history-empty">' +
        (changelog.length === 0 ? 'No changes recorded yet. Changes will appear here after the first save.' : 'No entries match your filter.') +
        '</div>';
      return;
    }

    const rows = filtered.map(entry => {
      const initials = (entry.userName||entry.userEmail||'?').charAt(0).toUpperCase();
      const oldVal   = entry.oldValue.trim();
      const newVal   = entry.newValue.trim();

      let changeHtml;
      if (!oldVal && newVal) {
        changeHtml = `<span class="cl-added">${escHtml(newVal)}</span>`;
      } else if (oldVal && !newVal) {
        changeHtml = `<span class="cl-removed">${escHtml(oldVal)}</span> → <em style="color:var(--muted);font-size:11px;">removed</em>`;
      } else {
        changeHtml = `<div class="cl-changed">
          <span class="cl-from">${escHtml(oldVal)}</span>
          <span class="cl-to">${escHtml(newVal)}</span>
        </div>`;
      }

      return `<tr>
        <td class="cl-time">${formatTimestamp(entry.timestamp)}</td>
        <td class="cl-user">
          <div class="cl-user-chip">
            <div class="cl-avatar">${initials}</div>
            <div>
              <div style="font-weight:600;font-size:12px;">${escHtml(entry.userName||entry.userEmail)}</div>
              <div style="font-size:11px;color:var(--muted);">${escHtml(entry.userEmail)}</div>
            </div>
          </div>
        </td>
        <td class="cl-prog">${escHtml(entry.programme)}</td>
        <td class="cl-qual">${escHtml(entry.qualification)}</td>
        <td class="cl-change-cell">${changeHtml}</td>
      </tr>`;
    }).join('');

    wrap.innerHTML = `
      <div class="changelog-count">${filtered.length} entr${filtered.length!==1?'ies':'y'}${progFilter||query?' (filtered)':''}</div>
      <div style="overflow-x:auto;">
        <table class="changelog-table">
          <thead>
            <tr>
              <th>When</th>
              <th>Who</th>
              <th>Programme</th>
              <th>Qualification</th>
              <th>Change</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  }

  // ── OUTPUT TABLE ───────────────────────────────────────────────
  function buildOutputTable(fields) {
    const tbody = document.getElementById('output-tbody');
    tbody.innerHTML = '';
    qualHeaders.forEach(h => {
      const req = fields[h] || '';
      if (!req.trim()) return;
      const tr=document.createElement('tr'), td1=document.createElement('td'), s=document.createElement('strong');
      s.textContent=h; td1.appendChild(s);
      const td2=document.createElement('td'); td2.textContent=req;
      tr.appendChild(td1); tr.appendChild(td2); tbody.appendChild(tr);
    });
    document.getElementById('copy-confirm').classList.remove('visible');
  }

  // ── COPY HELPERS ───────────────────────────────────────────────
  function copyTable(btn) {
    const rows = document.querySelectorAll('#output-tbody tr');
    _doCopy(btn, 'copy-confirm', rows);
  }

  function copyTableFromView(btn) {
    const rows = document.querySelectorAll('.view-table tbody tr');
    const cleanRows = Array.from(rows).map(row => ({
      qualName: row.cells[0].childNodes[0].textContent.trim(),
      reqText:  row.cells[1].textContent.trim(),
    }));

    let tableRows = '';
    cleanRows.forEach(({qualName,reqText}) => {
      tableRows += `<tr><td><strong>${qualName}</strong></td><td>${reqText}</td></tr>`;
    });

    const cleanHTML = `<h3>${currentProg}</h3><table border="1" cellpadding="8" cellspacing="0"><thead><tr><th>Qualification</th><th>Requirements</th></tr></thead><tbody>${tableRows}</tbody></table>`;
    const plainText = `${currentProg}\n\nQualification\tRequirements\n` + cleanRows.map(r=>r.qualName+'\t'+r.reqText).join('\n');

    navigator.clipboard.write([new ClipboardItem({
      'text/html':  new Blob([cleanHTML],  {type:'text/html'}),
      'text/plain': new Blob([plainText], {type:'text/plain'}),
    })]).then(() => {
      document.getElementById('view-copy-confirm').classList.add('visible');
      btn.textContent='✓ Copied!'; btn.style.background='#15803d';
      setTimeout(()=>{ btn.textContent='Copy Table'; btn.style.background=''; }, 2500);
    }).catch(()=>alert('Copy failed — please select the table manually.'));
  }

  function _doCopy(btn, confirmId, rows) {
    let tableRows = '';
    rows.forEach(row => {
      const qual = row.cells[0].querySelector('strong')?.textContent || row.cells[0].textContent;
      tableRows += `<tr><td><strong>${qual}</strong></td><td>${row.cells[1].textContent}</td></tr>`;
    });
    const cleanHTML = `<h3>${currentProg}</h3><table border="1" cellpadding="8" cellspacing="0"><thead><tr><th>Qualification</th><th>Requirements</th></tr></thead><tbody>${tableRows}</tbody></table>`;
    const plainText = `${currentProg}\n\nQualification\tRequirements\n` +
      Array.from(rows).map(r=>{
        const q=r.cells[0].querySelector('strong')?.textContent||r.cells[0].textContent;
        return q.trim()+'\t'+r.cells[1].textContent.trim();
      }).join('\n');

    navigator.clipboard.write([new ClipboardItem({
      'text/html':  new Blob([cleanHTML],  {type:'text/html'}),
      'text/plain': new Blob([plainText], {type:'text/plain'}),
    })]).then(()=>{
      document.getElementById(confirmId).classList.add('visible');
      btn.textContent='✓ Copied!'; btn.style.background='#15803d';
      setTimeout(()=>{ btn.textContent='Copy Table'; btn.style.background=''; }, 2500);
    }).catch(()=>alert('Copy failed — please select the table manually.'));
  }

  // ── TOAST ──────────────────────────────────────────────────────
  function showToast(type, html) {
    const icons = {info:'ℹ️',success:'✅',error:'❌',warn:'⚠️'};
    const t = document.getElementById('toast');
    t.className = `toast ${type} visible`;
    t.innerHTML = `<span>${icons[type]||''}</span><span>${html}</span>`;
    if (_toastTimer) clearTimeout(_toastTimer);
    _toastTimer = setTimeout(hideToast, type==='error'?9000:4500);
  }

  function hideToast() {
    const t = document.getElementById('toast');
    if (t) t.className = 'toast';
  }

  // ── API HELPERS ────────────────────────────────────────────────
  function apiUrl(range) {
    return `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${encodeURIComponent(range)}`;
  }

  function apiFetch(url, opts={}) {
    return fetch(url, {
      ...opts,
      headers: { Authorization:'Bearer '+accessToken, 'Content-Type':'application/json', ...(opts.headers||{}) },
    });
  }

  // ── UTILS ──────────────────────────────────────────────────────
  function escHtml(str) {
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function colLetter(n) {
    let s=''; n+=1;
    while(n>0){ n--; s=String.fromCharCode(65+(n%26))+s; n=Math.floor(n/26); }
    return s;
  }

  // Format ISO/stored timestamp to "15 Jan 2025, 14:32"
  function formatTimestamp(ts) {
    if (!ts) return '—';
    try {
      const d = new Date(ts.replace(' ','T')+'Z');
      if (isNaN(d)) return ts;
      return d.toLocaleDateString('en-GB', {day:'numeric',month:'short',year:'numeric'}) +
        ', ' + d.toLocaleTimeString('en-GB', {hour:'2-digit',minute:'2-digit'});
    } catch(_) { return ts; }
  }

  // Human-readable relative time: "2 hours ago", "3 days ago", etc.
  function relativeTime(ts) {
    if (!ts) return '';
    try {
      const d     = new Date(ts.replace(' ','T')+'Z');
      if (isNaN(d)) return ts;
      const diff  = Date.now() - d.getTime();
      const mins  = Math.floor(diff / 60000);
      const hours = Math.floor(diff / 3600000);
      const days  = Math.floor(diff / 86400000);
      if (mins < 2)   return 'just now';
      if (mins < 60)  return mins + 'm ago';
      if (hours < 24) return hours + 'h ago';
      if (days < 30)  return days + 'd ago';
      return formatTimestamp(ts).split(',')[0]; // just the date
    } catch(_) { return ts; }
  }

  // ── PUBLIC API ─────────────────────────────────────────────────
  return {
    init, showScreen,
    goToStep, goToView, goToEdit, openChangelog,
    saveConfig, signIn, signOut,
    filterProgList, clearSearch,
    openView, openEdit,
    onQualInput, saveRequirements,
    openInlineEdit, closeInlineEdit, saveInlineEdit,
    copyTable, copyTableFromView,
    filterChangelog,
  };

})();

document.addEventListener('DOMContentLoaded', () => App.init());