/* ═══════════════════════════════════════════════════════════════
   Entry Requirements Manager — app.js

   Screens:    setup → login → denied | app
   Steps:      1:select  2:view  3:edit  4:output
   Modal:      inline-edit for single qualification quick-save
═══════════════════════════════════════════════════════════════ */

const App = (() => {

  // ── CONFIG ─────────────────────────────────────────────────────
  const CFG = {
    get clientId()  { return localStorage.getItem('cfg_client_id')  || ''; },
    get sheetId()   { return localStorage.getItem('cfg_sheet_id')   || ''; },
    set clientId(v) { localStorage.setItem('cfg_client_id',  v); },
    set sheetId(v)  { localStorage.setItem('cfg_sheet_id',   v); },
  };

  // ── STATE ──────────────────────────────────────────────────────
  let accessToken    = null;
  let userEmail      = null;
  let userName       = null;
  let userPicture    = null;
  let allHeaders     = [];   // ["Programme Name", "SPM / O-Level", …]
  let qualHeaders    = [];   // allHeaders minus index 0
  let allData        = [];   // array of row objects
  let currentProg    = null; // programme name string
  let currentRowIdx  = null; // index into allData
  let inlineQualKey  = null; // qual header being inline-edited
  let _toastTimer    = null;

  const SHEET_NAME = 'Entry Requirements';
  const ACCESS_TAB = 'Access';
  const STEPS      = ['step-select', 'step-view', 'step-edit', 'step-output'];

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

  // goToStep — public, handles cleanup per step
  function goToStep(n) {
    hideToast();
    if (n === 1) {
      currentProg   = null;
      currentRowIdx = null;
      renderProgList(); // refresh fill counts after any saves
      // Restore search state
      const q = document.getElementById('prog-search').value;
      filterProgList(q);
    }
    showStep(n);
  }

  // Convenience helpers used by HTML buttons
  function goToView() {
    renderViewTable();
    showStep(2);
  }

  function goToEdit() {
    renderQualGrid(allData[currentRowIdx]);
    showStep(3);
  }

  // ── CONFIG SAVE ────────────────────────────────────────────────
  function saveConfig() {
    const sheetId  = document.getElementById('cfg-sheet-id').value.trim();
    const clientId = document.getElementById('cfg-client-id').value.trim();

    if (!sheetId)  { alert('Please enter your Spreadsheet ID.');    document.getElementById('cfg-sheet-id').focus();  return; }
    if (!clientId) { alert('Please enter your OAuth Client ID.');   document.getElementById('cfg-client-id').focus(); return; }

    CFG.sheetId  = sheetId;
    CFG.clientId = clientId;
    showScreen('screen-login');
  }

  // ── AUTH ───────────────────────────────────────────────────────
  function signIn() {
    if (!CFG.clientId || !CFG.sheetId) { showScreen('screen-setup'); return; }

    if (!window.google?.accounts?.oauth2) {
      alert('Google sign-in library not loaded. Check your internet connection and reload the page.');
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
        error_callback: (err) => {
          alert('Sign-in error: ' + (err.message || err.type || 'Unknown error') +
            '\n\nMake sure popups are allowed for this page.');
        },
      });
    } catch (err) {
      alert('Could not initialise Google sign-in: ' + err.message +
        '\n\nDouble-check your OAuth Client ID in the setup screen.');
      return;
    }

    tokenClient.requestAccessToken({ prompt: 'select_account' });
  }

  async function onTokenResponse(resp) {
    if (resp.error) {
      alert('Sign-in failed: ' + resp.error +
        (resp.error_description ? '\n' + resp.error_description : ''));
      return;
    }

    accessToken = resp.access_token;

    try {
      const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: 'Bearer ' + accessToken },
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const p   = await res.json();
      userEmail   = p.email;
      userName    = p.name;
      userPicture = p.picture;
    } catch (err) {
      alert('Signed in but could not fetch profile: ' + err.message);
      return;
    }

    await checkAccess();
  }

  async function checkAccess() {
    try {
      const url = apiUrl(ACCESS_TAB + '!A:A');
      const res = await apiFetch(url);
      if (!res.ok) {
        const e = await res.json().catch(() => ({}));
        throw new Error(e.error?.message || 'HTTP ' + res.status);
      }
      const data    = await res.json();
      const allowed = (data.values || []).flat()
        .map(e => String(e).trim().toLowerCase())
        .filter(e => e.includes('@'));

      if (allowed.includes(userEmail.toLowerCase())) {
        onAccessGranted();
      } else {
        onAccessDenied();
      }
    } catch (err) {
      alert('Could not verify access: ' + err.message +
        '\n\nMake sure:\n• The "' + ACCESS_TAB + '" tab exists in your Google Sheet\n• Your Spreadsheet ID is correct');
      showScreen('screen-login');
    }
  }

  function onAccessGranted() {
    const avatar = document.getElementById('user-avatar');
    if (userPicture) {
      avatar.innerHTML = `<img src="${userPicture}" alt="" />`;
    } else {
      avatar.textContent = (userName || userEmail).charAt(0).toUpperCase();
    }
    document.getElementById('user-name-display').textContent = userName || userEmail;
    document.getElementById('topbar-user').classList.add('visible');

    showScreen('screen-app');
    showStep(1);
    loadSheetData();
  }

  function onAccessDenied() {
    document.getElementById('denied-email').textContent = userEmail;
    showScreen('screen-denied');
  }

  function signOut() {
    accessToken = userEmail = userName = userPicture = null;
    allData = []; currentProg = null; currentRowIdx = null;
    document.getElementById('topbar-user').classList.remove('visible');
    if (window.google?.accounts?.oauth2 && accessToken) {
      google.accounts.oauth2.revoke(accessToken);
    }
    showScreen('screen-login');
  }

  // ── LOAD SHEET DATA ────────────────────────────────────────────
  async function loadSheetData() {
    renderProgListSkeleton();
    try {
      const res = await apiFetch(apiUrl(SHEET_NAME + '!A1:ZZ'));
      if (!res.ok) {
        const e = await res.json().catch(() => ({}));
        throw new Error(e.error?.message || 'HTTP ' + res.status);
      }
      const data = await res.json();
      const rows = data.values || [];

      if (rows.length < 2) {
        renderProgListMsg('No programmes found. Make sure the "' + SHEET_NAME + '" tab has a header row and at least one data row.');
        return;
      }

      allHeaders  = rows[0].map(h => String(h).trim());
      qualHeaders = allHeaders.slice(1);
      allData     = rows.slice(1)
        .filter(r => String(r[0] || '').trim() !== '')
        .map(row => {
          const obj = {};
          allHeaders.forEach((h, i) => { obj[h] = String(row[i] || '').trim(); });
          return obj;
        });

      renderProgList();
    } catch (err) {
      renderProgListMsg('Failed to load data: ' + err.message);
      showToast('error', '<strong>Could not load sheet data.</strong> ' + err.message);
    }
  }

  // ── STEP 1: PROGRAMME LIST ─────────────────────────────────────
  function renderProgListSkeleton() {
    document.getElementById('prog-list').innerHTML =
      [1,2,3,4].map(() => '<div class="skeleton"></div>').join('');
  }

  function renderProgListMsg(msg) {
    document.getElementById('prog-list').innerHTML =
      `<p class="prog-empty">${msg}</p>`;
  }

  function renderProgList() {
    const progKey = allHeaders[0];
    const total   = qualHeaders.length;

    document.getElementById('prog-list').innerHTML = allData.map((prog, i) => {
      const name   = escHtml(prog[progKey] || '(Unnamed)');
      const filled = qualHeaders.filter(h => prog[h] && prog[h].trim()).length;
      const pct    = total > 0 ? Math.round((filled / total) * 100) : 0;

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

    // Re-apply active search filter if any
    const q = document.getElementById('prog-search')?.value || '';
    if (q) filterProgList(q);
  }

  // ── SEARCH ─────────────────────────────────────────────────────
  function filterProgList(query) {
    const q         = query.trim().toLowerCase();
    const cards     = document.querySelectorAll('#prog-list .prog-card');
    const clearBtn  = document.getElementById('search-clear');
    const emptyMsg  = document.getElementById('prog-empty');
    let   visible   = 0;

    clearBtn.classList.toggle('hidden', q === '');

    cards.forEach(card => {
      const name    = card.dataset.name || '';
      const matches = !q || name.includes(q);
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

  // ── OPEN VIEW ──────────────────────────────────────────────────
  function openView(idx) {
    setCurrentProg(idx);
    renderViewTable();
    showStep(2);
  }

  function renderViewTable() {
    const prog  = allData[currentRowIdx];
    const filled = qualHeaders.filter(h => prog[h] && prog[h].trim());

    document.getElementById('view-prog-name').textContent = currentProg;
    document.getElementById('view-meta').textContent =
      filled.length + ' of ' + qualHeaders.length + ' qualifications apply to this programme';

    const wrap = document.getElementById('view-table-wrap');

    if (filled.length === 0) {
      wrap.innerHTML = `
        <div class="view-empty">
          No entry requirements have been filled in yet for this programme.<br>
          <button class="btn btn-primary btn-sm" style="margin-top:16px;" onclick="App.goToEdit()">Fill in requirements</button>
        </div>`;
      return;
    }

    const rows = filled.map(h => {
      const req     = escHtml(prog[h]);
      const hEsc    = escHtml(h);
      return `
        <tr>
          <td>
            ${hEsc}
            <button class="inline-edit-btn" onclick="App.openInlineEdit('${h.replace(/'/g, "\\'")}')">
              <svg width="10" height="10" viewBox="0 0 12 12" fill="none">
                <path d="M8 1.5l2.5 2.5-6 6H2v-2.5l6-6z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/>
              </svg>
              Edit
            </button>
          </td>
          <td>${req}</td>
        </tr>`;
    }).join('');

    wrap.innerHTML = `
      <table class="view-table">
        <thead>
          <tr><th>Qualification</th><th>Requirements</th></tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>`;

    document.getElementById('view-copy-confirm').classList.remove('visible');
  }

  // ── OPEN EDIT (full form) ──────────────────────────────────────
  function openEdit(idx) {
    setCurrentProg(idx);
    renderQualGrid(allData[idx]);
    document.getElementById('edit-prog-name').textContent = currentProg;
    showStep(3);
  }

  function setCurrentProg(idx) {
    currentProg   = allData[idx][allHeaders[0]];
    currentRowIdx = idx;
  }

  // ── STEP 3: QUAL GRID ──────────────────────────────────────────
  function renderQualGrid(prog) {
    document.getElementById('edit-prog-name').textContent = currentProg;
    document.getElementById('qual-grid').innerHTML = qualHeaders.map((header, i) => {
      const val    = prog[header] || '';
      const filled = val.trim() !== '';
      return `
        <div class="qual-item ${filled ? 'has-value' : ''}" id="qi-${i}">
          <div class="qual-header" onclick="document.getElementById('qa-${i}').focus()">
            <div class="qual-name">
              ${escHtml(header)}
              <span class="qual-badge ${filled ? 'filled' : 'empty'}" id="qb-${i}">
                ${filled ? 'Filled' : 'Empty'}
              </span>
            </div>
          </div>
          <div class="qual-body">
            <textarea
              id="qa-${i}"
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
    document.getElementById('qi-' + i).classList.toggle('has-value', filled);
    const badge = document.getElementById('qb-' + i);
    badge.className   = 'qual-badge ' + (filled ? 'filled' : 'empty');
    badge.textContent = filled ? 'Filled' : 'Empty';
    updateFilledCount();
  }

  function updateFilledCount() {
    const filled = document.querySelectorAll('#qual-grid .qual-item.has-value').length;
    document.getElementById('filled-count').textContent =
      filled + ' of ' + qualHeaders.length + ' filled';
  }

  // ── SAVE FULL FORM ─────────────────────────────────────────────
  async function saveRequirements() {
    const fields = {};
    qualHeaders.forEach((h, i) => {
      fields[h] = (document.getElementById('qa-' + i)?.value || '').trim();
    });

    const newRow   = allHeaders.map((h, i) => i === 0 ? currentProg : (fields[h] || ''));
    const sheetRow = currentRowIdx + 2;
    const range    = `${SHEET_NAME}!A${sheetRow}:${colLetter(allHeaders.length - 1)}${sheetRow}`;
    const url      = `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;

    const btn = document.getElementById('save-btn');
    btn.disabled = true;
    btn.innerHTML = '<div class="spinner"></div> Saving…';

    try {
      const res = await apiFetch(url, {
        method: 'PUT',
        body:   JSON.stringify({ range, majorDimension: 'ROWS', values: [newRow] }),
      });

      if (!res.ok) {
        const e = await res.json().catch(() => ({}));
        throw new Error(e.error?.message || 'HTTP ' + res.status);
      }

      // Update local cache
      allHeaders.forEach((h, i) => { allData[currentRowIdx][h] = newRow[i]; });

      // Build output and go to step 4
      document.getElementById('output-prog-name').textContent = currentProg;
      buildOutputTable(fields);
      showStep(4);
      showToast('success', `"${currentProg}" saved to Google Sheets.`);

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
    const currentVal = allData[currentRowIdx][qualKey] || '';

    document.getElementById('modal-prog-label').textContent  = currentProg;
    document.getElementById('modal-qual-name').textContent   = qualKey;
    document.getElementById('modal-textarea').value          = currentVal;

    document.getElementById('modal-overlay').classList.add('open');
    document.getElementById('inline-modal').classList.add('open');

    // Focus textarea after animation
    setTimeout(() => document.getElementById('modal-textarea').focus(), 100);
  }

  function closeInlineEdit() {
    document.getElementById('modal-overlay').classList.remove('open');
    document.getElementById('inline-modal').classList.remove('open');
    inlineQualKey = null;
  }

  async function saveInlineEdit() {
    if (!inlineQualKey) return;

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
        body:   JSON.stringify({ range, majorDimension: 'ROWS', values: [[newVal]] }),
      });

      if (!res.ok) {
        const e = await res.json().catch(() => ({}));
        throw new Error(e.error?.message || 'HTTP ' + res.status);
      }

      // Update local cache
      allData[currentRowIdx][inlineQualKey] = newVal;

      closeInlineEdit();
      renderViewTable(); // refresh the view table with new value
      showToast('success', `"${inlineQualKey}" updated successfully.`);

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

  // Close modal on Escape key
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeInlineEdit();
  });

  // ── BUILD OUTPUT TABLE ─────────────────────────────────────────
  function buildOutputTable(fields) {
    const tbody = document.getElementById('output-tbody');
    tbody.innerHTML = '';
    qualHeaders.forEach(h => {
      const req = fields[h] || '';
      if (!req.trim()) return;
      const tr  = document.createElement('tr');
      const td1 = document.createElement('td');
      const s   = document.createElement('strong');
      s.textContent = h;
      td1.appendChild(s);
      const td2 = document.createElement('td');
      td2.textContent = req;
      tr.appendChild(td1);
      tr.appendChild(td2);
      tbody.appendChild(tr);
    });
    document.getElementById('copy-confirm').classList.remove('visible');
  }

  // ── COPY TABLE (from output step) ─────────────────────────────
  function copyTable(btn) {
    const rows = document.querySelectorAll('#output-tbody tr');
    _doCopy(btn, 'copy-confirm', rows, currentProg);
  }

  // ── COPY TABLE (from view step) ───────────────────────────────
  function copyTableFromView(btn) {
    const rows = document.querySelectorAll('.view-table tbody tr');

    // Build row data from view table (qual name + req text, skip inline-edit btn text)
    const cleanRows = Array.from(rows).map(row => {
      // First cell contains qual name + the edit button — get just the text node
      const qualCell = row.cells[0];
      const qualName = qualCell.childNodes[0].textContent.trim();
      const reqText  = row.cells[1].textContent.trim();
      return { qualName, reqText };
    });

    let tableRows = '';
    cleanRows.forEach(({ qualName, reqText }) => {
      tableRows += `<tr><td><strong>${qualName}</strong></td><td>${reqText}</td></tr>`;
    });

    const cleanHTML =
      `<h3>${currentProg}</h3>` +
      `<table border="1" cellpadding="8" cellspacing="0">` +
      `<thead><tr><th>Qualification</th><th>Requirements</th></tr></thead>` +
      `<tbody>${tableRows}</tbody></table>`;

    const plainText =
      `${currentProg}\n\nQualification\tRequirements\n` +
      cleanRows.map(r => r.qualName + '\t' + r.reqText).join('\n');

    navigator.clipboard.write([
      new ClipboardItem({
        'text/html':  new Blob([cleanHTML],  { type: 'text/html'  }),
        'text/plain': new Blob([plainText], { type: 'text/plain' }),
      }),
    ]).then(() => {
      document.getElementById('view-copy-confirm').classList.add('visible');
      btn.textContent      = '✓ Copied!';
      btn.style.background = '#15803d';
      setTimeout(() => {
        btn.textContent      = 'Copy Table';
        btn.style.background = '';
      }, 2500);
    }).catch(() => {
      alert('Copy failed — please select the table manually and copy.');
    });
  }

  // Shared copy helper for output step
  function _doCopy(btn, confirmId, rows, progName) {
    let tableRows = '';
    rows.forEach(row => {
      const qual = row.cells[0].querySelector('strong')?.textContent || row.cells[0].textContent;
      const req  = row.cells[1].textContent;
      tableRows += `<tr><td><strong>${qual}</strong></td><td>${req}</td></tr>`;
    });

    const cleanHTML =
      `<h3>${progName}</h3>` +
      `<table border="1" cellpadding="8" cellspacing="0">` +
      `<thead><tr><th>Qualification</th><th>Requirements</th></tr></thead>` +
      `<tbody>${tableRows}</tbody></table>`;

    const plainText =
      `${progName}\n\nQualification\tRequirements\n` +
      Array.from(rows).map(r => {
        const q = r.cells[0].querySelector('strong')?.textContent || r.cells[0].textContent;
        return q.trim() + '\t' + r.cells[1].textContent.trim();
      }).join('\n');

    navigator.clipboard.write([
      new ClipboardItem({
        'text/html':  new Blob([cleanHTML],  { type: 'text/html'  }),
        'text/plain': new Blob([plainText], { type: 'text/plain' }),
      }),
    ]).then(() => {
      document.getElementById(confirmId).classList.add('visible');
      btn.textContent      = '✓ Copied!';
      btn.style.background = '#15803d';
      setTimeout(() => {
        btn.textContent      = 'Copy Table';
        btn.style.background = '';
      }, 2500);
    }).catch(() => {
      alert('Copy failed — please select the table manually and copy.');
    });
  }

  // ── TOAST ──────────────────────────────────────────────────────
  function showToast(type, html) {
    const icons = { info: 'ℹ️', success: '✅', error: '❌', warn: '⚠️' };
    const t = document.getElementById('toast');
    t.className = `toast ${type} visible`;
    t.innerHTML = `<span>${icons[type] || ''}</span><span>${html}</span>`;
    if (_toastTimer) clearTimeout(_toastTimer);
    _toastTimer = setTimeout(hideToast, type === 'error' ? 9000 : 4000);
  }

  function hideToast() {
    const t = document.getElementById('toast');
    if (t) t.className = 'toast';
  }

  // ── API HELPERS ────────────────────────────────────────────────
  function apiUrl(range) {
    return `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${encodeURIComponent(range)}`;
  }

  function apiFetch(url, opts = {}) {
    return fetch(url, {
      ...opts,
      headers: {
        Authorization:  'Bearer ' + accessToken,
        'Content-Type': 'application/json',
        ...(opts.headers || {}),
      },
    });
  }

  // ── UTILS ──────────────────────────────────────────────────────
  function escHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function colLetter(n) {
    // 0-based index → spreadsheet column letter (0→A, 25→Z, 26→AA…)
    let s = '';
    n += 1;
    while (n > 0) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  }

  // ── PUBLIC API ─────────────────────────────────────────────────
  return {
    // Boot
    init,
    // Navigation
    showScreen,
    goToStep,
    goToView,
    goToEdit,
    // Auth
    saveConfig,
    signIn,
    signOut,
    // Prog list
    filterProgList,
    clearSearch,
    openView,
    openEdit,
    // Edit form
    onQualInput,
    saveRequirements,
    // Inline edit modal
    openInlineEdit,
    closeInlineEdit,
    saveInlineEdit,
    // Copy
    copyTable,
    copyTableFromView,
  };

})();

document.addEventListener('DOMContentLoaded', () => App.init());