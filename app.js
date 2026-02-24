/* ═══════════════════════════════════════════════════════════════
   Entry Requirements Manager — app.js
   
   Architecture:
   - App object contains all state and methods
   - showScreen() strictly controls which screen is visible
   - showStep()   strictly controls which step panel is visible
   - No href="#", no form submits — zero accidental scroll/nav
═══════════════════════════════════════════════════════════════ */

const App = (() => {

  // ─── CONFIG (persisted in localStorage) ──────────────────────
  const CFG = {
    get clientId()  { return localStorage.getItem('cfg_client_id')  || ''; },
    get sheetId()   { return localStorage.getItem('cfg_sheet_id')   || ''; },
    set clientId(v) { localStorage.setItem('cfg_client_id',  v); },
    set sheetId(v)  { localStorage.setItem('cfg_sheet_id',   v); },
  };

  // ─── STATE ────────────────────────────────────────────────────
  let accessToken   = null;
  let userEmail     = null;
  let userName      = null;
  let userPicture   = null;
  let allHeaders    = [];   // ["Programme Name", "SPM / O-Level", …]
  let qualHeaders   = [];   // allHeaders minus index 0
  let allData       = [];   // array of row objects
  let currentProg   = null;
  let currentRowIdx = null;
  let _toastTimer   = null;

  const SHEET_NAME = 'Entry Requirements';
  const ACCESS_TAB = 'Access';

  // ─── INIT ─────────────────────────────────────────────────────
  function init() {
    if (!CFG.clientId || !CFG.sheetId) {
      showScreen('screen-setup');
    } else {
      showScreen('screen-login');
    }
  }

  // ─── SCREEN MANAGEMENT ────────────────────────────────────────
  // One screen visible at a time. Screens are top-level sections.
  function showScreen(id) {
    document.querySelectorAll('.screen').forEach(el => {
      el.classList.remove('active');
    });
    const target = document.getElementById(id);
    if (target) target.classList.add('active');
    window.scrollTo(0, 0);
  }

  // ─── STEP MANAGEMENT ──────────────────────────────────────────
  // One step panel visible at a time within screen-app.
  function showStep(n) {
    document.querySelectorAll('.step-panel').forEach(el => {
      el.classList.remove('active');
    });
    const panels = ['step-select', 'step-edit', 'step-output'];
    const target = document.getElementById(panels[n - 1]);
    if (target) target.classList.add('active');
    window.scrollTo(0, 0);
  }

  // Public alias used by HTML onclick attributes
  function goToStep(n) {
    if (n === 1) {
      currentProg    = null;
      currentRowIdx  = null;
      hideToast();
      renderProgList(); // refresh fill counts
    }
    showStep(n);
  }

  // ─── CONFIG SAVE ──────────────────────────────────────────────
  function saveConfig() {
    const clientId = document.getElementById('cfg-client-id').value.trim();
    const sheetId  = document.getElementById('cfg-sheet-id').value.trim();

    if (!sheetId) {
      alert('Please enter your Spreadsheet ID.');
      document.getElementById('cfg-sheet-id').focus();
      return;
    }
    if (!clientId) {
      alert('Please enter your OAuth Client ID.');
      document.getElementById('cfg-client-id').focus();
      return;
    }

    CFG.clientId = clientId;
    CFG.sheetId  = sheetId;
    showScreen('screen-login');
  }

  // ─── AUTH: SIGN IN ────────────────────────────────────────────
  function signIn() {
    if (!CFG.clientId || !CFG.sheetId) {
      showScreen('screen-setup');
      return;
    }

    if (!window.google?.accounts?.oauth2) {
      alert('Google sign-in library not loaded. Please check your internet connection and reload the page.');
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
        callback: onTokenResponse,
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

  async function onTokenResponse(tokenResponse) {
    if (tokenResponse.error) {
      alert('Sign-in failed: ' + tokenResponse.error +
        (tokenResponse.error_description ? '\n' + tokenResponse.error_description : ''));
      return;
    }

    accessToken = tokenResponse.access_token;

    // Fetch user profile
    try {
      const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: 'Bearer ' + accessToken },
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const profile = await res.json();
      userEmail   = profile.email;
      userName    = profile.name;
      userPicture = profile.picture;
    } catch (err) {
      alert('Signed in but could not fetch your profile: ' + err.message);
      return;
    }

    await checkAccess();
  }

  // ─── AUTH: ACCESS CHECK ───────────────────────────────────────
  async function checkAccess() {
    try {
      const range = encodeURIComponent(ACCESS_TAB + '!A:A');
      const url   = `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${range}`;
      const res   = await apiFetch(url);

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error?.message || 'HTTP ' + res.status);
      }

      const data    = await res.json();
      const allowed = (data.values || [])
        .flat()
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
    // Update topbar user display
    const avatar = document.getElementById('user-avatar');
    if (userPicture) {
      avatar.innerHTML = `<img src="${userPicture}" alt="" />`;
    } else {
      avatar.textContent = (userName || userEmail).charAt(0).toUpperCase();
    }
    document.getElementById('user-name-display').textContent = userName || userEmail;
    document.getElementById('topbar-user').classList.add('visible');

    // Show app screen and load data
    showScreen('screen-app');
    showStep(1);
    loadSheetData();
  }

  function onAccessDenied() {
    document.getElementById('denied-email').textContent = userEmail;
    showScreen('screen-denied');
  }

  // ─── AUTH: SIGN OUT ───────────────────────────────────────────
  function signOut() {
    accessToken   = null;
    userEmail     = null;
    userName      = null;
    userPicture   = null;
    allData       = [];
    currentProg   = null;
    currentRowIdx = null;

    document.getElementById('topbar-user').classList.remove('visible');

    // Revoke token so Google shows account picker next time
    if (window.google?.accounts?.oauth2 && accessToken) {
      google.accounts.oauth2.revoke(accessToken);
    }

    showScreen('screen-login');
  }

  // ─── SHEETS: LOAD DATA ────────────────────────────────────────
  async function loadSheetData() {
    renderProgListSkeleton();

    try {
      const range = encodeURIComponent(SHEET_NAME + '!A1:ZZ');
      const url   = `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${range}`;
      const res   = await apiFetch(url);

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error?.message || 'HTTP ' + res.status);
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

  // ─── STEP 1: PROGRAMME LIST ───────────────────────────────────
  function renderProgListSkeleton() {
    document.getElementById('prog-list').innerHTML =
      [1, 2, 3, 4].map(() => '<div class="skeleton"></div>').join('');
  }

  function renderProgListMsg(msg) {
    document.getElementById('prog-list').innerHTML =
      `<p style="padding:24px 0;text-align:center;color:var(--muted);font-size:13px;">${msg}</p>`;
  }

  function renderProgList() {
    const progKey = allHeaders[0];
    const total   = qualHeaders.length;

    if (allData.length === 0) {
      renderProgListMsg('No programmes found in the sheet.');
      return;
    }

    document.getElementById('prog-list').innerHTML = allData.map((prog, i) => {
      const name   = escHtml(prog[progKey] || '(Unnamed)');
      const filled = qualHeaders.filter(h => prog[h] && prog[h].trim()).length;
      return `
        <div class="prog-card" data-idx="${i}" onclick="App.selectProgramme(${i})">
          <div class="prog-icon">
            <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
              <rect x="3" y="2" width="14" height="16" rx="2" stroke="currentColor" stroke-width="1.5"/>
              <path d="M6 7h8M6 10h8M6 13h5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
            </svg>
          </div>
          <div>
            <div class="prog-name">${name}</div>
            <div class="prog-meta">${filled} of ${total} qualifications filled</div>
          </div>
        </div>`;
    }).join('');
  }

  // ─── STEP 2: EDIT REQUIREMENTS ───────────────────────────────
  function selectProgramme(idx) {
    // Highlight selected card briefly
    document.querySelectorAll('.prog-card').forEach(c => c.classList.remove('selected'));
    const card = document.querySelector(`.prog-card[data-idx="${idx}"]`);
    if (card) card.classList.add('selected');

    currentProg    = allData[idx][allHeaders[0]];
    currentRowIdx  = idx;

    document.getElementById('edit-prog-name').textContent = currentProg;
    renderQualGrid(allData[idx]);
    showStep(2);
  }

  function renderQualGrid(prog) {
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
              placeholder="Leave blank if this qualification is not accepted for this programme…"
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
    badge.className  = 'qual-badge ' + (filled ? 'filled' : 'empty');
    badge.textContent = filled ? 'Filled' : 'Empty';
    updateFilledCount();
  }

  function updateFilledCount() {
    const filled = document.querySelectorAll('.qual-item.has-value').length;
    document.getElementById('filled-count').textContent =
      `${filled} of ${qualHeaders.length} filled`;
  }

  // ─── STEP 2 → SAVE TO SHEETS ──────────────────────────────────
  async function saveRequirements() {
    // Collect values
    const fields = {};
    qualHeaders.forEach((h, i) => {
      fields[h] = (document.getElementById('qa-' + i)?.value || '').trim();
    });

    // Build full row in header order
    const newRow = allHeaders.map((h, i) =>
      i === 0 ? currentProg : (fields[h] || '')
    );

    // Sheet row number (1-based; row 1 = headers, so data starts at 2)
    const sheetRow = currentRowIdx + 2;
    const endCol   = colLetter(allHeaders.length - 1);
    const range    = `${SHEET_NAME}!A${sheetRow}:${endCol}${sheetRow}`;
    const url      = `https://sheets.googleapis.com/v4/spreadsheets/${CFG.sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;

    // Update button state
    const btn = document.getElementById('save-btn');
    btn.disabled = true;
    btn.innerHTML = '<div class="spinner"></div> Saving…';

    try {
      const res = await apiFetch(url, {
        method: 'PUT',
        body: JSON.stringify({
          range,
          majorDimension: 'ROWS',
          values: [newRow],
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.error?.message || 'HTTP ' + res.status);
      }

      // Update local cache so the programme list shows fresh counts
      allHeaders.forEach((h, i) => { allData[currentRowIdx][h] = newRow[i]; });

      buildOutputTable(fields);
      showStep(3);
      showToast('success', `"${currentProg}" saved to Google Sheets successfully.`);

    } catch (err) {
      showToast('error', '<strong>Save failed:</strong> ' + err.message);
    } finally {
      btn.disabled = false;
      btn.innerHTML = `
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
          <path d="M2 7l3.5 3.5L12 3" stroke="white" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        Save &amp; Generate Table`;
    }
  }

  // ─── STEP 3: BUILD OUTPUT TABLE ───────────────────────────────
  function buildOutputTable(fields) {
    const tbody = document.getElementById('output-tbody');
    tbody.innerHTML = '';

    qualHeaders.forEach(h => {
      const req = fields[h] || '';
      if (!req.trim()) return; // skip blank — not applicable

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

  // ─── STEP 3: COPY TABLE ───────────────────────────────────────
  // Builds a clean HTML string from scratch — no computed styles
  function copyTable(btn) {
    const rows = document.querySelectorAll('#output-tbody tr');
    let tableRows = '';
    rows.forEach(row => {
      const qual = row.cells[0].querySelector('strong').textContent;
      const req  = row.cells[1].textContent;
      tableRows += `<tr><td><strong>${qual}</strong></td><td>${req}</td></tr>`;
    });

    const cleanHTML =
      `<h3>${currentProg}</h3>` +
      `<table border="1" cellpadding="8" cellspacing="0">` +
      `<thead><tr><th>Qualification</th><th>Requirements</th></tr></thead>` +
      `<tbody>${tableRows}</tbody>` +
      `</table>`;

    const plainText =
      `${currentProg}\n\nQualification\tRequirements\n` +
      Array.from(rows)
        .map(r => r.cells[0].textContent.trim() + '\t' + r.cells[1].textContent.trim())
        .join('\n');

    navigator.clipboard.write([
      new ClipboardItem({
        'text/html':  new Blob([cleanHTML],  { type: 'text/html'  }),
        'text/plain': new Blob([plainText], { type: 'text/plain' }),
      }),
    ]).then(() => {
      document.getElementById('copy-confirm').classList.add('visible');
      btn.textContent      = '✓ Copied!';
      btn.style.background = '#15803d';
      setTimeout(() => {
        btn.textContent      = 'Copy Table';
        btn.style.background = '';
      }, 2500);
    }).catch(() => {
      alert('Copy failed — please select the table manually and copy (Ctrl+C / Cmd+C).');
    });
  }

  // ─── TOAST ────────────────────────────────────────────────────
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

  // ─── API HELPER ───────────────────────────────────────────────
  function apiFetch(url, opts = {}) {
    return fetch(url, {
      ...opts,
      headers: {
        Authorization: 'Bearer ' + accessToken,
        'Content-Type': 'application/json',
        ...(opts.headers || {}),
      },
    });
  }

  // ─── UTILS ────────────────────────────────────────────────────
  function escHtml(str) {
    return String(str)
      .replace(/&/g,  '&amp;')
      .replace(/</g,  '&lt;')
      .replace(/>/g,  '&gt;')
      .replace(/"/g,  '&quot;');
  }

  // Convert 0-based column index to spreadsheet letter (0→A, 25→Z, 26→AA…)
  function colLetter(n) {
    let s = '';
    n += 1;
    while (n > 0) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  }

  // ─── PUBLIC API ───────────────────────────────────────────────
  return {
    init,
    showScreen,
    saveConfig,
    signIn,
    signOut,
    goToStep,
    selectProgramme,
    onQualInput,
    saveRequirements,
    copyTable,
  };

})();

// ─── BOOT ─────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => App.init());