// Google Sheets live JSONP sync for SD Daily Report Generator (works on file:// with zero CORS)
(function (root) {
  const SD = (root.SD = root.SD || {});

  SD.DEFAULT_SHEET_URL =
    'https://docs.google.com/spreadsheets/d/1_86VrIo-lmZ63YrKsWsWTA62bnf7HcOaBssSctYixoM/edit?usp=sharing';

  const MONTH_MAP = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
  };

  SD.extractSheetId = function (urlOrId) {
    const str = String(urlOrId || '').trim();
    const match = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : str;
  };

  /**
   * Parses gviz date cell (supports "Date(2026,10,8)" 0-indexed month and "8-Nov-26").
   */
  function normalizeGvizDate(cell) {
    if (!cell) return null;
    const pad = (n) => String(n).padStart(2, '0');

    // gviz Date(Y, M, D) where M is 0-indexed
    if (typeof cell.v === 'string') {
      const gvizMatch = cell.v.match(/^Date\((\d{4}),\s*(\d{1,2}),\s*(\d{1,2})\)$/);
      if (gvizMatch) {
        const y = Number(gvizMatch[1]);
        const m = Number(gvizMatch[2]) + 1;
        const d = Number(gvizMatch[3]);
        return `${y}-${pad(m)}-${pad(d)}`;
      }
    }

    const raw = String(cell.f || cell.v || '').trim();
    if (!raw || raw.toLowerCase() === 'total' || raw.toLowerCase() === 'grand total') return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

    const m = raw.match(/^(\d{1,2})[-/ ]([A-Za-z]{3,9})[-/ ](\d{2,4})$/);
    if (m) {
      const day = Number(m[1]);
      const monKey = m[2].slice(0, 3).toLowerCase();
      const month = MONTH_MAP[monKey];
      let year = Number(m[3]);
      if (year < 100) year += 2000;
      if (month && day >= 1 && day <= 31) {
        return `${year}-${pad(month)}-${pad(day)}`;
      }
    }

    return null;
  }

  /**
   * Loads gviz JSON via <script> JSONP in browser (bypasses file:// CORS),
   * or via fetch() in Node.js test environments.
   */
  function loadGvizData(sheetId) {
    if (typeof document === 'undefined') {
      const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?headers=2&tqx=out:json`;
      return fetch(url).then(async (res) => {
        if (!res.ok) {
          throw new Error(`Google Sheets returned HTTP ${res.status}.`);
        }
        const txt = await res.text();
        if (txt.trim().startsWith('<')) {
          throw new Error('Google Sheets returned a sign-in page. Ensure "Anyone with the link (Viewer)" is enabled.');
        }
        return JSON.parse(txt.substring(txt.indexOf('{'), txt.lastIndexOf('}') + 1));
      });
    }

    return new Promise((resolve, reject) => {
      const cbName = `__sdGvizCb_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
      const script = document.createElement('script');
      let settled = false;

      const cleanup = () => {
        delete root[cbName];
        if (script.parentNode) script.parentNode.removeChild(script);
      };

      const timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        cleanup();
        reject(new Error('Timed out connecting to Google Sheets. Check your internet connection or sheet sharing permissions.'));
      }, 15000);

      root[cbName] = (payload) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        cleanup();
        if (payload && payload.status === 'error') {
          const msg = payload.errors?.[0]?.detailed_message || payload.errors?.[0]?.message || 'Google Sheets query error.';
          reject(new Error(msg));
        } else {
          resolve(payload);
        }
      };

      script.onerror = () => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        cleanup();
        reject(new Error('Could not load Google Sheet. Check your internet connection and ensure the sheet is shared as "Anyone with the link (Viewer)".'));
      };

      script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?headers=2&tqx=responseHandler:${cbName}`;
      document.head.appendChild(script);
    });
  }

  /**
   * Fetches the Google Sheet and extracts daily Complete Voice & Non-Voice counts
   * for the specified Processor Name.
   */
  SD.syncFromGoogleSheet = async function (sheetUrlOrId, processorName) {
    const cleanName = String(processorName || '').trim();
    if (!cleanName) {
      throw new Error('Please enter a Processor Name before clicking Sync.');
    }

    const sheetId = SD.extractSheetId(sheetUrlOrId || SD.DEFAULT_SHEET_URL);
    if (!sheetId) {
      throw new Error('Invalid Google Sheet URL or ID.');
    }

    const payload = await loadGvizData(sheetId);
    const cols = payload?.table?.cols || [];
    const rows = payload?.table?.rows || [];

    if (cols.length < 10 || rows.length === 0) {
      throw new Error('The Google Sheet does not contain expected audit columns or rows.');
    }

    const availableNames = [];
    let matchedColIdx = -1;
    let matchedExactName = '';

    for (let c = 5; c < cols.length; c++) {
      const label = String(cols[c]?.label || '').trim();
      if (!label) continue;
      if (/^grand\s*total/i.test(label) || /^daily\s*audit/i.test(label)) break;

      // Strip trailing "Distribuation Count" / "Distribution Count" from merged 2-row header
      if (/distribu[a-z]*\s*count$/i.test(label)) {
        const userName = label.replace(/\s*distribu[a-z]*\s*count$/i, '').trim();
        availableNames.push(userName);
        if (userName.toLowerCase() === cleanName.toLowerCase()) {
          matchedColIdx = c;
          matchedExactName = userName;
        }
      }
    }

    if (matchedColIdx === -1) {
      const list = availableNames.length ? availableNames.join(', ') : 'None found';
      throw new Error(`Name "${cleanName}" not found in sheet. Available names: ${list}`);
    }

    // Find "Complete" sub-column within this user's 4-column block
    let completeCol = -1;
    for (let offset = 0; offset < 4; offset++) {
      const subLabel = String(cols[matchedColIdx + offset]?.label || '').trim().toLowerCase();
      if (subLabel === 'complete') {
        completeCol = matchedColIdx + offset;
        break;
      }
    }

    if (completeCol === -1) {
      throw new Error(`Could not find the "Complete" sub-column under "${matchedExactName}".`);
    }

    const dailyCounts = {};
    for (const row of rows) {
      const cells = row.c || [];
      const isoDate = normalizeGvizDate(cells[0]);
      if (!isoDate) continue;

      const l1 = String(cells[3]?.f || cells[3]?.v || '').trim().toLowerCase();
      const val = Number(cells[completeCol]?.v) || 0;

      if (!dailyCounts[isoDate]) {
        dailyCounts[isoDate] = { voice: 0, nonVoice: 0 };
      }

      if (l1.includes('digital')) {
        dailyCounts[isoDate].nonVoice += val;
      } else if (l1.includes('inbound') || l1.includes('outbound')) {
        dailyCounts[isoDate].voice += val;
      }
    }

    return {
      matchedName: matchedExactName,
      availableNames,
      dailyCounts
    };
  };
})(typeof globalThis !== 'undefined' ? globalThis : window);
