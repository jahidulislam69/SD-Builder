// Excel Audit.xlsx importer using ExcelJS
(function (root) {
  const SD = (root.SD = root.SD || {});

  function unwrapCell(val) {
    if (val && typeof val === 'object' && 'result' in val && !(val instanceof Date)) {
      return val.result;
    }
    return val;
  }

  function toIsoDateUTC(val) {
    if (val instanceof Date && !isNaN(val.getTime())) {
      const pad = (n) => String(n).padStart(2, '0');
      return `${val.getUTCFullYear()}-${pad(val.getUTCMonth() + 1)}-${pad(val.getUTCDate())}`;
    }
    if (typeof val === 'string') {
      const trimmed = val.trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
      const parsed = new Date(trimmed);
      if (!isNaN(parsed.getTime())) {
        const pad = (n) => String(n).padStart(2, '0');
        return `${parsed.getFullYear()}-${pad(parsed.getMonth() + 1)}-${pad(parsed.getDate())}`;
      }
    }
    return null;
  }

  function findSundayOfWeek(isoDateStr) {
    const [y, m, d] = isoDateStr.split('-').map(Number);
    const dt = new Date(y, m - 1, d, 12, 0, 0);
    dt.setDate(dt.getDate() - dt.getDay());
    const pad = (n) => String(n).padStart(2, '0');
    return `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
  }

  /**
   * Parses an Audit.xlsx ArrayBuffer/Buffer and aggregates daily Voice & Non-Voice counts.
   * Rule: L1 Activities 'Digital' -> Non-Voice; 'Inbound' / 'Outbound' -> Voice.
   */
  SD.parseAuditWorkbook = async function (buffer) {
    const ExcelJS = root.ExcelJS;
    if (!ExcelJS) {
      throw new Error('ExcelJS library is not loaded.');
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const sheet = workbook.worksheets[0];
    if (!sheet) {
      throw new Error('No worksheet found in the uploaded Excel file.');
    }

    // Detect column indices from header row (fallback to 1=Date, 4=L1 Activities, last=Audit)
    const headerRow = sheet.getRow(1);
    let dateCol = 1;
    let l1Col = 4;
    let countCol = sheet.columnCount || 6;

    headerRow.eachCell((cell, colNum) => {
      const txt = String(unwrapCell(cell.value) || '').trim().toLowerCase();
      if (txt === 'date') dateCol = colNum;
      else if (txt.includes('l1')) l1Col = colNum;
      else if (txt === 'audit' || txt.includes('count') || txt.includes('total')) countCol = colNum;
    });

    const dailyCounts = {};
    let totalVoice = 0;
    let totalNonVoice = 0;

    for (let r = 2; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const rawDate = unwrapCell(row.getCell(dateCol).value);
      const isoDate = toIsoDateUTC(rawDate);
      if (!isoDate) continue; // Skip 'Total' or blank rows

      const l1Raw = String(unwrapCell(row.getCell(l1Col).value) || '').trim().toLowerCase();
      const count = Number(unwrapCell(row.getCell(countCol).value)) || 0;

      if (!dailyCounts[isoDate]) {
        dailyCounts[isoDate] = { voice: 0, nonVoice: 0 };
      }

      if (l1Raw.includes('digital')) {
        dailyCounts[isoDate].nonVoice += count;
        totalNonVoice += count;
      } else if (l1Raw.includes('inbound') || l1Raw.includes('outbound')) {
        dailyCounts[isoDate].voice += count;
        totalVoice += count;
      }
    }

    const sortedDates = Object.keys(dailyCounts).sort();
    if (sortedDates.length === 0) {
      throw new Error('No valid dated audit rows found in the uploaded file.');
    }

    const weekSunday = findSundayOfWeek(sortedDates[0]);

    return {
      weekSunday,
      dailyCounts,
      sortedDates,
      totalVoice,
      totalNonVoice
    };
  };
})(typeof globalThis !== 'undefined' ? globalThis : window);
