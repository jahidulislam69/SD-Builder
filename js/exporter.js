// Excel generator engine using ExcelJS
(function (root) {
  const SD = (root.SD = root.SD || {});

  SD.exportReportToExcel = async function (weekData, profile) {
    const ExcelJS = root.ExcelJS;
    if (!ExcelJS) {
      throw new Error('ExcelJS library is not loaded.');
    }

    const { APP_CONSTANTS, generateReportRows } = SD;
    const workbook = new ExcelJS.Workbook();
    workbook.creator = profile.name || 'QA Executive';
    workbook.lastModifiedBy = profile.name || 'QA Executive';
    workbook.created = new Date();
    workbook.modified = new Date();

    // 1. Create SD Data sheet first (so it opens by default)
    const sdSheet = workbook.addWorksheet('SD Data', {
      views: [{ state: 'normal', activeCell: 'A1', showGridLines: true }]
    });

    // 2. Create Sheet2 (Data validation lookups)
    const sheet2 = workbook.addWorksheet('Sheet2');
    sheet2.getCell('A3').value = 'Production';
    sheet2.getCell('C3').value = 'Yes';
    sheet2.getCell('A4').value = 'Non Production';
    sheet2.getCell('C4').value = 'No';

    // Set column widths
    sdSheet.columns = APP_CONSTANTS.COLUMN_WIDTHS.map((width, idx) => ({
      key: `col_${idx + 1}`,
      width
    }));

    // Add header row
    const headerRow = sdSheet.addRow(APP_CONSTANTS.HEADERS);
    headerRow.height = 30;

    headerRow.eachCell((cell) => {
      cell.font = {
        name: 'Calibri',
        size: 11,
        bold: true,
        color: { argb: 'FFFFFFFF' }
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1F4E79' }
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFBFBFBF' } },
        left: { style: 'thin', color: { argb: 'FFBFBFBF' } },
        bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
        right: { style: 'thin', color: { argb: 'FFBFBFBF' } }
      };
    });

    // Generate data rows
    const rawRows = generateReportRows(weekData, profile);

    rawRows.forEach((rowValues) => {
      const row = sdSheet.addRow(rowValues);
      row.height = 20;

      row.eachCell((cell, colNumber) => {
        cell.font = {
          name: 'Calibri',
          size: 11
        };

        cell.border = {
          top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
          left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
          bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
          right: { style: 'thin', color: { argb: 'FFD9D9D9' } }
        };

        if ([1, 4, 6, 7, 14, 15, 18, 19].includes(colNumber)) {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        } else if ([11, 12, 13].includes(colNumber)) {
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
          cell.numFmt = '[$-409]d-mmm-yyyy;@';
        } else {
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
        }
      });

      const rowIdx = row.number;
      sdSheet.getCell(`F${rowIdx}`).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['Sheet2!$A$3:$A$4']
      };
      sdSheet.getCell(`G${rowIdx}`).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: ['Sheet2!$C$3:$C$4']
      };
    });

    const buffer = await workbook.xlsx.writeBuffer();

    const lastDay = weekData[weekData.length - 1];
    const endDate = lastDay.dateMeta.completionDate;
    const dd = String(endDate.getDate()).padStart(2, '0');
    const mm = String(endDate.getMonth() + 1).padStart(2, '0');
    const yyyy = endDate.getFullYear();
    const safeName = (profile.name || 'QA_Report').trim().replace(/[^a-zA-Z0-9 _-]/g, '');
    const fileName = `Daily SD_${safeName}_${dd}-${mm}-${yyyy}.xlsx`;

    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    return { fileName, rowCount: rawRows.length };
  };
})(typeof globalThis !== 'undefined' ? globalThis : window);
