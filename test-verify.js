// Comprehensive verification script for SD Report Generator
import assert from 'node:assert';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
globalThis.ExcelJS = ExcelJS;
import './js/presets.js';
import './js/calculator.js';
import './js/importer.js';
import './js/sheets-sync.js';

const { resolveWeekDates, calculateDayAllocation, generateReportRows, parseAuditWorkbook, syncFromGoogleSheet, DEFAULT_SHEET_URL, APP_CONSTANTS } = globalThis.SD;

console.log('--- 1. Testing resolveWeekDates & Date Math ---');
const dates = resolveWeekDates('2026-11-08');
assert.strictEqual(dates.length, 5, 'Must resolve 5 days (Sunday to Thursday)');
assert.strictEqual(dates[0].dayName, 'Sunday');
assert.strictEqual(dates[4].dayName, 'Thursday');

// Test 3-day receiving date offset rule
const pad = (n) => String(n).padStart(2, '0');
const toIso = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

const sundayComp = toIso(dates[0].completionDate);
const sundayRec = toIso(dates[0].receivingDate);
assert.strictEqual(sundayComp, '2026-11-08');
assert.strictEqual(sundayRec, '2026-11-05', 'Receiving date must be 3 calendar days prior');
console.log('✓ Date resolution and 3-day receiving date offset passed!');

console.log('--- 2. Testing calculateDayAllocation & 420-Minute Balancing ---');
// Case A: Standard Day with Meeting/Calibration deduction
const dayWithMeeting = {
  type: 'NORMAL',
  voiceCount: 31,
  nonVoiceCount: 30,
  sessions: [
    { source: 'Calibration', time: 60 }
  ]
};
const allocMeeting = calculateDayAllocation(dayWithMeeting);
assert.strictEqual(allocMeeting.baseAuditPool, 420);
assert.strictEqual(allocMeeting.sessionMinutes, 60);
assert.strictEqual(allocMeeting.availableAuditPool, 360, 'Audit pool must deduct meeting/calibration time (420 - 60 = 360)');
assert.strictEqual(allocMeeting.voiceTime + allocMeeting.nonVoiceTime, 360, 'Voice + Non-Voice must equal 360m');
assert.strictEqual(allocMeeting.totalDayTime, 540, 'Total day time must equal 540 minutes');
assert.strictEqual(allocMeeting.isBalanced, true);
console.log(`✓ Meeting deduction: 420 - 60 = ${allocMeeting.availableAuditPool}m audit pool (Voice: ${allocMeeting.voiceTime}m, NV: ${allocMeeting.nonVoiceTime}m)`);

// Case B: Half-Day Leave
const halfDay = {
  type: 'HALF_DAY_LEAVE',
  voiceCount: 0,
  nonVoiceCount: 30,
  sessions: []
};
const allocHalfDay = calculateDayAllocation(halfDay);
assert.strictEqual(allocHalfDay.baseAuditPool, 210, 'Half day audit target is 210m');
assert.strictEqual(allocHalfDay.leaveTime, 270, 'Half day leave row takes 270m');
assert.strictEqual(allocHalfDay.breakTime, 30);
assert.strictEqual(allocHalfDay.docTime, 30);
assert.strictEqual(allocHalfDay.totalDayTime, 540, 'Half day (270 + 210 + 30 + 30) must equal 540m');
assert.strictEqual(allocHalfDay.isBalanced, true);
console.log('✓ Half-day allocation passed: 270m leave + 210m audit + 30m break + 30m doc = 540m total');

// Case C: Full Day Leave
const fullLeave = {
  type: 'SICK_LEAVE',
  voiceCount: 0,
  nonVoiceCount: 0,
  sessions: []
};
const allocLeave = calculateDayAllocation(fullLeave);
assert.strictEqual(allocLeave.totalDayTime, 0);
assert.strictEqual(allocLeave.isBalanced, true);
console.log('✓ Full day leave passed: 0m total');

console.log('--- 3. Testing generateReportRows Schema & SL numbering ---');
const weekData = dates.map((dateMeta, idx) => ({
  index: idx,
  type: (idx === 3) ? 'SICK_LEAVE' : 'NORMAL',
  voiceCount: 40,
  nonVoiceCount: 20,
  sessions: (idx === 1) ? [{ source: 'Calibration', level1: 'Calibration', level2: 'Calibration with Mia Bhai', level3: 'Calibration about fail calls', quantity: 1, time: 60 }] : [],
  dateMeta
}));

const profile = {
  name: 'Intakhab Didar Rony',
  qcBy: 'Md Helal Hossain',
  unit: 'QA Executive'
};

const rows = generateReportRows(weekData, profile);
assert.ok(rows.length >= 15, `Generated rows count (${rows.length}) should be at least 15`);
rows.forEach((row, i) => {
  assert.strictEqual(row[0], i + 1, 'SL no must be continuous sequence starting at 1');
  assert.strictEqual(row.length, 19, 'Each row must have exactly 19 columns');
  assert.strictEqual(row[4], 'QA Executive', 'Unit must be QA Executive');
  assert.strictEqual(row[15], 'Md Helal Hossain', 'QC By must match profile');
  assert.strictEqual(row[16], 'Office Work', 'Category must be Office Work');
});
console.log(`✓ Row schema valid: ${rows.length} rows generated with 19 columns each`);

console.log('--- 4. Testing Excel Workbook Output with ExcelJS ---');
const wb = new ExcelJS.Workbook();
const sData = wb.addWorksheet('SD Data');
sData.columns = APP_CONSTANTS.COLUMN_WIDTHS.map((width, idx) => ({
  key: `col_${idx + 1}`,
  width
}));

sData.addRow(APP_CONSTANTS.HEADERS);
rows.forEach(r => sData.addRow(r));

const s2 = wb.addWorksheet('Sheet2');
s2.getCell('A3').value = 'Production';
s2.getCell('C3').value = 'Yes';
s2.getCell('A4').value = 'Non Production';
s2.getCell('C4').value = 'No';

const testFilePath = 'd:\\SD Report Generator\\test_generated.xlsx';
await wb.xlsx.writeFile(testFilePath);
assert.ok(fs.existsSync(testFilePath), 'Generated excel file must exist on disk');
console.log('✓ Successfully wrote test_generated.xlsx');

console.log('--- 5. Testing Audit.xlsx Importer ---');
const auditBuffer = fs.readFileSync('d:\\SD Report Generator\\Audit.xlsx');
const imported = await parseAuditWorkbook(auditBuffer);
assert.strictEqual(imported.weekSunday, '2026-11-08');
assert.deepStrictEqual(imported.dailyCounts['2026-11-08'], { voice: 61, nonVoice: 0 });
assert.deepStrictEqual(imported.dailyCounts['2026-11-09'], { voice: 40, nonVoice: 10 });
assert.deepStrictEqual(imported.dailyCounts['2026-11-10'], { voice: 26, nonVoice: 34 });
assert.deepStrictEqual(imported.dailyCounts['2026-11-11'], { voice: 0, nonVoice: 20 });
assert.deepStrictEqual(imported.dailyCounts['2026-11-12'], { voice: 42, nonVoice: 8 });
console.log('✓ Audit.xlsx parsed cleanly:', imported.totalVoice, 'Voice,', imported.totalNonVoice, 'Non-Voice');

console.log('--- 6. Testing Live Google Sheet Sync ---');
const sheetRes = await syncFromGoogleSheet(DEFAULT_SHEET_URL, 'Intakhab Didar Rony');
assert.strictEqual(sheetRes.matchedName, 'Intakhab Didar Rony');
assert.deepStrictEqual(sheetRes.dailyCounts['2026-11-08'], { voice: 51, nonVoice: 0 });
assert.deepStrictEqual(sheetRes.dailyCounts['2026-11-09'], { voice: 31, nonVoice: 30 });
assert.deepStrictEqual(sheetRes.dailyCounts['2026-11-10'], { voice: 26, nonVoice: 35 });
assert.deepStrictEqual(sheetRes.dailyCounts['2026-11-11'], { voice: 0, nonVoice: 0 });
assert.deepStrictEqual(sheetRes.dailyCounts['2026-11-12'], { voice: 46, nonVoice: 14 });
await assert.rejects(() => syncFromGoogleSheet(DEFAULT_SHEET_URL, 'Unknown Person'), /not found in sheet/);
console.log('✓ Live Google Sheet Sync passed for Intakhab Didar Rony & invalid-name error handling!');

console.log('\n=== ALL TESTS PASSED SUCCESSFULLY! ===');
