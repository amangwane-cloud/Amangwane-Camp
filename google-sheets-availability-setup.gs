/**
 * Builds the Amangwane year-view availability board.
 * Run buildAvailabilityBoard() once in the Google Sheets workbook.
 */
function buildAvailabilityBoard() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Availability Board';
  const existing = workbook.getSheetByName(sheetName);
  const sheet = existing || workbook.insertSheet(sheetName);
  sheet.clear();
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  const headers = ['Date', 'Day', 'Holiday / School Holiday',
    'Tent 1', 'Tent 2', 'Tent 3', 'Tent 4', 'Tent 5', 'Tent 6', 'Tent 7', 'Tent 8', 'Tent 9', 'Tent 10'];
  const rows = [headers];
  const start = new Date(2026, 0, 1);
  const end = new Date(2026, 11, 31);
  const holidays = southAfrican2026Holidays();

  for (let date = new Date(start); date <= end; date.setDate(date.getDate() + 1)) {
    const key = formatDate(date);
    const marker = holidays[key] || '';
    rows.push([new Date(date), Utilities.formatDate(date, Session.getScriptTimeZone(), 'EEE'), marker,
      'Available', 'Available', 'Available', 'Available', 'Available', 'Available', 'Available', 'Available', 'Available', 'Available']);
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(2, 1, rows.length - 1, 1).setNumberFormat('dd mmm yyyy');
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1C1C2A').setFontColor('#FFFFFF');
  sheet.getRange(1, 1, rows.length, headers.length).setVerticalAlignment('middle');
  sheet.getRange(2, 1, rows.length - 1, 3).setBackground('#F7F3EA');
  sheet.getRange(2, 4, rows.length - 1, 10).setBackground('#EAF3EA').setFontColor('#2A6E2A');
  sheet.getRange(1, 1, rows.length, headers.length).createFilter();
  sheet.setColumnWidth(1, 105);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidths(4, 10, 95);
  sheet.getRange(1, 1, rows.length, headers.length).setWrap(true);

  const holidayRows = [];
  for (let row = 2; row <= rows.length; row++) {
    if (rows[row - 1][2]) holidayRows.push(row);
  }
  holidayRows.forEach(row => sheet.getRange(row, 1, 1, headers.length).setBackground('#D9D9D9').setFontColor('#555555'));

  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Available').setBackground('#EAF3EA').setFontColor('#2A6E2A').setRanges([sheet.getRange(2, 4, rows.length - 1, 10)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Confirmed').setBackground('#DCEFF6').setFontColor('#075D79').setRanges([sheet.getRange(2, 4, rows.length - 1, 10)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Deposit pending').setBackground('#FFF4C7').setFontColor('#7B6000').setRanges([sheet.getRange(2, 4, rows.length - 1, 10)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Maintenance').setBackground('#EEEEEE').setFontColor('#777777').setRanges([sheet.getRange(2, 4, rows.length - 1, 10)]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function southAfrican2026Holidays() {
  const dates = {
    '2026-01-01': 'Public holiday: New Year\'s Day',
    '2026-03-21': 'Public holiday: Human Rights Day',
    '2026-04-03': 'Public holiday: Good Friday',
    '2026-04-06': 'Public holiday: Family Day',
    '2026-04-27': 'Public holiday: Freedom Day',
    '2026-05-01': 'Public holiday: Workers\' Day',
    '2026-06-16': 'Public holiday: Youth Day',
    '2026-08-09': 'Public holiday: National Women\'s Day',
    '2026-08-10': 'Public holiday: Public holiday observed',
    '2026-09-24': 'Public holiday: Heritage Day',
    '2026-12-16': 'Public holiday: Day of Reconciliation',
    '2026-12-25': 'Public holiday: Christmas Day',
    '2026-12-26': 'Public holiday: Day of Goodwill',
    '2026-12-28': 'Public holiday: Public holiday observed'
  };
  const schoolBreaks = [
    ['2026-01-01', '2026-01-13', 'School holiday'],
    ['2026-03-28', '2026-04-07', 'School holiday'],
    ['2026-06-27', '2026-07-20', 'School holiday'],
    ['2026-09-24', '2026-10-05', 'School holiday'],
    ['2026-12-10', '2026-12-31', 'School holiday']
  ];
  schoolBreaks.forEach(([from, to, label]) => {
    const date = new Date(from + 'T00:00:00');
    const end = new Date(to + 'T00:00:00');
    while (date <= end) {
      const key = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      dates[key] = dates[key] ? dates[key] + ' + ' + label : label;
      date.setDate(date.getDate() + 1);
    }
  });
  return dates;
}
