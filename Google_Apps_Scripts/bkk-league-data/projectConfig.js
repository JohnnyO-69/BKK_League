// Maps each Apps Script project to its bound spreadsheet.
// ScriptApp.getScriptId() returns the ID of whichever deployment is running,
// so the same source file works correctly for both projects.
// Updated: 2026-04-06 – bkk-league-data
const SCRIPT_TO_SPREADSHEET_ = {
  // BKK League Data
  '1gP5BJz1Lpz3-XLQmiKYIi_y0nARwfpED8zjoyhK6gO1OgBAvtYllamCt': '1Kcv1y5bQX8YGxnIIXyKYj5QkcQQO_qBO5Zcvt6lSMAU',
  // Team Sheet
  '1CWO4vZaW5FTQ9yRjghI4zg3yryKUt0jgW-y7tT8lLHxffNSA3Hg31vKl': '1zz1rk8E_r3dDxkMUR1_30igrPXKVZu7ju0SkRMaYJf4'
};

function getLeagueSpreadsheet_() {
  const scriptId = ScriptApp.getScriptId();
  const spreadsheetId = SCRIPT_TO_SPREADSHEET_[scriptId];
  if (!spreadsheetId) throw new Error('getLeagueSpreadsheet_: no spreadsheet mapped for script ID ' + scriptId);
  return SpreadsheetApp.openById(spreadsheetId);
}