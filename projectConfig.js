const LEAGUE_SPREADSHEET_ID = '1Kcv1y5bQX8YGxnIIXyKYj5QkcQQO_qBO5Zcvt6lSMAU';
const LEAGUE_SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1Kcv1y5bQX8YGxnIIXyKYj5QkcQQO_qBO5Zcvt6lSMAU/edit?gid=680648455#gid=680648455';

function getLeagueSpreadsheet_() {
  return SpreadsheetApp.openById(LEAGUE_SPREADSHEET_ID);
}