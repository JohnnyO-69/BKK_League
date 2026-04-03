/********************************************************************
 *     NEXT OPPONENT DETECTION – SILENT BATCH VERSION (NO POPUPS)
 *     Works perfectly in background, on triggers, or in batch
 ********************************************************************/
function buildNextTeamForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  const paramsSheet = ss.getSheetByName('Parameters');
  if (!fixturesSheet) {
    Logger.log('ERROR: Sheet "Fixtures" not found');
    return;
  }
  if (!paramsSheet) {
    Logger.log('ERROR: Sheet "Parameters" not found');
    return;
  }

  const data = fixturesSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('Fixtures sheet is empty');
    return;
  }

  const header = data.shift();
  const idxDate    = header.indexOf('Match Date');
  const idxHome    = header.indexOf('Home Team');
  const idxAway    = header.indexOf('Away Team');
  const idxHomeId  = header.indexOf('Home Team ID');
  const idxAwayId  = header.indexOf('Away Team ID');

  if ([idxDate, idxHome, idxAway, idxHomeId, idxAwayId].some(i => i === -1)) {
    Logger.log('ERROR: Missing required columns in Fixtures sheet');
    return;
  }

  const myTeam = String(paramsSheet.getRange('F7').getValue() || '').trim();

  if (!myTeam) {
    Logger.log('ERROR: Missing team name in Parameters!F7');
    return;
  }

  // Today at 00:00:00 in Thai time
  const todayThai = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' }));
  todayThai.setHours(0, 0, 0, 0);

  function parseMatchDate(v) {
    if (!v) return null;
    let d;
    if (v instanceof Date) {
      d = new Date(v);
    } else {
      const s = String(v).trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        d = new Date(s);
      } else {
        d = new Date(s);
      }
    }
    if (isNaN(d?.getTime())) return null;
    d.setHours(0, 0, 0, 0);
    return d;
  }

  const ourMatches = data
    .map(row => {
      const date = parseMatchDate(row[idxDate]);
      if (!date) return null;
      const home = String(row[idxHome] || '').trim();
      const away = String(row[idxAway] || '').trim();
      if (home !== myTeam && away !== myTeam) return null;

      return {
        date,
        home,
        away,
        homeId: row[idxHomeId],
        awayId: row[idxAwayId]
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.date - b.date);

  if (ourMatches.length === 0) {
    Logger.log(`No matches found for ${myTeam}`);
    return;
  }

  const nextMatch = ourMatches.find(m => m.date >= todayThai);

  if (!nextMatch) {
    Logger.log(`No upcoming match found for ${myTeam}`);
    return;
  }

  const opponentName = nextMatch.home === myTeam ? nextMatch.away : nextMatch.home;
  const opponentId   = nextMatch.home === myTeam ? nextMatch.awayId : nextMatch.homeId;

  // Silent logging only (visible in View → Logs)
  Logger.log(`Next opponent detected: ${opponentName} (ID: ${opponentId}) on ${nextMatch.date.toLocaleDateString('en-GB')}`);

  // Build the two reports silently
  buildLast3FormForTeam(opponentId, opponentName, 'NextTeam3MatchForm');
  buildSeasonFormForTeam(opponentId, opponentName, 'NextTeamSeasonForm');
  buildNextTeamOOP();
}