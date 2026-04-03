/********************************************************************
 *     FINAL PREDICTION REFRESH – WHOLE NUMBERS ONLY
 *     • 10-10 → Draw = 28% guaranteed
 *     • 11-9 / 9-11 → Draw = 22%
 *     • 12-8 / 8-12 → Draw = 15%
 *     • Everything else → 8%
 ********************************************************************/
function refreshPrediction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('NextTeam3MatchForm');
  const ourStatsSheet = ss.getSheetByName('Last3MatchForm');
  const oppStatsSheet = ss.getSheetByName('NextTeam3MatchForm');
  const fixturesSheet = ss.getSheetByName('Fixtures');
  const paramsSheet = ss.getSheetByName('Parameters');
  if (!sheet) return;
  if (!ourStatsSheet) return;
  if (!oppStatsSheet) return;
  if (!fixturesSheet) return;
  if (!paramsSheet) return;

  function extractTeamName(value) {
    return String(value || '')
      .replace(/^.*[–-]\s*/, '')
      .trim();
  }

  function parseMatchDate(value) {
    if (!value) return null;
    const date = value instanceof Date ? new Date(value) : new Date(String(value).trim());
    if (isNaN(date.getTime())) return null;
    date.setHours(0, 0, 0, 0);
    return date;
  }

  function getNextOpponentName() {
    const data = fixturesSheet.getDataRange().getValues();
    if (data.length <= 1) return '';

    const header = data.shift();
    const idxDate = header.indexOf('Match Date');
    const idxHome = header.indexOf('Home Team');
    const idxAway = header.indexOf('Away Team');

    if ([idxDate, idxHome, idxAway].some(i => i === -1)) return '';

    const todayThai = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' }));
    todayThai.setHours(0, 0, 0, 0);

    const ourMatches = data
      .map(row => {
        const date = parseMatchDate(row[idxDate]);
        if (!date) return null;

        const home = String(row[idxHome] || '').trim();
        const away = String(row[idxAway] || '').trim();
        if (home !== ourTeam && away !== ourTeam) return null;

        return { date, home, away };
      })
      .filter(Boolean)
      .sort((a, b) => a.date - b.date);

    const nextMatch = ourMatches.find(match => match.date >= todayThai);
    if (!nextMatch) return '';

    return nextMatch.home === ourTeam ? nextMatch.away : nextMatch.home;
  }

  const r = 26;
  const scoreRow = r + 3;
  const probRow  = r + 5;

  // 1. FULL CLEAR
  sheet.getRange('F26:L31').clearContent().clearFormat();

  // 2. TITLE + MERGED EMPTY ROWS
  sheet.getRange(r, 6, 1, 7).merge()
    .setValue('PREDICTED SCORE (20 frames)')
    .setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center')
    .setBackground('#fff2cc');

  sheet.getRange('F27:L27').merge();
  sheet.getRange('F30:L30').merge();

  // 3. TEAM NAMES
  const ourTeam = String(paramsSheet.getRange('F7').getValue() || '').trim() ||
    sheet.getRange('A1').getValue().match(/– (.+)$/)?.[1]?.trim() ||
    'Home';
  const nextOpponentName = getNextOpponentName();
  const opponentCandidates = [
    nextOpponentName,
    oppStatsSheet.getRange('A1').getValue(),
    oppStatsSheet.getRange('A22').getValue(),
    oppStatsSheet.getRange('F1').getValue(),
    oppStatsSheet.getRange('F22').getValue()
  ]
    .map(extractTeamName)
    .filter(name => name && name !== ourTeam);

  const oppTeam = opponentCandidates[0] || 'Away';

  sheet.getRange(r + 2, 6, 1, 3).merge().setValue(ourTeam)
    .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange(r + 2, 10, 1, 3).merge().setValue(oppTeam)
    .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');

  // 4. STRENGTH CALCULATION (unchanged)
  const ourWinPct  = (ourStatsSheet.getRange('G5').getValue()  + ourStatsSheet.getRange('J5').getValue())  / 2;
  const oppWinPct  = (oppStatsSheet.getRange('G5').getValue() + oppStatsSheet.getRange('J5').getValue()) / 2;
  const ourPPG     = ourStatsSheet.getRange('L5').getValue();
  const oppPPG     = oppStatsSheet.getRange('L5').getValue();

  const ourStrength = ourWinPct + ourPPG * 0.4;
  const oppStrength = oppWinPct + oppPPG * 0.4;
  const total = ourStrength + oppStrength || 1;

  // FORCE WHOLE NUMBER FROM THE VERY BEGINNING
  let finalScore = Math.round(20 * ourStrength / total);   // ← already whole
  if (finalScore > 15) finalScore += 1;
  if (finalScore < 5)  finalScore -= 1;
  finalScore = Math.max(0, Math.min(20, finalScore));

  // 5. BIG SCORE (always whole numbers)
  sheet.getRange(scoreRow, 6, 1, 3).merge()
    .setValue(finalScore)
    .setFontSize(42).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(scoreRow, 9).setValue('–')
    .setFontSize(42).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(scoreRow, 10, 1, 3).merge()
    .setValue(20 - finalScore)
    .setFontSize(42).setFontWeight('bold').setHorizontalAlignment('center');

  // 6. DRAW PROBABILITY BASED ON WHOLE-NUMBER DIFFERENCE
  const diffFromTen = Math.abs(finalScore - 10);
  const drawProb = diffFromTen === 0 ? "28%" :
                   diffFromTen === 1 ? "22%" :
                   diffFromTen === 2 ? "15%" : "8%";

  const probText = `${ourTeam}: ${Math.round(finalScore/20*100)}%  Draw: ${drawProb}  ${oppTeam}: ${Math.round((20-finalScore)/20*100)}%`;

  sheet.getRange(probRow, 6, 1, 7).merge()
    .setValue(probText)
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  // 7. YELLOW BOX
  sheet.getRange(r, 6, 6, 7)
    .setBackground('#fff2cc')
    .setBorder(true, true, true, true, true, true);
}