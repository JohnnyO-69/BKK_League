/********************************************************************
 *     MATCH FORM – CANONICAL CURRENT IMPLEMENTATION
 ********************************************************************/
function buildCurrentMatchFormSmart() {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  const paramsSheet = ss.getSheetByName('Parameters');
  if (!fixturesSheet) return;
  if (!paramsSheet) return;

  const canonicalOutName = 'CurrentMatchForm';
  const legacyOutName = 'CurrentMatchFormV2';
  let out = ss.getSheetByName(canonicalOutName);
  if (!out) {
    out = ss.getSheetByName(legacyOutName);
    if (out) {
      out.setName(canonicalOutName);
    } else {
      out = ss.insertSheet(canonicalOutName);
    }
  }

  // Only clear values, preserve formatting
  out.getDataRange().clearContent();

  const teamName = String(paramsSheet.getRange('F7').getValue() || '').trim();
  const teamId = Number(paramsSheet.getRange('F8').getValue());

  if (!teamName || Number.isNaN(teamId)) {
    out.getRange('A1').setValue('Missing team name/id in Parameters!F7:F8');
    return;
  }

  const today = new Date(
    new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' })
  );

  const values = fixturesSheet.getDataRange().getValues();
  const header = values.shift();

  const idxMatchId = header.indexOf('Match ID');
  const idxHome = header.indexOf('Home Team');
  const idxAway = header.indexOf('Away Team');
  const idxDate = header.indexOf('Match Date');

  if ([idxMatchId, idxHome, idxAway, idxDate].some(i => i < 0)) return;

  function toDate(value) {
    return value instanceof Date ? value : new Date(String(value || '').trim());
  }

  const matches = values
    .filter(row => (row[idxHome] === teamName || row[idxAway] === teamName) && row[idxMatchId])
    .map(row => ({
      id: row[idxMatchId],
      date: toDate(row[idxDate]),
      home: row[idxHome],
      away: row[idxAway]
    }))
    .filter(match => !isNaN(match.date.getTime()))
    .sort((left, right) => left.date - right.date);

  if (!matches.length) {
    out.getRange('A1').setValue(`No matches found for ${teamName}`);
    return;
  }

  const past = matches.filter(match => match.date <= today);
  const future = matches.filter(match => match.date > today);

  let chosen;

  if (future.length > 0) {
    const nextFuture = future[0];
    chosen = past.length > 0 ? past[past.length - 1] : nextFuture;
  } else {
    chosen = past[past.length - 1];
  }

  const matchId = chosen.id;
  const matchDate = chosen.date;
  const isHome = chosen.home === teamName;
  const opponent = isHome ? chosen.away : chosen.home;

  function fetchJson(url) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      return response.getResponseCode() === 200 ? JSON.parse(response.getContentText()) : null;
    } catch {
      return null;
    }
  }

  const matchJson = fetchJson('https://api.bkkleague.com/match/details/' + matchId);
  if (!matchJson || !Array.isArray(matchJson.data)) {
    out.getRange('A1').setValue('Unable to load match details');
    return;
  }

  const stats = {};

  function addStat(player, isDouble, won) {
    if (!player) return;

    if (!stats[player]) {
      stats[player] = {
        singlesPlayed: 0,
        singlesWon: 0,
        doublesPlayed: 0,
        doublesWon: 0,
        points: 0
      };
    }

    const playerStats = stats[player];

    if (isDouble) {
      playerStats.doublesPlayed++;
      if (won) {
        playerStats.doublesWon++;
        playerStats.points += 0.5;
      }
    } else {
      playerStats.singlesPlayed++;
      if (won) {
        playerStats.singlesWon++;
        playerStats.points += 1;
      }
    }
  }

  let homeFrames = 0;
  matchJson.data.forEach(frame => {
    const homePlayers = (frame.homePlayers || []).map(player => player.nickName || player.nickname || '').filter(Boolean);
    const awayPlayers = (frame.awayPlayers || []).map(player => player.nickName || player.nickname || '').filter(Boolean);

    const isDouble = homePlayers.length > 1;
    const homeWon = frame.homeWin === 1;

    const homeId = frame.homeTeamId || frame.home_team_id || frame.homeTeamid;
    const awayId = frame.awayTeamId || frame.away_team_id || frame.awayTeamid;

    const ourHome = String(homeId) === String(teamId);
    const ourAway = String(awayId) === String(teamId);

    if (!ourHome && !ourAway) return;

    const ourPlayers = ourHome ? homePlayers : awayPlayers;
    const won = ourHome ? homeWon : !homeWon;

    ourPlayers.forEach(player => addStat(player, isDouble, won));

    if (frame.homeWin !== undefined && homeWon) {
      homeFrames++;
    }
  });

  const decidedFrames = matchJson.data.filter(frame => frame.homeWin !== undefined).length;
  const awayFrames = decidedFrames - homeFrames;
  const ourScore = isHome ? homeFrames : awayFrames;
  const opponentScore = isHome ? awayFrames : homeFrames;
  const result = ourScore > opponentScore ? ' W' : ourScore === opponentScore ? ' D' : ' L';

  const matchSummary = decidedFrames > 0
    ? `Match ID: ${matchId} • ${opponent} (${ourScore}-${opponentScore})${result}`
    : `Match ID: ${matchId} • vs ${opponent}`;

  const rows = Object.entries(stats).map(([name, playerStats]) => {
    const totalGames = playerStats.singlesPlayed + playerStats.doublesPlayed;
    const singlesPct = playerStats.singlesPlayed ? playerStats.singlesWon / playerStats.singlesPlayed : 0;
    const doublesPct = playerStats.doublesPlayed ? playerStats.doublesWon / playerStats.doublesPlayed : 0;
    const ppg = totalGames ? playerStats.points / totalGames : 0;

    return [
      name,
      totalGames,
      playerStats.singlesPlayed,
      playerStats.singlesWon,
      singlesPct,
      playerStats.doublesPlayed,
      playerStats.doublesWon,
      doublesPct,
      playerStats.points,
      ppg
    ];
  });

  rows.sort((left, right) => {
    if (right[9] !== left[9]) return right[9] - left[9];
    if (right[8] !== left[8]) return right[8] - left[8];
    if (right[1] !== left[1]) return right[1] - left[1];
    return left[0].localeCompare(right[0]);
  });

  const totalCols = 10;
  const thaiDate = matchDate.toLocaleDateString('en-CA', { timeZone: 'Asia/Bangkok' });
  const nowStr = new Date().toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' });

  out.getRange(1, 1, 1, totalCols).merge().setValue(
    'Match Form – ' + (isHome ? teamName + ' vs ' + opponent : opponent + ' vs ' + teamName)
  ).setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('white')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(2, 1, 1, totalCols).merge().setValue(`Match Date: ${thaiDate} (Thai Time)`)
    .setFontStyle('italic').setFontColor('#666').setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(3, 1, 1, totalCols).merge().setValue(`Last Refresh: ${nowStr}`)
    .setFontStyle('italic').setFontColor('#666').setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(4, 1, 1, totalCols).merge().setValue(matchSummary)
    .setFontWeight('bold').setFontColor('#333').setHorizontalAlignment('center').setVerticalAlignment('middle');

  const colHeaders = [
    'Player',
    'Games Played',
    'Singles Played',
    'Singles Won',
    'Singles Win %',
    'Doubles Played',
    'Doubles Won',
    'Doubles Win %',
    'Points',
    'Points Per Game'
  ];

  out.getRange(6, 1, 1, totalCols).setValues([colHeaders])
    .setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center').setWrap(true);

  if (rows.length) {
    out.getRange(7, 1, rows.length, totalCols).setValues(rows);

    out.getRange(7, 5, rows.length, 1).setNumberFormat('0.0%');
    out.getRange(7, 8, rows.length, 1).setNumberFormat('0.0%');
    out.getRange(7, 10, rows.length, 1).setNumberFormat('0.00');
    out.getRange(7, 2, rows.length, 9).setHorizontalAlignment('center');
  } else {
    out.getRange(7, 1).setValue('No player stats available');
  }

  for (let column = 1; column <= totalCols; column++) out.setColumnWidth(column, 100);
  const lastRow = rows.length ? rows.length + 6 : 8;
  out.getRange(1, 1, lastRow, totalCols).setBorder(true, true, true, true, true, true);
}

function handleCurrentMatchFormEdit_(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  Logger.log('Edited sheet: ' + sheetName + ', range: ' + range.getA1Notation());

  if ((sheetName === 'CurrentMatchForm' || sheetName === 'CurrentMatchFormV2') &&
      range.getRow() === 18 &&
      range.getColumn() === 1) {

    sheet.getRange('A18').clearContent();
    buildCurrentMatchFormSmart();
  }
}