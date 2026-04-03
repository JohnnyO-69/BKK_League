/********************************************************************
 *     MATCH FORM V2 – WITH FORMATTING SIMILAR TO LAST 3 FORM
 ********************************************************************/
function buildCurrentMatchFormSmartV2() {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  const paramsSheet = ss.getSheetByName('Parameters');
  if (!fixturesSheet) return;
  if (!paramsSheet) return;

  const outName = 'CurrentMatchFormV2';
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);

  // Only clear values, preserve formatting
  out.getDataRange().clearContent();

  const teamName = String(paramsSheet.getRange('F7').getValue() || '').trim();
  const teamId = Number(paramsSheet.getRange('F8').getValue());

  if (!teamName || Number.isNaN(teamId)) {
    out.getRange('A1').setValue('Missing team name/id in Parameters!F7:F8');
    return;
  }

  // Thailand timezone "today"
  const today = new Date(
    new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' })
  );

  // Load fixtures
  const values = fixturesSheet.getDataRange().getValues();
  const header = values.shift();

  const idxMatchId = header.indexOf('Match ID');
  const idxHome = header.indexOf('Home Team');
  const idxAway = header.indexOf('Away Team');
  const idxDate = header.indexOf('Match Date');

  if ([idxMatchId, idxHome, idxAway, idxDate].some(i => i < 0)) return;

  function toDate(v) {
    return v instanceof Date ? v : new Date(String(v || '').trim());
  }

  // All matches involving The Game 8B
  const matches = values
    .filter(r => (r[idxHome] === teamName || r[idxAway] === teamName) && r[idxMatchId])
    .map(r => ({
      id: r[idxMatchId],
      date: toDate(r[idxDate]),
      home: r[idxHome],
      away: r[idxAway]
    }))
    .filter(m => !isNaN(m.date.getTime()))
    .sort((a, b) => a.date - b.date);

  if (!matches.length) {
    out.getRange('A1').setValue(`No matches found for ${teamName}`);
    return;
  }

  // Split past and future
  const past = matches.filter(m => m.date <= today);
  const future = matches.filter(m => m.date > today);

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

  // Fetch match frames
  function fetchJson(url) {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
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

    const s = stats[player];

    if (isDouble) {
      s.doublesPlayed++;
      if (won) {
        s.doublesWon++;
        s.points += 0.5;
      }
    } else {
      s.singlesPlayed++;
      if (won) {
        s.singlesWon++;
        s.points += 1;
      }
    }
  }

  // Process frames and calculate scores
  let homeFrames = 0;
  matchJson.data.forEach(f => {
    const hp = (f.homePlayers || []).map(p => p.nickName || p.nickname || '').filter(Boolean);
    const ap = (f.awayPlayers || []).map(p => p.nickName || p.nickname || '').filter(Boolean);

    const isDouble = hp.length > 1;
    const homeWon = f.homeWin === 1;

    const homeId = f.homeTeamId || f.home_team_id || f.homeTeamid;
    const awayId = f.awayTeamId || f.away_team_id || f.awayTeamid;

    const ourHome = String(homeId) === String(teamId);
    const ourAway = String(awayId) === String(teamId);

    if (!ourHome && !ourAway) return;

    const ourPlayers = ourHome ? hp : ap;
    const won = ourHome ? homeWon : !homeWon;

    ourPlayers.forEach(p => addStat(p, isDouble, won));

    if (f.homeWin !== undefined) {  // Assume only count decided frames
      if (homeWon) homeFrames++;
    }
  });

  let decidedFrames = matchJson.data.filter(f => f.homeWin !== undefined).length;
  let awayFrames = decidedFrames - homeFrames;

  const weHome = isHome;
  const ourScore = weHome ? homeFrames : awayFrames;
  const oppScore = weHome ? awayFrames : homeFrames;
  const res = ourScore > oppScore ? ' W' : ourScore === oppScore ? ' D' : ' L';

  let matchStr;
  if (decidedFrames > 0) {
    matchStr = `Match ID: ${matchId} • ${opponent} (${ourScore}-${oppScore})${res}`;
  } else {
    matchStr = `Match ID: ${matchId} • vs ${opponent}`;
  }

  // Build rows
  const rows = Object.entries(stats).map(([name, s]) => {
    const totalGames = s.singlesPlayed + s.doublesPlayed;
    const singlesPct = s.singlesPlayed ? s.singlesWon / s.singlesPlayed : 0;
    const doublesPct = s.doublesPlayed ? s.doublesWon / s.doublesPlayed : 0;
    const ppg = totalGames ? s.points / totalGames : 0;

    return [
      name,
      totalGames,
      s.singlesPlayed,
      s.singlesWon,
      singlesPct,
      s.doublesPlayed,
      s.doublesWon,
      doublesPct,
      s.points,
      ppg
    ];
  });

  // Sort by PPG, then points, then games, then name
  rows.sort((a, b) => {
    if (b[9] !== a[9]) return b[9] - a[9];
    if (b[8] !== a[8]) return b[8] - a[8];
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  });

  const totalCols = 10;
  const thaiDate = matchDate.toLocaleDateString('en-CA', { timeZone: 'Asia/Bangkok' });
  const now = new Date();
  const nowStr = now.toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' });

  // === PERFECT HEADER BLOCK WITH SIMILAR COLORS ===
  out.getRange(1, 1, 1, totalCols).merge().setValue(
    'Match Form – ' + (isHome ? teamName + ' vs ' + opponent : opponent + ' vs ' + teamName)
  ).setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('white')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(2, 1, 1, totalCols).merge().setValue(`Match Date: ${thaiDate} (Thai Time)`)
    .setFontStyle('italic').setFontColor('#666').setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(3, 1, 1, totalCols).merge().setValue(`Last Refresh: ${nowStr}`)
    .setFontStyle('italic').setFontColor('#666').setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(4, 1, 1, totalCols).merge().setValue(matchStr)
    .setFontWeight('bold').setFontColor('#333').setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Column headers row 6 (skip row 5 like in Last3Form)
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

  // Data rows from row 7
  if (rows.length) {
    out.getRange(7, 1, rows.length, totalCols).setValues(rows);

    out.getRange(7, 5, rows.length, 1).setNumberFormat('0.0%');
    out.getRange(7, 8, rows.length, 1).setNumberFormat('0.0%');
    out.getRange(7, 10, rows.length, 1).setNumberFormat('0.00');

    out.getRange(7, 2, rows.length, 9).setHorizontalAlignment('center');
  } else {
    out.getRange(7, 1).setValue('No player stats available');
  }

  for (let i = 1; i <= totalCols; i++) out.setColumnWidth(i, 100);
  const lastRow = rows.length ? rows.length + 6 : 8;
  out.getRange(1, 1, lastRow, totalCols).setBorder(true, true, true, true, true, true);
}

function handleCurrentMatchFormV2Edit_(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();

  Logger.log('Edited sheet: ' + sheet.getName() + ', range: ' + range.getA1Notation());

  if (sheet.getName() === 'CurrentMatchFormV2' &&
      range.getRow() === 18 &&
      range.getColumn() === 1) {

    buildCurrentMatchFormSmartV2();

    // Clear the trigger cell text only
    sheet.getRange('A18').clearContent();
  }
}