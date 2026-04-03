function buildCurrentMatchFormSmart() {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  if (!fixturesSheet) return;

  const outName = 'CurrentMatchForm';
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);

  // Only clear values, preserve formatting
  out.getDataRange().clearContent();

  const teamName = 'The Game 8B';
  const teamId = 2327;

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
    out.getRange('A1').setValue('No matches found for The Game 8B');
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

  // Process frames
  matchJson.data.forEach(f => {
    const hp = (f.homePlayers || []).map(p => p.nickName);
    const ap = (f.awayPlayers || []).map(p => p.nickName);

    const isDouble = hp.length > 1;
    const homeWon = f.homeWin === 1;

    const homeId = f.homeTeamId || f.home_team_id || f.homeTeamid;
    const awayId = f.awayTeamId || f.away_team_id || f.awayTeamid;

    const ourHome = homeId === teamId;
    const ourAway = awayId === teamId;

    if (!ourHome && !ourAway) return;

    const ourPlayers = ourHome ? hp : ap;
    const won = ourHome ? homeWon : !homeWon;

    ourPlayers.forEach(p => addStat(p, isDouble, won));
  });

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

  // Headers
  out.getRange(1, 1).setValue(
    'Current Match Form – ' + (isHome ? teamName + ' vs ' + opponent : opponent + ' vs ' + teamName)
  );
  out.getRange(2, 1).setValue('Date: ' + matchDate.toISOString().split('T')[0]);
  out.getRange(3, 1).setValue('Match ID: ' + matchId);

  out.getRange(1, 1, 1, 10).mergeAcross();
  out.getRange(2, 1, 1, 10).mergeAcross();
  out.getRange(3, 1, 1, 10).mergeAcross();

  out.getRange(1, 1).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);
  out.getRange(2, 1).setFontWeight('bold').setHorizontalAlignment('center');
  out.getRange(3, 1).setFontWeight('bold').setHorizontalAlignment('center');

  // Last refresh line in row 4
  const refreshedAt = new Date(
    new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' })
  );
  out.getRange(4, 1).setValue('Last refresh: ' + refreshedAt.toLocaleString('en-GB', {
    timeZone: 'Asia/Bangkok'
  }));
  out.getRange(4, 1, 1, 10).mergeAcross();
  out.getRange(4, 1).setHorizontalAlignment('center').setFontStyle('italic');

  // Column headers row 5
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

  out.getRange(5, 1, 1, 10)
    .setValues([colHeaders])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);

  // Data rows from row 6
  if (rows.length) {
    out.getRange(6, 1, rows.length, 10).setValues(rows);

    out.getRange(6, 5, rows.length, 1).setNumberFormat('0.0%');
    out.getRange(6, 8, rows.length, 1).setNumberFormat('0.0%');
    out.getRange(6, 10, rows.length, 1).setNumberFormat('0.00');

    out.getRange(6, 2, rows.length, 9).setHorizontalAlignment('center');
  } else {
    out.getRange(6, 1).setValue('No player stats available');
  }

  const lastRow = rows.length ? rows.length + 5 : 6;
  out.getRange(1, 1, lastRow, 10).setBorder(true, true, true, true, true, true);
}

// onEdit trigger for A18 in CurrentMatchForm
function onEdit(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();

  // Debug log so you can see what is being edited if needed
  Logger.log('Edited sheet: ' + sheet.getName() + ', range: ' + range.getA1Notation());

  if (sheet.getName() === 'CurrentMatchForm' &&
      range.getRow() === 18 &&
      range.getColumn() === 1) {

    buildCurrentMatchFormSmart();

    // Clear the trigger cell text only
    sheet.getRange('A18').clearContent();
  }
}
