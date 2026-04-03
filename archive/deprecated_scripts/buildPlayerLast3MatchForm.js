function buildPlayerLast3MatchForm() {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  const paramsSheet = ss.getSheetByName('Parameters');
  if (!fixturesSheet) return;
  if (!paramsSheet) return;

  const outName = 'PlayerLast3MatchForm';
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear();

  const teamName = String(paramsSheet.getRange('F7').getValue() || '').trim();
  const teamId = Number(paramsSheet.getRange('F8').getValue());
  const delayMs = 400;

  if (!teamName || Number.isNaN(teamId)) {
    out.getRange('A1').setValue('Missing team name/id in Parameters!F7:F8');
    return;
  }

  const now = new Date();
  const sixWeeksMs = 42 * 24 * 60 * 60 * 1000;

  const values = fixturesSheet.getDataRange().getValues();
  const header = values.shift();
  const idxMatchId = header.indexOf('Match ID');
  const idxHome = header.indexOf('Home Team');
  const idxAway = header.indexOf('Away Team');
  const idxDate = header.indexOf('Match Date');

  if ([idxMatchId, idxHome, idxAway, idxDate].some(i => i < 0)) return;

  const toDate = v => {
    if (v instanceof Date) return v;
    const d = new Date(String(v || '').trim());
    return isNaN(d) ? new Date(0) : d;
  };

  const matches = values
    .filter(r => (r[idxHome] === teamName || r[idxAway] === teamName) && r[idxMatchId])
    .map(r => ({ id: r[idxMatchId], date: toDate(r[idxDate]) }))
    .sort((a, b) => a.date - b.date);

  if (!matches.length) {
    out.getRange('A1').setValue(`No matches found for ${teamName}`);
    return;
  }

  const playerMatches = {};

  const fetchJson = url => {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
    } catch {
      return null;
    }
  };

  const addGame = (player, matchId, matchDate, isDouble, won) => {
    if (!player) return;
    if (!playerMatches[player]) playerMatches[player] = {};

    if (!playerMatches[player][matchId]) {
      playerMatches[player][matchId] = {
        date: matchDate,
        singlesPlayed: 0,
        singlesWon: 0,
        doublesPlayed: 0,
        doublesWon: 0,
        points: 0
      };
    }

    const x = playerMatches[player][matchId];

    if (isDouble) {
      x.doublesPlayed++;
      if (won) {
        x.doublesWon++;
        x.points += 0.5;
      }
    } else {
      x.singlesPlayed++;
      if (won) {
        x.singlesWon++;
        x.points += 1;
      }
    }
  };

  matches.forEach(m => {
    const json = fetchJson(`https://api.bkkleague.com/match/details/${m.id}`);
    if (!json || !Array.isArray(json.data)) return;

    json.data.forEach(f => {
      const hp = (f.homePlayers || []).map(p => p.nickName);
      const ap = (f.awayPlayers || []).map(p => p.nickName);
      const isDouble = hp.length > 1;

      const homeWon = f.homeWin === 1;
      const homeId = f.homeTeamId || f.home_team_id || f.homeTeamid;
      const awayId = f.awayTeamId || f.away_team_id || f.awayTeamid;

      const weAreHome = homeId === teamId;
      const weAreAway = awayId === teamId;
      if (!weAreHome && !weAreAway) return;

      const ourPlayers = weAreHome ? hp : ap;
      const won = weAreHome ? homeWon : !homeWon;

      ourPlayers.forEach(p => addGame(p, m.id, m.date, isDouble, won));
    });

    Utilities.sleep(delayMs);
  });

  const rows = Object.entries(playerMatches).map(([player, matchStats]) => {
    const games = Object.values(matchStats).sort((a, b) => b.date - a.date);
    const last3 = games.slice(0, 3);

    const singlesPlayed = sum(last3, 'singlesPlayed');
    const singlesWon = sum(last3, 'singlesWon');
    const doublesPlayed = sum(last3, 'doublesPlayed');
    const doublesWon = sum(last3, 'doublesWon');
    const points = sum(last3, 'points');
    const totalGames = singlesPlayed + doublesPlayed;

    const ppg = totalGames ? points / totalGames : 0;

    const lastPlayed = last3.length ? last3[0].date : new Date(0);
    const active = now - lastPlayed <= sixWeeksMs;

    const singlesPct = singlesPlayed ? singlesWon / singlesPlayed : 0;
    const doublesPct = doublesPlayed ? doublesWon / doublesPlayed : 0;

    return {
      player,
      totalGames,
      singlesPlayed,
      singlesWon,
      singlesPct,
      doublesPlayed,
      doublesWon,
      doublesPct,
      points,
      ppg,
      lastPlayed,
      active
    };
  });

  const activeRows = rows.filter(r => r.active);

  activeRows.sort((a, b) => {
    if (b.ppg !== a.ppg) return b.ppg - a.ppg;
    if (b.points !== a.points) return b.points - a.points;
    if (b.totalGames !== a.totalGames) return b.totalGames - a.totalGames;
    return a.player.localeCompare(b.player);
  });

  const outRows = activeRows.map(r => [
    r.player,
    r.totalGames,
    r.singlesPlayed,
    r.singlesWon,
    r.singlesPct,
    r.doublesPlayed,
    r.doublesWon,
    r.doublesPct,
    r.points,
    r.ppg
  ]);

  const outHeaders = [
    'Player',
    'Games Played (Last 3 Matches)',
    'Singles Played',
    'Singles Won',
    'Singles Win %',
    'Doubles Played',
    'Doubles Won',
    'Doubles Win %',
    'Points',
    'Points Per Game'
  ];

  // Write header
  out.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);

  // Format header
  out.getRange(1, 1, 1, outHeaders.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (outRows.length) {

    // Clear formatting to avoid typed columns
    out.getRange(2, 1, outRows.length, outHeaders.length).clearFormat();

    // Write rows
    out.getRange(2, 1, outRows.length, outHeaders.length).setValues(outRows);

    const startRow = 2;

    // Apply % formatting
    out.getRange(startRow, 5, outRows.length, 1).setNumberFormat("0.0%");
    out.getRange(startRow, 8, outRows.length, 1).setNumberFormat("0.0%");

    // Format PPG
    out.getRange(startRow, 10, outRows.length, 1).setNumberFormat("0.00");

    // Center align B-J
    out.getRange(startRow, 2, outRows.length, 9).setHorizontalAlignment("center");

  } else {
    out.getRange(2, 1).setValue("No active players in last 6 weeks");
  }

  // Set column widths
  for (let col = 1; col <= outHeaders.length; col++) {
    out.setColumnWidth(col, 100);
  }

  // Border around data
  const fullRange = out.getRange(1, 1, activeRows.length + 1, outHeaders.length);
  fullRange.setBorder(true, true, true, true, true, true);

}

function sum(arr, key) {
  return arr.reduce((t, x) => t + (x[key] || 0), 0);
}
