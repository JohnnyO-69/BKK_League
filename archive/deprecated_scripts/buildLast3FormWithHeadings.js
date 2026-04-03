/*************************************************
 *  LAST 3 MATCH FORM WITH HEADINGS (GENERIC)
 *************************************************/
function buildLast3FormWithHeadings(teamId, teamName, outputSheetName) {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  if (!fixturesSheet) return;

  let out = ss.getSheetByName(outputSheetName);
  if (!out) out = ss.insertSheet(outputSheetName);

  out.getDataRange().clearContent();

  const delayMs = 400;
  const now = new Date();
  const sixWeeksMs = 42 * 24 * 60 * 60 * 1000;

  const today = new Date(
    new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' })
  );

  const values = fixturesSheet.getDataRange().getValues();
  const header = values.shift();

  const idxMatchId    = header.indexOf('Match ID');
  const idxHome       = header.indexOf('Home Team');
  const idxAway       = header.indexOf('Away Team');
  const idxDate       = header.indexOf('Match Date');
  const idxHomeFrames = header.indexOf('Home Frames');
  const idxAwayFrames = header.indexOf('Away Frames');

  function toDate(v) {
    if (v instanceof Date) return v;
    const d = new Date(String(v || '').trim());
    return isNaN(d.getTime()) ? new Date(0) : d;
  }

  const matchesRaw = values
    .filter(r => (r[idxHome] === teamName || r[idxAway] === teamName) && r[idxMatchId])
    .map(r => ({
      id: r[idxMatchId],
      date: toDate(r[idxDate]),
      home: r[idxHome],
      away: r[idxAway],
      homeFrames: idxHomeFrames >= 0 ? r[idxHomeFrames] : '',
      awayFrames: idxAwayFrames >= 0 ? r[idxAwayFrames] : ''
    }))
    .filter(m => !isNaN(m.date.getTime()));

  if (!matchesRaw.length) {
    out.getRange('A1').setValue('No matches found for ' + teamName);
    return;
  }

  const completed = matchesRaw
    .filter(m => m.date <= today)
    .filter(m =>
      m.homeFrames !== '' &&
      m.awayFrames !== ''
    )
    .sort((a, b) => a.date - b.date);

  if (!completed.length) {
    out.getRange('A1').setValue('No completed matches with scores for ' + teamName);
    return;
  }

  const mostRecentMatch = completed[completed.length - 1];
  const last3Matches = completed.slice(-3);

  const playerMatches = {};

  function fetchJson(url) {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
    } catch {
      return null;
    }
  }

  function addGame(player, matchId, matchDate, isDouble, won) {
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
  }

  completed.forEach(m => {
    const json = fetchJson('https://api.bkkleague.com/match/details/' + m.id);
    if (!json || !Array.isArray(json.data)) return;

    json.data.forEach(f => {
      const hp = (f.homePlayers || []).map(p => p.nickName);
      const ap = (f.awayPlayers || []).map(p => p.nickName);
      const isDouble = hp.length > 1;
      const homeWon = f.homeWin === 1;

      const homeId = f.homeTeamId || f.home_team_id || f.homeTeamid;
      const awayId = f.awayTeamId || f.away_team_id || f.awayTeamid;

      const weHome = homeId === teamId;
      const weAway = awayId === teamId;

      if (!weHome && !weAway) return;

      const ourPlayers = weHome ? hp : ap;
      const won = weHome ? homeWon : !homeWon;

      ourPlayers.forEach(p => addGame(p, m.id, m.date, isDouble, won));
    });

    Utilities.sleep(delayMs);
  });

  const rowsTemp = Object.entries(playerMatches).map(([player, statsObj]) => {
    const games = Object.values(statsObj).sort((a, b) => b.date - a.date);
    const last3 = games.slice(0, 3);

    const singlesPlayed = sum(last3, 'singlesPlayed');
    const singlesWon = sum(last3, 'singlesWon');
    const doublesPlayed = sum(last3, 'doublesPlayed');
    const doublesWon = sum(last3, 'doublesWon');
    const points = sum(last3, 'points');
    const totalGames = singlesPlayed + doublesPlayed;

    const lastPlayed = last3.length ? last3[0].date : new Date(0);
    const active = now - lastPlayed <= sixWeeksMs;

    const singlesPct = singlesPlayed ? singlesWon / singlesPlayed : 0;
    const doublesPct = doublesPlayed ? doublesWon / doublesPlayed : 0;
    const ppg = totalGames ? points / totalGames : 0;

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
      active
    };
  });

  const rows = rowsTemp
    .filter(r => r.active)
    .sort((a, b) => {
      if (b.ppg !== a.ppg) return b.ppg - a.ppg;
      if (b.points !== a.points) return b.points - a.points;
      if (b.totalGames !== a.totalGames) return b.totalGames - a.totalGames;
      return a.player.localeCompare(b.player);
    })
    .map(r => [
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

  /*************************************************
   *      HEADING BLOCK
   *************************************************/
  out.getRange(1, 1).setValue('Last 3 Match Form – ' + teamName);
  out.getRange(2, 1).setValue(
    'Most Recent Match Date: ' + mostRecentMatch.date.toISOString().split('T')[0]
  );
  out.getRange(3, 1).setValue('Matches Considered: Last 3 Completed Matches');

  out.getRange(1, 1, 1, 10).mergeAcross().setFontWeight('bold')
     .setHorizontalAlignment('center').setFontSize(12);

  out.getRange(2, 1, 1, 10).mergeAcross().setFontWeight('bold')
     .setHorizontalAlignment('center');

  out.getRange(3, 1, 1, 10).mergeAcross().setFontWeight('bold')
     .setHorizontalAlignment('center');

  // Last refresh
  const refreshedAt = new Date(
    new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' })
  );

  out.getRange(4, 1).setValue(
    'Last Refresh: ' +
      refreshedAt.toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' })
  );
  out.getRange(4, 1, 1, 10).mergeAcross()
     .setHorizontalAlignment('center')
     .setFontStyle('italic');

  /*************************************************
   *      LAST 3 MATCH RESULTS (W L D)
   *************************************************/
  const last3Summary = last3Matches.map(m => {
    const isHome = m.home === teamName;
    const opp = isHome ? m.away : m.home;
    const forPts = isHome ? m.homeFrames : m.awayFrames;
    const agPts = isHome ? m.awayFrames : m.homeFrames;

    let r;
    if (forPts > agPts) r = 'W';
    else if (forPts < agPts) r = 'L';
    else r = 'D';

    return `${opp} (${forPts}-${agPts}) ${r}`;
  });

  out.getRange(5, 1).setValue('Last 3 Matches:   ' + last3Summary.join('     '));
  out.getRange(5, 1, 1, 10).mergeAcross()
     .setFontWeight('bold')
     .setHorizontalAlignment('center');

  /*************************************************
   *      PLAYER TABLE HEADER
   *************************************************/
  const headerRow = [
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

  out.getRange(6, 1, 1, 10).setValues([headerRow])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);

  /*************************************************
   *      PLAYER DATA ROWS
   *************************************************/
  if (rows.length) {
    out.getRange(7, 1, rows.length, 10).setValues(rows);

    out.getRange(7, 5, rows.length, 1).setNumberFormat('0.0%');
    out.getRange(7, 8, rows.length, 1).setNumberFormat('0.0%');
    out.getRange(7, 10, rows.length, 1).setNumberFormat('0.00');

    out.getRange(7, 2, rows.length, 9).setHorizontalAlignment('center');
  } else {
    out.getRange(7, 1).setValue('No active players in last 6 weeks');
  }

  const lastRow = rows.length ? rows.length + 6 : 7;
  out.getRange(1, 1, lastRow, 10).setBorder(true, true, true, true, true, true);
}

function sum(arr, key) {
  return arr.reduce((t, x) => t + (x[key] || 0), 0);
}

/*************************************************
 *      FIND NEXT OPPONENT FOR THE GAME 8B
 *************************************************/
function buildNextTeam3MatchForm() {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  if (!fixturesSheet) return;

  const values = fixturesSheet.getDataRange().getValues();
  const header = values.shift();

  const idxDate   = header.indexOf('Match Date');
  const idxHome   = header.indexOf('Home Team');
  const idxAway   = header.indexOf('Away Team');
  const idxHomeId = header.indexOf('Home Team ID');
  const idxAwayId = header.indexOf('Away Team ID');

  const team = 'The Game 8B';

  const today = new Date(
    new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' })
  );

  const upcoming = values
    .map(r => ({
      date: new Date(r[idxDate]),
      home: r[idxHome],
      away: r[idxAway],
      homeId: r[idxHomeId],
      awayId: r[idxAwayId]
    }))
    .filter(m => !isNaN(m.date.getTime()))
    .filter(m => m.date >= today)
    .filter(m => m.home === team || m.away === team)
    .sort((a, b) => a.date - b.date);

  const nextMatch = upcoming[0];
  if (!nextMatch) {
    SpreadsheetApp.getUi().alert('No upcoming matches found for The Game 8B');
    return;
  }

  const oppName = nextMatch.home === team ? nextMatch.away : nextMatch.home;
  const oppId   = nextMatch.home === team ? nextMatch.awayId : nextMatch.homeId;

  buildLast3FormWithHeadings(oppId, oppName, 'NextTeam3MatchForm');
}
