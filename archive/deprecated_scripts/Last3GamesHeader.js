/*************************************************
 *  COMMON UTILS
 *************************************************/
function sum(arr, key) {
  return arr.reduce((t, x) => t + (x[key] || 0), 0);
}

function toThaiToday() {
  return new Date(
    new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' })
  );
}

function toDateSafe(v) {
  if (v instanceof Date) return v;
  const d = new Date(String(v || '').trim());
  return isNaN(d.getTime()) ? new Date(0) : d;
}

function fetchJson(url) {
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return res.getResponseCode() === 200
      ? JSON.parse(res.getContentText())
      : null;
  } catch (err) {
    Logger.log('Fetch error: ' + err);
    return null;
  }
}

/*************************************************
 *  REUSABLE: LAST 3 MATCH FORM WITH HEADINGS
 *  - teamId: numeric ID (e.g. 2327)
 *  - teamName: string (e.g. "The Game 8B")
 *  - outputSheetName: target sheet name
 *************************************************/
function buildLast3FormWithHeadings(teamId, teamName, outputSheetName) {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  if (!fixturesSheet) return;

  let out = ss.getSheetByName(outputSheetName);
  if (!out) out = ss.insertSheet(outputSheetName);

  // Keep formatting, clear values
  out.getDataRange().clearContent();

  const delayMs = 400;
  const now = toThaiToday();
  const sixWeeksMs = 42 * 24 * 60 * 60 * 1000;

  const values = fixturesSheet.getDataRange().getValues();
  const header = values.shift();

  const idxMatchId    = header.indexOf('Match ID');
  const idxHome       = header.indexOf('Home Team');
  const idxAway       = header.indexOf('Away Team');
  const idxDate       = header.indexOf('Match Date');
  const idxHomeFrames = header.indexOf('Home Frames');
  const idxAwayFrames = header.indexOf('Away Frames');

  if ([idxMatchId, idxHome, idxAway, idxDate].some(i => i < 0)) {
    out.getRange('A1').setValue('Fixtures sheet missing required columns');
    return;
  }

  const matchesRaw = values
    .filter(r => (r[idxHome] === teamName || r[idxAway] === teamName) && r[idxMatchId])
    .map(r => ({
      id: r[idxMatchId],
      date: toDateSafe(r[idxDate]),
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

  // Only completed matches (date <= today and scores present)
  const today = toThaiToday();

  const completed = matchesRaw
    .filter(m => m.date <= today)
    .filter(m =>
      m.homeFrames !== '' && m.homeFrames != null &&
      m.awayFrames !== '' && m.awayFrames != null
    )
    .sort((a, b) => a.date - b.date);

  if (!completed.length) {
    out.getRange('A1').setValue('No completed matches with scores for ' + teamName);
    return;
  }

  const mostRecentMatch = completed[completed.length - 1];

  // For W/L/D strip which side we were on
  function resultForMatch(m) {
    const home = m.home;
    const away = m.away;
    const hf = Number(m.homeFrames);
    const af = Number(m.awayFrames);

    const weHome = home === teamName;
    const weAway = away === teamName;

    let res = 'D';
    if (weHome) {
      if (hf > af) res = 'W';
      else if (hf < af) res = 'L';
    } else if (weAway) {
      if (af > hf) res = 'W';
      else if (af < hf) res = 'L';
    }

    const opp = weHome ? away : home;
    const score = weHome ? hf + ' - ' + af : af + ' - ' + hf;
    return { opp, score, res };
  }

  // Collect last 3 completed matches results (for summary row)
  const last3Completed = completed.slice(-3).map(resultForMatch);

  // Player per-match stats
  const playerMatches = {};

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

      ourPlayers.forEach(player => {
        if (!player) return;

        if (!playerMatches[player]) playerMatches[player] = {};

        if (!playerMatches[player][m.id]) {
          playerMatches[player][m.id] = {
            date: m.date,
            singlesPlayed: 0,
            singlesWon: 0,
            doublesPlayed: 0,
            doublesWon: 0,
            points: 0
          };
        }

        const rec = playerMatches[player][m.id];

        if (isDouble) {
          rec.doublesPlayed++;
          if (won) {
            rec.doublesWon++;
            rec.points += 0.5;
          }
        } else {
          rec.singlesPlayed++;
          if (won) {
            rec.singlesWon++;
            rec.points += 1;
          }
        }
      });
    });

    Utilities.sleep(delayMs);
  });

  // Per player, last 3 matches only
  const rowsObj = Object.entries(playerMatches).map(([player, matchStats]) => {
    const matchesArr = Object.values(matchStats).sort((a, b) => b.date - a.date);
    const last3 = matchesArr.slice(0, 3);

    const singlesPlayed = sum(last3, 'singlesPlayed');
    const singlesWon = sum(last3, 'singlesWon');
    const doublesPlayed = sum(last3, 'doublesPlayed');
    const doublesWon = sum(last3, 'doublesWon');
    const points = sum(last3, 'points');
    const matchesPlayed = last3.length; // distinct matches in last 3
    const totalGames = singlesPlayed + doublesPlayed;

    const lastPlayed = last3.length ? last3[0].date : new Date(0);
    const active = now - lastPlayed <= sixWeeksMs;

    const singlesPct = singlesPlayed ? singlesWon / singlesPlayed : 0;
    const doublesPct = doublesPlayed ? doublesWon / doublesPlayed : 0;
    const ppg = totalGames ? points / totalGames : 0;

    return {
      player,
      matchesPlayed,
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

  const rows = rowsObj
    .filter(r => r.active)
    .sort((a, b) => {
      if (b.ppg !== a.ppg) return b.ppg - a.ppg;
      if (b.points !== a.points) return b.points - a.points;
      if (b.totalGames !== a.totalGames) return b.totalGames - a.totalGames;
      return a.player.localeCompare(b.player);
    })
    .map(r => [
      r.player,
      r.matchesPlayed,
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

  // Heading block
  out.getRange(1, 1).setValue('Last 3 Match Form – ' + teamName);
  out.getRange(2, 1).setValue(
    'Most Recent Match Date: ' + mostRecentMatch.date.toISOString().split('T')[0]
  );
  out.getRange(3, 1).setValue('Matches Considered: Last 3 Completed Matches');

  out.getRange(1, 1, 1, 12).mergeAcross()
    .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);
  out.getRange(2, 1, 1, 12).mergeAcross()
    .setFontWeight('bold').setHorizontalAlignment('center');
  out.getRange(3, 1, 1, 12).mergeAcross()
    .setFontWeight('bold').setHorizontalAlignment('center');

  // Last refresh
  const refreshedAt = toThaiToday();
  out.getRange(4, 1).setValue(
    'Last Refresh: ' +
    refreshedAt.toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' })
  );
  out.getRange(4, 1, 1, 12).mergeAcross()
    .setHorizontalAlignment('center').setFontStyle('italic');

  // Last 3 match W/L/D string
  const last3Str = last3Completed
    .map(m => m.opp + ' ' + m.score + ' (' + m.res + ')')
    .join('  |  ');

  out.getRange(5, 1).setValue('Last 3 match results: ' + last3Str);
  out.getRange(5, 1, 1, 12).mergeAcross()
    .setHorizontalAlignment('center');

  // Column headers
  const headerRow = [
    'Player',
    'Matches Played',
    'Games Played',
    'Singles Played',
    'Singles Won',
    'Singles Win %',
    'Doubles Played',
    'Doubles Won',
    'Doubles Win %',
    'Points',
    'Points Per Game',
    '' // keep width consistent if you want an extra buffer col
  ];

  out.getRange(6, 1, 1, 12).setValues([headerRow])
    .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);

  // Data
  if (rows.length) {
    out.getRange(7, 1, rows.length, 11).setValues(rows);

    // Percent and PPG formatting
    out.getRange(7, 6, rows.length, 1).setNumberFormat('0.0%');  // Singles %
    out.getRange(7, 9, rows.length, 1).setNumberFormat('0.0%');  // Doubles %
    out.getRange(7, 11, rows.length, 1).setNumberFormat('0.00'); // PPG

    // Center numeric columns (B → K)
    out.getRange(7, 2, rows.length, 10).setHorizontalAlignment('center');
  } else {
    out.getRange(7, 1).setValue('No active players in last 6 weeks');
  }

  const lastRow = rows.length ? rows.length + 6 : 7;
  out.getRange(1, 1, lastRow, 11).setBorder(true, true, true, true, true, true);

  // Optional: width
  for (let col = 1; col <= 11; col++) {
    out.setColumnWidth(col, 100);
  }
}

/*************************************************
 *  REUSABLE: SEASON FORM WITH HEADINGS
 *  - teamId, teamName, outputSheetName
 *  - counts matchesPlayed per player
 *************************************************/
function buildSeasonFormWithHeadings(teamId, teamName, outputSheetName) {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  if (!fixturesSheet) return;

  let sheet = ss.getSheetByName(outputSheetName);
  if (!sheet) sheet = ss.insertSheet(outputSheetName);
  sheet.getDataRange().clearContent();

  const delayMs = 400;
  const now = toThaiToday();
  const sixWeeksMs = 42 * 24 * 60 * 60 * 1000;
  const today = toThaiToday();

  const fixtureData = fixturesSheet.getDataRange().getValues();
  const header = fixtureData.shift();

  const idxId         = header.indexOf('Match ID');
  const idxHome       = header.indexOf('Home Team');
  const idxAway       = header.indexOf('Away Team');
  const idxDate       = header.indexOf('Match Date');
  const idxHomeFrames = header.indexOf('Home Frames');
  const idxAwayFrames = header.indexOf('Away Frames');

  if ([idxId, idxHome, idxAway, idxDate].some(i => i < 0)) {
    sheet.getRange('A1').setValue('Fixtures sheet missing required columns');
    return;
  }

  const matches = fixtureData
    .filter(r => (r[idxHome] === teamName || r[idxAway] === teamName) && r[idxId])
    .map(r => ({
      id: r[idxId],
      date: toDateSafe(r[idxDate]),
      home: r[idxHome],
      away: r[idxAway],
      homeFrames: idxHomeFrames >= 0 ? r[idxHomeFrames] : '',
      awayFrames: idxAwayFrames >= 0 ? r[idxAwayFrames] : ''
    }))
    .filter(m => !isNaN(m.date.getTime()))
    .filter(m =>
      m.date <= today &&
      m.homeFrames !== '' && m.homeFrames != null &&
      m.awayFrames !== '' && m.awayFrames != null
    )
    .sort((a, b) => a.date - b.date);

  if (!matches.length) {
    sheet.getRange('A1').setValue('No completed matches with scores for ' + teamName);
    return;
  }

  const playerStats = {};
  const teamMatchesPlayed = matches.length;
  const lastTeamMatch = matches[matches.length - 1];

  function record(player, matchId, matchDate, won, isDouble) {
    if (!player) return;

    if (!playerStats[player]) {
      playerStats[player] = {
        matches: {},
        singlesPlayed: 0,
        singlesWon: 0,
        doublesPlayed: 0,
        doublesWon: 0,
        points: 0,
        lastPlayed: new Date(0)
      };
    }

    const s = playerStats[player];

    // Track match participation
    if (!s.matches[matchId]) {
      s.matches[matchId] = true;
    }

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

    if (matchDate > s.lastPlayed) s.lastPlayed = matchDate;
  }

  matches.forEach(m => {
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

      ourPlayers.forEach(p => record(p, m.id, m.date, won, isDouble));
    });

    Utilities.sleep(delayMs);
  });

  const rowsRaw = Object.entries(playerStats)
    .filter(([, s]) => now - s.lastPlayed <= sixWeeksMs)
    .map(([player, s]) => {
      const matchesPlayed = Object.keys(s.matches).length;
      const totalGames = s.singlesPlayed + s.doublesPlayed;
      const ppg = totalGames ? s.points / totalGames : 0;
      const singlesPct = s.singlesPlayed ? s.singlesWon / s.singlesPlayed : 0;
      const doublesPct = s.doublesPlayed ? s.doublesWon / s.doublesPlayed : 0;

      return {
        player,
        matchesPlayed,
        totalGames,
        singlesPlayed: s.singlesPlayed,
        singlesWon: s.singlesWon,
        singlesPct,
        doublesPlayed: s.doublesPlayed,
        doublesWon: s.doublesWon,
        doublesPct,
        points: s.points,
        ppg
      };
    });

  const rows = rowsRaw
    .sort((a, b) => {
      if (b.ppg !== a.ppg) return b.ppg - a.ppg;
      if (b.points !== a.points) return b.points - a.points;
      if (b.totalGames !== a.totalGames) return b.totalGames - a.totalGames;
      return a.player.localeCompare(b.player);
    })
    .map(r => [
      r.player,
      r.matchesPlayed,
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

  // Heading
  sheet.getRange(1, 1).setValue('Season Form – ' + teamName);
  sheet.getRange(2, 1).setValue('Team Matches Played: ' + teamMatchesPlayed);
  sheet.getRange(3, 1).setValue(
    'Most Recent Match Date: ' + lastTeamMatch.date.toISOString().split('T')[0]
  );

  sheet.getRange(1, 1, 1, 12).mergeAcross()
    .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);
  sheet.getRange(2, 1, 1, 12).mergeAcross()
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(3, 1, 1, 12).mergeAcross()
    .setFontWeight('bold').setHorizontalAlignment('center');

  const refreshedAt = toThaiToday();
  sheet.getRange(4, 1).setValue(
    'Last Refresh: ' +
    refreshedAt.toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' })
  );
  sheet.getRange(4, 1, 1, 12).mergeAcross()
    .setHorizontalAlignment('center').setFontStyle('italic');

  const headers = [
    'Player',
    'Matches Played',
    'Games Played',
    'Singles Played',
    'Singles Won',
    'Singles Win %',
    'Doubles Played',
    'Doubles Won',
    'Doubles Win %',
    'Points',
    'Points Per Game',
    ''
  ];

  sheet.getRange(5, 1, 1, 12).setValues([headers])
    .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);

  if (rows.length) {
    sheet.getRange(6, 1, rows.length, 11).clearFormat();
    sheet.getRange(6, 1, rows.length, 11).setValues(rows);

    sheet.getRange(6, 6, rows.length, 1).setNumberFormat('0.0%');  // Singles %
    sheet.getRange(6, 9, rows.length, 1).setNumberFormat('0.0%');  // Doubles %
    sheet.getRange(6, 11, rows.length, 1).setNumberFormat('0.00'); // PPG

    sheet.getRange(6, 2, rows.length, 10).setHorizontalAlignment('center');
  } else {
    sheet.getRange(6, 1).setValue('No active players in last 6 weeks');
  }

  const lastRow = rows.length ? rows.length + 5 : 6;
  sheet.getRange(1, 1, lastRow, 11).setBorder(true, true, true, true, true, true);

  for (let col = 1; col <= 11; col++) {
    sheet.setColumnWidth(col, 100);
  }
}

/*************************************************
 *  WRAPPERS
 *************************************************/

// The Game 8B – season form
function buildTheGame8BSeasonForm() {
  const teamId = 2327;
  const teamName = 'The Game 8B';
  buildSeasonFormWithHeadings(teamId, teamName, 'TheGame8BForm');
}

// Next opponent – season form
function buildNextTeamSeasonForm() {
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

  const myTeam = 'The Game 8B';
  const today = toThaiToday();

  const upcoming = values
    .map(r => ({
      date: toDateSafe(r[idxDate]),
      home: r[idxHome],
      away: r[idxAway],
      homeId: r[idxHomeId],
      awayId: r[idxAwayId]
    }))
    .filter(m => !isNaN(m.date.getTime()))
    .filter(m => m.date >= today)
    .filter(m => m.home === myTeam || m.away === myTeam)
    .sort((a, b) => a.date - b.date);

  const nextMatch = upcoming[0];
  if (!nextMatch) {
    SpreadsheetApp.getUi().alert('No upcoming matches found for The Game 8B');
    return;
  }

  const oppName = nextMatch.home === myTeam ? nextMatch.away : nextMatch.home;
  const oppId   = nextMatch.home === myTeam ? nextMatch.awayId : nextMatch.homeId;

  buildSeasonFormWithHeadings(oppId, oppName, 'NextTeamSeasonForm');
}

// The Game 8B – last 3 match form (you already have)
function buildTheGameLast3MatchForm() {
  const teamId = 2327;
  const teamName = 'The Game 8B';
  buildLast3FormWithHeadings(teamId, teamName, 'PlayerLast3MatchForm');
}

// Next opponent – last 3 match form
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

  const myTeam = 'The Game 8B';
  const today = toThaiToday();

  const upcoming = values
    .map(r => ({
      date: toDateSafe(r[idxDate]),
      home: r[idxHome],
      away: r[idxAway],
      homeId: r[idxHomeId],
      awayId: r[idxAwayId]
    }))
    .filter(m => !isNaN(m.date.getTime()))
    .filter(m => m.date >= today)
    .filter(m => m.home === myTeam || m.away === myTeam)
    .sort((a, b) => a.date - b.date);

  const nextMatch = upcoming[0];
  if (!nextMatch) {
    SpreadsheetApp.getUi().alert('No upcoming matches found for The Game 8B');
    return;
  }

  const oppName = nextMatch.home === myTeam ? nextMatch.away : nextMatch.home;
  const oppId   = nextMatch.home === myTeam ? nextMatch.awayId : nextMatch.homeId;

  buildLast3FormWithHeadings(oppId, oppName, 'NextTeam3MatchForm');
}
