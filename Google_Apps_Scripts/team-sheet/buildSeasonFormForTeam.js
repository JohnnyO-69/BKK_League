/********************************************************************
 *     SEASON FORM – LAST 3 MATCHES NOW OLDEST → NEWEST
 *     Full streak + Last 3 chronological + Trend arrows
 ********************************************************************/
function buildSeasonFormForTeam(teamId, teamName, outputSheetName) {
  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  if (!fixturesSheet) {
    SpreadsheetApp.getUi().alert('Sheet "Fixtures" not found!');
    return;
  }

  let out = ss.getSheetByName(outputSheetName);
  if (!out) out = ss.insertSheet(outputSheetName);
  out.clear();

  const delayMs = 400;
  const now = new Date();
  const sixWeeksMs = 42 * 24 * 60 * 60 * 1000;

  const values = fixturesSheet.getDataRange().getValues();
  const header = values.shift();

  const idxMatchId     = header.indexOf('Match ID');
  const idxHome        = header.indexOf('Home Team');
  const idxAway        = header.indexOf('Away Team');
  const idxDate        = header.indexOf('Match Date');
  const idxHomeFrames  = header.indexOf('Home Frames');
  const idxAwayFrames  = header.indexOf('Away Frames');

  if ([idxMatchId, idxHome, idxAway, idxDate].some(i => i < 0)) {
    out.getRange('A1').setValue('Required columns not found in Fixtures sheet');
    return;
  }

  const toDate = v => v instanceof Date ? v : (d => isNaN(d.getTime()) ? new Date(0) : d)(new Date(String(v || '').trim()));

  const matchesRaw = values
    .filter(r => (r[idxHome] === teamName || r[idxAway] === teamName) && r[idxMatchId])
    .map(r => ({
      id: String(r[idxMatchId]),
      date: toDate(r[idxDate]),
      home: r[idxHome],
      away: r[idxAway],
      homeFrames: idxHomeFrames >= 0 ? Number(r[idxHomeFrames] || 0) : 0,
      awayFrames: idxAwayFrames >= 0 ? Number(r[idxAwayFrames] || 0) : 0
    }))
    .filter(m => m.date.getTime() > 0);

  const completed = matchesRaw
    .filter(m => m.date <= now && m.homeFrames + m.awayFrames > 0)
    .sort((a, b) => a.date - b.date);  // oldest first

  if (!completed.length) {
    out.getRange('A1').setValue(`No completed matches found for ${teamName}`);
    return;
  }

  const mostRecentMatch = completed[completed.length - 1];
  const thaiDate = mostRecentMatch.date.toLocaleDateString('en-CA', { timeZone: 'Asia/Bangkok' });

  // FULL SEASON STREAK (oldest → newest)
  const fullStreak = completed.map(m => {
    const weHome = m.home === teamName;
    const ourScore = weHome ? m.homeFrames : m.awayFrames;
    const oppScore = weHome ? m.awayFrames : m.homeFrames;
    return ourScore > oppScore ? 'W' : ourScore < oppScore ? 'L' : 'D';
  }).join('-');

  // LAST 3 MATCHES — NOW OLDEST → NEWEST (chronological)
  const last3 = completed.slice(-3);
  const last3Str = last3.map(m => {
    const weHome = m.home === teamName;
    const opp = weHome ? m.away : m.home;
    const ourScore = weHome ? m.homeFrames : m.awayFrames;
    const oppScore = weHome ? m.awayFrames : m.homeFrames;
    const res = ourScore > oppScore ? ' W' : ourScore < oppScore ? ' L' : ' D';
    return `${opp} (${ourScore}-${oppScore})${res}`;
  }).join(' • ');  // already oldest → newest

  // === PLAYER STATS (unchanged) ===
  const playerMatches = {};
  const frameCache = {};

  const fetchJson = url => {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
    } catch { return null; }
  };

  const recordGame = (player, matchId, matchDate, isDouble, won) => {
    if (!player) return;
    if (!playerMatches[player]) playerMatches[player] = {};
    if (!playerMatches[player][matchId]) {
      playerMatches[player][matchId] = { date: matchDate, sP: 0, sW: 0, dP: 0, dW: 0, points: 0 };
    }
    const g = playerMatches[player][matchId];
    if (isDouble) { g.dP++; if (won) { g.dW++; g.points += 0.5; } }
    else { g.sP++; if (won) { g.sW++; g.points += 1; } }
  };

  completed.forEach(m => {
    const json = fetchJson(`https://api.bkkleague.com/match/details/${m.id}`);
    if (!json || !Array.isArray(json.data)) {
      Utilities.sleep(delayMs);
      return;
    }
    frameCache[m.id] = json.data;
    json.data.forEach(f => {
      const hp = (f.homePlayers || []).map(p => p.nickName || p.nickname || '').filter(Boolean);
      const ap = (f.awayPlayers || []).map(p => p.nickName || p.nickname || '').filter(Boolean);
      const isDouble = hp.length > 1;
      const homeWon = f.homeWin === 1;
      const homeId = f.homeTeamId || f.home_team_id || f.homeTeamid;
      const awayId = f.awayTeamId || f.away_team_id || f.awayTeamid;
      const weAreHome = String(homeId) === String(teamId);
      const weAreAway = String(awayId) === String(teamId);
      if (!weAreHome && !weAreAway) return;
      const ourPlayers = weAreHome ? hp : ap;
      const won = weAreHome ? homeWon : !homeWon;
      ourPlayers.forEach(p => recordGame(p, m.id, m.date, isDouble, won));
    });
    Utilities.sleep(delayMs);
  });

  const rows = Object.keys(playerMatches).map(player => {
    const matchList = Object.values(playerMatches[player]);
    if (!matchList.length) return null;

    const totalPoints = matchList.reduce((s, g) => s + g.points, 0);
    const totalGames  = matchList.reduce((s, g) => s + g.sP + g.dP, 0);
    const ppg = totalGames > 0 ? totalPoints / totalGames : 0;
    const lastPlayed = Math.max(...matchList.map(g => g.date.getTime()));
    if ((now - lastPlayed) > sixWeeksMs) return null;

    let trend = 'New';
    if (matchList.length >= 2) {
      const sorted = matchList.sort((a, b) => b.date - a.date);
      const recent = sorted[0];
      const earlier = sorted.slice(1);
      const earlierPoints = earlier.reduce((s, g) => s + g.points, 0);
      const earlierGames = earlier.reduce((s, g) => s + g.sP + g.dP, 0);
      const earlierPPG = earlierGames > 0 ? earlierPoints / earlierGames : 0;
      const recentPPG = (recent.sP + recent.dP) > 0 ? recent.points / (recent.sP + recent.dP) : 0;
      if (recentPPG > earlierPPG + 0.05) trend = 'Up';
      else if (recentPPG < earlierPPG - 0.05) trend = 'Down';
      else trend = 'Same';
    }

    const sP = matchList.reduce((s, g) => s + g.sP, 0);
    const sW = matchList.reduce((s, g) => s + g.sW, 0);
    const dP = matchList.reduce((s, g) => s + g.dP, 0);
    const dW = matchList.reduce((s, g) => s + g.dW, 0);
    const matchesPlayed = matchList.length;

    return { player, trend, matchesPlayed, totalGames: sP + dP, sP, sW, singlesPct: sP ? sW/sP : 0, dP, dW, doublesPct: dP ? dW/dP : 0, points: totalPoints, ppg };
  }).filter(Boolean);

  rows.sort((a, b) => b.ppg - a.ppg || b.points - a.points || b.totalGames - a.totalGames || a.player.localeCompare(b.player));

  const headers = ['Player','Trend','Matches','Games','Singles','Won','Win%','Doubles','Won','Win%','Points','PPG'];
  const totalCols = headers.length;

  // === PERFECT HEADER BLOCK ===
  out.getRange(1, 1, 1, totalCols).merge().setValue(`Season Form – ${teamName}`)
    .setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('white')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(2, 1, 1, totalCols).merge().setValue(`Most Recent Match Date: ${thaiDate} (Thai Time)`)
    .setFontStyle('italic').setFontColor('#666').setHorizontalAlignment('center').setVerticalAlignment('middle');

  const nowStr = now.toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' });
  out.getRange(3, 1, 1, totalCols).merge().setValue(`Last Refresh: ${nowStr}`)
    .setFontStyle('italic').setFontColor('#666').setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(4, 1, 1, totalCols).merge().setValue(`Last 3 Matches: ${last3Str || 'None'}`)
    .setFontWeight('bold').setFontColor('#333').setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(5, 1, 1, totalCols).merge().setValue(`Season Results: ${fullStreak || 'None'}`)
    .setFontWeight('bold').setFontSize(11).setFontColor('#222')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  out.getRange(7, 1, 1, totalCols).setValues([headers])
    .setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center').setWrap(true);

  if (rows.length > 0) {
    const data = rows.map(r => [
      r.player, '', r.matchesPlayed, r.totalGames,
      r.sP, r.sW, r.singlesPct,
      r.dP, r.dW, r.doublesPct,
      r.points, r.ppg
    ]);
    out.getRange(8, 1, data.length, totalCols).setValues(data);

    out.getRange(8, 7, data.length, 1).setNumberFormat('0.0%');
    out.getRange(8, 10, data.length, 1).setNumberFormat('0.0%');
    out.getRange(8, 12, data.length, 1).setNumberFormat('0.00');
    out.getRange(8, 3, data.length, 10).setHorizontalAlignment('center');

    rows.forEach((r, i) => {
      const cell = out.getRange(8 + i, 2);
      const map = { 'Up': ['Up','#0f9d58'], 'Down': ['Down','#db4437'], 'Same': ['Same','#f4b400'], 'New': ['New','#4285f4'] };
      const [txt, col] = map[r.trend] || ['New','#4285f4'];
      cell.setValue(txt).setFontColor(col).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    });
  } else {
    out.getRange('A8').setValue('No active players in last 6 weeks');
  }

  for (let i = 1; i <= totalCols; i++) out.setColumnWidth(i, i === 2 ? 80 : 100);
  const lastRow = rows.length ? rows.length + 7 : 9;
  out.getRange(1, 1, lastRow, totalCols).setBorder(true, true, true, true, true, true);

  return frameCache;
}