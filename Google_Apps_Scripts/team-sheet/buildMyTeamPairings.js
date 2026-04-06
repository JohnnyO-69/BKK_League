/********************************************************************
 *     MY TEAM PAIRINGS – Historical doubles pair analysis
 *     Reads all completed season matches, fetches frame data, and
 *     aggregates won/played for every doubles pairing used.
 *
 *     Sheet: MyTeamPairings
 *     Called by: buildMyTeamForm()
 ********************************************************************/
function buildMyTeamPairings(frameCache) {
  console.log('START buildMyTeamPairings');

  const ss = getLeagueSpreadsheet_();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  const paramsSheet = ss.getSheetByName('Parameters');
  const last3Sheet = ss.getSheetByName('Last3MatchForm');
  if (!fixturesSheet) { console.log('ERROR: Fixtures sheet not found'); return; }
  if (!paramsSheet) { console.log('ERROR: Parameters sheet not found'); return; }
  if (!last3Sheet) { console.log('ERROR: Last3MatchForm sheet not found'); return; }

  const outName = 'MyTeamPairings';
  let out = ss.getSheetByName(outName);
  if (!out) {
    console.log('Creating sheet:', outName);
    out = ss.insertSheet(outName);
  }
  out.clear();

  const teamName = String(paramsSheet.getRange('F7').getValue() || '').trim();
  const teamId = Number(paramsSheet.getRange('F8').getValue());
  console.log('Team name:', teamName, '| Team ID:', teamId);

  if (!teamName || Number.isNaN(teamId)) {
    console.log('ERROR: Missing team name/id in Parameters!F7:F8');
    out.getRange('A1').setValue('Missing team name/id in Parameters!F7:F8');
    return;
  }

  // ── Read completed matches from Fixtures ──────────────────────────
  const values = fixturesSheet.getDataRange().getValues();
  const header = values.shift();
  console.log('Fixtures headers:', header.join(', '));

  const idxMatchId    = header.indexOf('Match ID');
  const idxHome       = header.indexOf('Home Team');
  const idxAway       = header.indexOf('Away Team');
  const idxDate       = header.indexOf('Match Date');
  const idxHomeFrames = header.indexOf('Home Frames');
  const idxAwayFrames = header.indexOf('Away Frames');

  console.log('Column indices – MatchID:', idxMatchId, 'Home:', idxHome, 'Away:', idxAway, 'Date:', idxDate, 'HomeFrames:', idxHomeFrames, 'AwayFrames:', idxAwayFrames);

  if ([idxMatchId, idxHome, idxAway, idxDate].some(i => i < 0)) {
    console.log('ERROR: Required columns not found in Fixtures sheet');
    out.getRange('A1').setValue('Required columns not found in Fixtures sheet');
    return;
  }

  const toDate = v => v instanceof Date ? v : (d => isNaN(d.getTime()) ? new Date(0) : d)(new Date(String(v || '').trim()));
  const now = new Date();

  const teamRows = values.filter(r => r[idxHome] === teamName || r[idxAway] === teamName);
  console.log('Rows matching team name:', teamRows.length);

  const hasFrameScores = idxHomeFrames >= 0 && idxAwayFrames >= 0;
  console.log('Home/Away Frames columns present:', hasFrameScores);

  const completed = values
    .filter(r =>
      (r[idxHome] === teamName || r[idxAway] === teamName) &&
      r[idxMatchId] &&
      idxHomeFrames >= 0 && idxAwayFrames >= 0 &&
      (Number(r[idxHomeFrames] || 0) + Number(r[idxAwayFrames] || 0)) > 0
    )
    .map(r => ({
      id: String(r[idxMatchId]),
      date: toDate(r[idxDate]),
      home: r[idxHome],
      away: r[idxAway]
    }))
    .filter(m => m.date.getTime() > 0 && m.date <= now)
    .sort((a, b) => a.date - b.date);

  console.log('Completed matches found:', completed.length, '| IDs:', completed.map(m => m.id).join(', '));

  if (!completed.length) {
    console.log('ERROR: No completed matches found for', teamName);
    out.getRange('A1').setValue(`No completed matches found for ${teamName}`);
    return;
  }

  // ── Read filtered player pool from Last3MatchForm (same as MyTeamOOP) ─
  const last3Headers = last3Sheet.getRange(6, 1, 1, Math.min(last3Sheet.getLastColumn(), 20)).getValues()[0];
  const last3HeaderMap = buildMyTeamOOPHeaderMap_(last3Headers);
  const playerPool = last3HeaderMap.player !== -1
    ? buildMyTeamOOPVisiblePlayerPool_(last3Sheet, last3HeaderMap)
    : [];
  const selectedNames = new Set(playerPool.map(p => p.name));
  console.log('Visible players from Last3MatchForm:', playerPool.map(p => p.name).join(', '));

  console.log('frameCache provided:', !!frameCache, '| Cached match IDs:', frameCache ? Object.keys(frameCache).join(', ') : 'none');

  // ── Fetch frame data and aggregate pairing stats ──────────────────
  const delayMs = 400;

  // pairings[pairKey] = { players: [a, b], played: 0, won: 0 }
  const pairings = {};
  // individual[playerName] = { played: 0, won: 0 }
  const individual = {};

  const fetchJson = url => {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
    } catch { return null; }
  };

  const recordPair = (playerA, playerB, won) => {
    const sorted = [playerA, playerB].sort((a, b) => a.localeCompare(b));
    const key = sorted.join(' + ');
    if (!pairings[key]) pairings[key] = { players: sorted, played: 0, won: 0 };
    pairings[key].played++;
    if (won) pairings[key].won++;
  };

  const recordIndividual = (player, won) => {
    if (!individual[player]) individual[player] = { played: 0, won: 0 };
    individual[player].played++;
    if (won) individual[player].won++;
  };

  completed.forEach(m => {
    const cachedFrames = frameCache && frameCache[m.id];
    console.log('Match', m.id, '– using cache:', !!cachedFrames);
    const frames = cachedFrames || (() => {
      const json = fetchJson(`https://api.bkkleague.com/match/details/${m.id}`);
      if (!json || !Array.isArray(json.data)) {
        console.log('Match', m.id, '– fetch failed or no data array');
        Utilities.sleep(delayMs);
        return null;
      }
      Utilities.sleep(delayMs);
      return json.data;
    })();
    if (!frames) { console.log('Match', m.id, '– no frames, skipping'); return; }

    console.log('Match', m.id, '– total frames:', frames.length);
    let doublesFrameCount = 0;

    frames.forEach(frame => {
      const hp = (frame.homePlayers || []).map(p => p.nickName || p.nickname || '').filter(Boolean);
      const ap = (frame.awayPlayers || []).map(p => p.nickName || p.nickname || '').filter(Boolean);

      const isDouble = hp.length > 1 || ap.length > 1;
      if (!isDouble) return;
      doublesFrameCount++;

      const homeId = frame.homeTeamId || frame.home_team_id || frame.homeTeamid;
      const awayId = frame.awayTeamId || frame.away_team_id || frame.awayTeamid;
      const weAreHome = String(homeId) === String(teamId);
      const weAreAway = String(awayId) === String(teamId);
      if (!weAreHome && !weAreAway) return;

      const ourPlayers = weAreHome ? hp : ap;
      const homeWon = frame.homeWin === 1;
      const won = weAreHome ? homeWon : !homeWon;

      if (ourPlayers.length === 2) {
        recordPair(ourPlayers[0], ourPlayers[1], won);
      }
      ourPlayers.forEach(player => recordIndividual(player, won));
    });
    console.log('Match', m.id, '– doubles frames processed:', doublesFrameCount);
  });

  console.log('Total unique pairings found:', Object.keys(pairings).length);
  console.log('Total individual doubles players:', Object.keys(individual).length);
  if (Object.keys(pairings).length > 0) {
    console.log('Pairings:', Object.keys(pairings).join(' | '));
  }

  // ── Sort pairs: win% desc, then played desc, then name asc ────────
  // Only include pairs where both players are currently selected in Last3MatchForm
  const pairRows = Object.values(pairings)
    .filter(pair => !selectedNames.size || (selectedNames.has(pair.players[0]) && selectedNames.has(pair.players[1])))
    .sort((a, b) => {
      const aWinPct = a.played ? a.won / a.played : 0;
      const bWinPct = b.played ? b.won / b.played : 0;
      if (bWinPct !== aWinPct) return bWinPct - aWinPct;
      if (b.played !== a.played) return b.played - a.played;
      return a.players.join(' + ').localeCompare(b.players.join(' + '));
    })
    .map(pair => [
      pair.players.join(' + '),
      pair.played,
      pair.won,
      pair.played ? pair.won / pair.played : 0
    ]);

  console.log('Pairs after player filter:', pairRows.length);

  // ── Sort individual: win% desc, played desc, name asc ─────────────
  // Only include players currently selected in Last3MatchForm
  const individualRows = Object.entries(individual)
    .filter(([name]) => !selectedNames.size || selectedNames.has(name))
    .sort(([nameA, a], [nameB, b]) => {
      const aWinPct = a.played ? a.won / a.played : 0;
      const bWinPct = b.played ? b.won / b.played : 0;
      if (bWinPct !== aWinPct) return bWinPct - aWinPct;
      if (b.played !== a.played) return b.played - a.played;
      return nameA.localeCompare(nameB);
    })
    .map(([player, stats]) => [
      player,
      stats.played,
      stats.won,
      stats.played ? stats.won / stats.played : 0
    ]);

  // ── Write output sheet ─────────────────────────────────────────────
  const refreshStamp = now.toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' });
  const matchCount = completed.length;

  const PAIR_COLS = 4;

  // Title
  out.getRange(1, 1, 1, PAIR_COLS).merge()
    .setValue(`Doubles Pairings Analysis – ${teamName}`)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#274e13')
    .setFontColor('white');

  const playerList = playerPool.length ? playerPool.map(p => p.name).join(', ') : 'all players';
  out.getRange(2, 1, 1, PAIR_COLS).merge()
    .setValue(`Based on ${matchCount} completed match${matchCount === 1 ? '' : 'es'} | Filtered to: ${playerList} | Refreshed ${refreshStamp}`)
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setFontColor('#666')
    .setWrap(true);

  // Pairs table header
  const PAIRS_HEADER_ROW = 4;
  out.getRange(PAIRS_HEADER_ROW, 1, 1, PAIR_COLS).setValues([[
    'Pair',
    'Played Together',
    'Won',
    'Win %'
  ]])
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (pairRows.length) {
    out.getRange(PAIRS_HEADER_ROW + 1, 1, pairRows.length, PAIR_COLS).setValues(pairRows);
    out.getRange(PAIRS_HEADER_ROW + 1, 4, pairRows.length, 1).setNumberFormat('0.0%');
    out.getRange(PAIRS_HEADER_ROW + 1, 2, pairRows.length, 3).setHorizontalAlignment('center');
    out.getRange(PAIRS_HEADER_ROW + 1, 1, pairRows.length, 1).setHorizontalAlignment('left');
  } else {
    out.getRange(PAIRS_HEADER_ROW + 1, 1).setValue('No doubles frames found');
  }

  // Individual doubles summary
  const IND_HEADER_ROW = PAIRS_HEADER_ROW + (pairRows.length || 1) + 3;
  out.getRange(IND_HEADER_ROW, 1, 1, PAIR_COLS).merge()
    .setValue('Individual Doubles Summary')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#fce5cd')
    .setHorizontalAlignment('center');

  out.getRange(IND_HEADER_ROW + 1, 1, 1, PAIR_COLS).setValues([[
    'Player',
    'Doubles Played',
    'Won',
    'Win %'
  ]])
    .setFontWeight('bold')
    .setBackground('#fce5cd')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (individualRows.length) {
    out.getRange(IND_HEADER_ROW + 2, 1, individualRows.length, PAIR_COLS).setValues(individualRows);
    out.getRange(IND_HEADER_ROW + 2, 4, individualRows.length, 1).setNumberFormat('0.0%');
    out.getRange(IND_HEADER_ROW + 2, 2, individualRows.length, 3).setHorizontalAlignment('center');
    out.getRange(IND_HEADER_ROW + 2, 1, individualRows.length, 1).setHorizontalAlignment('left');
  } else {
    out.getRange(IND_HEADER_ROW + 2, 1).setValue('No data');
  }

  // Borders and column widths
  const lastRow = IND_HEADER_ROW + 2 + Math.max(individualRows.length, 1);
  out.getRange(1, 1, lastRow, PAIR_COLS).setBorder(true, true, true, true, true, true);

  out.setColumnWidth(1, 220);
  out.setColumnWidth(2, 120);
  out.setColumnWidth(3, 80);
  out.setColumnWidth(4, 80);

  console.log('END buildMyTeamPairings – sheet written successfully');
}
