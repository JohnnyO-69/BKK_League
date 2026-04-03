function buildNextTeamOOP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  const paramsSheet = ss.getSheetByName('Parameters');
  if (!fixturesSheet) return;
  if (!paramsSheet) return;

  const outName = 'NextTeamOOP';
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear();

  const myTeam = String(paramsSheet.getRange('F7').getValue() || '').trim();
  if (!myTeam) {
    out.getRange('A1').setValue('Missing team name in Parameters!F7');
    return;
  }

  const values = fixturesSheet.getDataRange().getValues();
  if (values.length <= 1) {
    out.getRange('A1').setValue('Fixtures sheet is empty');
    return;
  }

  const header = values.shift();
  const idxMatchId = header.indexOf('Match ID');
  const idxStatus = header.indexOf('Status');
  const idxDate = header.indexOf('Match Date');
  const idxHome = header.indexOf('Home Team');
  const idxAway = header.indexOf('Away Team');
  const idxHomeFrames = header.indexOf('Home Frames');
  const idxAwayFrames = header.indexOf('Away Frames');
  const idxHomeId = header.indexOf('Home Team ID');
  const idxAwayId = header.indexOf('Away Team ID');

  if ([idxMatchId, idxDate, idxHome, idxAway, idxHomeId, idxAwayId].some(i => i === -1)) {
    out.getRange('A1').setValue('Missing required columns in Fixtures');
    return;
  }

  const todayThai = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Bangkok' }));
  todayThai.setHours(0, 0, 0, 0);

  const fixtureRows = values
    .map(row => buildNextTeamOOPFixtureRow_(row, {
      idxMatchId,
      idxStatus,
      idxDate,
      idxHome,
      idxAway,
      idxHomeFrames,
      idxAwayFrames,
      idxHomeId,
      idxAwayId
    }))
    .filter(Boolean);

  const ourMatches = fixtureRows
    .filter(match => match.home === myTeam || match.away === myTeam)
    .sort((a, b) => a.date - b.date);

  if (!ourMatches.length) {
    out.getRange('A1').setValue(`No fixtures found for ${myTeam}`);
    return;
  }

  const nextMatch = ourMatches.find(match => match.date >= todayThai);
  if (!nextMatch) {
    out.getRange('A1').setValue(`No upcoming match found for ${myTeam}`);
    return;
  }

  const opponentName = nextMatch.home === myTeam ? nextMatch.away : nextMatch.home;
  const opponentId = Number(nextMatch.home === myTeam ? nextMatch.awayId : nextMatch.homeId);

  if (Number.isNaN(opponentId)) {
    out.getRange('A1').setValue(`Could not determine next opponent ID for ${opponentName}`);
    return;
  }

  const completedMatches = fixtureRows
    .filter(match => {
      const involvesOpponent = Number(match.homeId) === opponentId || Number(match.awayId) === opponentId;
      return involvesOpponent && match.date < nextMatch.date && buildNextTeamOOPIsCompleted_(match);
    })
    .sort((a, b) => b.date - a.date)
    .slice(0, 5);

  if (!completedMatches.length) {
    out.getRange('A1').setValue(`No completed matches found for ${opponentName}`);
    return;
  }

  const slotStats = new Map();
  const playerStats = new Map();
  const matchesUsed = [];

  completedMatches.forEach((match, index) => {
    const weight = Math.pow(0.82, index);
    const json = buildNextTeamOOPFetchJson_(`https://api.bkkleague.com/match/details/${match.id}`);
    if (!json || !Array.isArray(json.data)) return;

    const orderedFrames = buildNextTeamOOPOrderFrames_(json.data);

    matchesUsed.push({
      id: match.id,
      date: match.date,
      weight,
      home: match.home,
      away: match.away
    });

    orderedFrames.forEach((item, sequenceIndex) => {
      const frame = item.frame;
      const homeId = Number(frame.homeTeamId || frame.home_team_id || frame.homeTeamid);
      const awayId = Number(frame.awayTeamId || frame.away_team_id || frame.awayTeamid);
      const isOpponentHome = homeId === opponentId;
      const isOpponentAway = awayId === opponentId;
      if (!isOpponentHome && !isOpponentAway) return;

      const players = buildNextTeamOOPExtractPlayerNames_(
        isOpponentHome ? frame.homePlayers : frame.awayPlayers
      );

      if (!players.length) return;

      const slot = sequenceIndex + 1;
      const format = players.length > 1 ? 'Doubles' : 'Singles';
      const lineup = players.join(' + ');

      buildNextTeamOOPRecordSlot_(slotStats, {
        slot,
        format,
        lineup,
        weight,
        date: match.date,
        matchId: match.id
      });

      players.forEach(player => {
        buildNextTeamOOPRecordPlayer_(playerStats, {
          player,
          slot,
          format,
          weight,
          date: match.date
        });
      });
    });
  });

  if (!matchesUsed.length || !slotStats.size) {
    out.getRange('A1').setValue(`Unable to build order-of-play prediction for ${opponentName}`);
    return;
  }

  const slotKeys = Array.from(slotStats.keys()).sort((a, b) => a - b);
  const optimizedChoices = buildNextTeamOOPOptimizeLineup_(slotKeys, slotStats);
  const slotPredictions = optimizedChoices.map(choice => {
    const alternative = choice.candidates.find(candidate => candidate.lineup !== choice.chosen.lineup) || null;
    const confidence = choice.totalWeight ? choice.chosen.weightedScore / choice.totalWeight : 0;

    return [
      choice.slot,
      choice.chosen.format,
      choice.chosen.lineup,
      confidence,
      choice.chosen.matchCount,
      choice.chosen.weightedScore,
      buildNextTeamOOPFormatDate_(choice.chosen.lastSeen),
      alternative ? `${alternative.lineup} (${Math.round((alternative.weightedScore / choice.totalWeight) * 100)}%)` : ''
    ];
  });

  const playerRows = Array.from(playerStats.values())
    .sort((a, b) => {
      if (b.weightedAppearances !== a.weightedAppearances) return b.weightedAppearances - a.weightedAppearances;
      return a.player.localeCompare(b.player);
    })
    .map(stat => [
      stat.player,
      stat.weightedAppearances,
      stat.singlesCount,
      stat.doublesCount,
      stat.weightedAppearances ? stat.weightedSlotTotal / stat.weightedAppearances : 0,
      buildNextTeamOOPFormatDate_(stat.lastSeen)
    ]);

  const title = `Next Opponent Order Of Play – ${opponentName}`;
  const nextMatchDate = buildNextTeamOOPFormatDate_(nextMatch.date);
  const refreshStamp = new Date().toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' });
  const matchesUsedText = matchesUsed
    .map(match => `${buildNextTeamOOPFormatDate_(match.date)} #${match.id} (${match.home} vs ${match.away})`)
    .join(' | ');

  out.getRange('A1:H1').merge()
    .setValue(title)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#1f4e78')
    .setFontColor('white');

  out.getRange('A2:H2').merge()
    .setValue(`Next match: ${nextMatch.home} vs ${nextMatch.away} on ${nextMatchDate}`)
    .setHorizontalAlignment('center')
    .setFontStyle('italic');

  out.getRange('A3:H3').merge()
    .setValue(`Using ${matchesUsed.length} recent completed matches | Refreshed ${refreshStamp}`)
    .setHorizontalAlignment('center')
    .setFontStyle('italic');

  out.getRange('A4:H4').merge()
    .setValue(matchesUsedText)
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setBackground('#ddebf7');

  out.getRange(6, 1, 1, 8).setValues([[
    'Slot',
    'Format',
    'Predicted Player(s)',
    'Confidence',
    'Seen In Matches',
    'Weighted Score',
    'Last Seen',
    'Next Best Option'
  ]])
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (slotPredictions.length) {
    out.getRange(7, 1, slotPredictions.length, 8).setValues(slotPredictions);
    out.getRange(7, 4, slotPredictions.length, 1).setNumberFormat('0.0%');
    out.getRange(7, 6, slotPredictions.length, 1).setNumberFormat('0.00');
    out.getRange(7, 1, slotPredictions.length, 8).setHorizontalAlignment('center');
    out.getRange(7, 3, slotPredictions.length, 1).setHorizontalAlignment('left');
    out.getRange(7, 8, slotPredictions.length, 1).setHorizontalAlignment('left');
  } else {
    out.getRange('A7').setValue('No slot predictions available');
  }

  const playerHeaderRow = Math.max(9 + slotPredictions.length, 12);
  out.getRange(playerHeaderRow, 1, 1, 6).setValues([[
    'Player',
    'Weighted Appearances',
    'Singles Frames',
    'Doubles Frames',
    'Average Slot',
    'Last Seen'
  ]])
    .setFontWeight('bold')
    .setBackground('#fce5cd')
    .setHorizontalAlignment('center');

  if (playerRows.length) {
    out.getRange(playerHeaderRow + 1, 1, playerRows.length, 6).setValues(playerRows);
    out.getRange(playerHeaderRow + 1, 2, playerRows.length, 1).setNumberFormat('0.00');
    out.getRange(playerHeaderRow + 1, 5, playerRows.length, 1).setNumberFormat('0.0');
    out.getRange(playerHeaderRow + 1, 1, playerRows.length, 6).setHorizontalAlignment('center');
    out.getRange(playerHeaderRow + 1, 1, playerRows.length, 1).setHorizontalAlignment('left');
  }

  out.getRange(1, 1, Math.max(playerHeaderRow + Math.max(playerRows.length, 1), 7 + Math.max(slotPredictions.length, 1)), 8)
    .setBorder(true, true, true, true, true, true);

  for (let col = 1; col <= 8; col++) {
    out.autoResizeColumn(col);
  }

  out.setColumnWidth(1, 60);
  out.setColumnWidth(2, 90);
  out.setColumnWidth(3, 180);
  out.setColumnWidth(4, 90);
  out.setColumnWidth(5, 95);
  out.setColumnWidth(6, 95);
  out.setColumnWidth(7, 90);
  out.setColumnWidth(8, 220);
}

function buildNextTeamOOPFixtureRow_(row, indexes) {
  const date = buildNextTeamOOPParseDate_(row[indexes.idxDate]);
  if (!date) return null;

  return {
    id: row[indexes.idxMatchId],
    status: indexes.idxStatus === -1 ? '' : String(row[indexes.idxStatus] || '').trim(),
    date,
    home: String(row[indexes.idxHome] || '').trim(),
    away: String(row[indexes.idxAway] || '').trim(),
    homeFrames: indexes.idxHomeFrames === -1 ? '' : row[indexes.idxHomeFrames],
    awayFrames: indexes.idxAwayFrames === -1 ? '' : row[indexes.idxAwayFrames],
    homeId: row[indexes.idxHomeId],
    awayId: row[indexes.idxAwayId]
  };
}

function buildNextTeamOOPParseDate_(value) {
  if (!value) return null;
  const date = value instanceof Date ? new Date(value) : new Date(String(value).trim());
  if (isNaN(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

function buildNextTeamOOPIsCompleted_(match) {
  const status = String(match.status || '').toLowerCase();
  if (status === 'completed') return true;

  const hasScores = match.homeFrames !== '' && match.awayFrames !== '' &&
    !Number.isNaN(Number(match.homeFrames)) && !Number.isNaN(Number(match.awayFrames));

  return hasScores;
}

function buildNextTeamOOPFetchJson_(url) {
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return response.getResponseCode() === 200 ? JSON.parse(response.getContentText()) : null;
  } catch (error) {
    Logger.log(`NextTeamOOP fetch failed for ${url}: ${error}`);
    return null;
  }
}

function buildNextTeamOOPOrderFrames_(frames) {
  return frames
    .map((frame, index) => ({
      frame,
      index,
      timestamp: buildNextTeamOOPFrameTimestamp_(frame)
    }))
    .sort((a, b) => {
      const aHasTimestamp = a.timestamp !== null;
      const bHasTimestamp = b.timestamp !== null;

      if (aHasTimestamp && bHasTimestamp && a.timestamp !== b.timestamp) {
        return a.timestamp - b.timestamp;
      }

      if (aHasTimestamp !== bHasTimestamp) {
        return aHasTimestamp ? -1 : 1;
      }

      return a.index - b.index;
    });
}

function buildNextTeamOOPFrameTimestamp_(frame) {
  const candidates = [
    frame.playedAt,
    frame.played_at,
    frame.startedAt,
    frame.started_at,
    frame.updatedAt,
    frame.updated_at,
    frame.createdAt,
    frame.created_at,
    frame.finishedAt,
    frame.finished_at,
    frame.time,
    frame.date
  ];

  for (let index = 0; index < candidates.length; index++) {
    const value = candidates[index];
    if (!value) continue;

    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.getTime();
    }
  }

  return null;
}

function buildNextTeamOOPExtractPlayerNames_(players) {
  return (players || [])
    .map(player => {
      const firstName = String(player.firstName || player.first_name || '').trim();
      const lastName = String(player.lastName || player.last_name || '').trim();
      const fullName = [firstName, lastName].filter(Boolean).join(' ');

      return String(
        player.nickName ||
        player.nickname ||
        player.nick_name ||
        player.displayName ||
        player.display_name ||
        player.fullName ||
        player.full_name ||
        player.name ||
        fullName ||
        ''
      ).trim();
    })
    .filter(Boolean);
}

function buildNextTeamOOPOptimizeLineup_(slotKeys, slotStats) {
  const beamWidth = 12;
  const maxCandidatesPerSlot = 4;
  let states = [{
    score: 0,
    choices: [],
    previousChosen: null,
    recentPlayers: []
  }];

  slotKeys.forEach(slot => {
    const bucket = slotStats.get(slot);
    const candidates = Array.from(bucket.candidates.values())
      .sort((a, b) => {
        if (b.weightedScore !== a.weightedScore) return b.weightedScore - a.weightedScore;
        if (b.matchCount !== a.matchCount) return b.matchCount - a.matchCount;
        return a.lineup.localeCompare(b.lineup);
      })
      .slice(0, maxCandidatesPerSlot);

    if (!candidates.length) return;

    const nextStates = [];

    states.forEach(state => {
      candidates.forEach(candidate => {
        const previousChosen = state.previousChosen;
        const baseScore = buildNextTeamOOPCandidateBaseScore_(candidate, bucket.totalWeight);
        const adjacencyPenalty = buildNextTeamOOPAdjacencyPenalty_(candidate, previousChosen);
        const usagePenalty = buildNextTeamOOPUsagePenalty_(candidate, state.recentPlayers);
        const adjustedScore = baseScore - adjacencyPenalty - usagePenalty;
        const chosen = {
          slot,
          format: candidate.format,
          lineup: candidate.lineup,
          players: buildNextTeamOOPSplitLineup_(candidate.lineup)
        };

        nextStates.push({
          score: state.score + adjustedScore,
          choices: state.choices.concat({
            slot,
            chosen: candidate,
            candidates,
            totalWeight: bucket.totalWeight
          }),
          previousChosen: chosen,
          recentPlayers: buildNextTeamOOPBuildRecentPlayers_(state.recentPlayers, chosen.players)
        });
      });
    });

    states = nextStates
      .sort((a, b) => b.score - a.score)
      .slice(0, beamWidth);
  });

  return states.length ? states[0].choices : [];
}

function buildNextTeamOOPCandidateBaseScore_(candidate, totalWeight) {
  return totalWeight ? candidate.weightedScore / totalWeight : 0;
}

function buildNextTeamOOPUsagePenalty_(candidate, recentPlayers) {
  const candidatePlayers = buildNextTeamOOPSplitLineup_(candidate.lineup);
  if (!candidatePlayers.length || !recentPlayers.length) return 0;

  let penalty = 0;
  candidatePlayers.forEach(player => {
    const firstBack = recentPlayers[0] || [];
    const secondBack = recentPlayers[1] || [];

    if (firstBack.includes(player)) penalty += 0.12;
    if (secondBack.includes(player)) penalty += 0.05;
  });

  return penalty;
}

function buildNextTeamOOPBuildRecentPlayers_(recentPlayers, currentPlayers) {
  return [currentPlayers].concat(recentPlayers).slice(0, 2);
}

function buildNextTeamOOPAdjacencyPenalty_(candidate, previousChosen) {
  if (!previousChosen) return 0;

  const candidatePlayers = buildNextTeamOOPSplitLineup_(candidate.lineup);
  const previousPlayers = previousChosen.players || [];
  const overlapCount = candidatePlayers.filter(player => previousPlayers.includes(player)).length;

  if (!overlapCount) return 0;

  if (candidate.format === 'Singles' && previousChosen.format === 'Singles') {
    return 0.25 * overlapCount;
  }

  if (candidate.format === 'Doubles' && previousChosen.format === 'Doubles') {
    return 0.18 * overlapCount;
  }

  return 0.08 * overlapCount;
}

function buildNextTeamOOPSplitLineup_(lineup) {
  return String(lineup || '')
    .split(' + ')
    .map(name => name.trim())
    .filter(Boolean);
}

function buildNextTeamOOPRecordSlot_(slotStats, details) {
  if (!slotStats.has(details.slot)) {
    slotStats.set(details.slot, {
      totalWeight: 0,
      candidates: new Map()
    });
  }

  const bucket = slotStats.get(details.slot);
  bucket.totalWeight += details.weight;

  if (!bucket.candidates.has(details.lineup)) {
    bucket.candidates.set(details.lineup, {
      lineup: details.lineup,
      format: details.format,
      weightedScore: 0,
      matchCount: 0,
      lastSeen: new Date(0),
      matchIds: new Set()
    });
  }

  const candidate = bucket.candidates.get(details.lineup);
  candidate.weightedScore += details.weight;
  candidate.matchIds.add(String(details.matchId));
  candidate.matchCount = candidate.matchIds.size;
  if (details.date > candidate.lastSeen) candidate.lastSeen = details.date;
}

function buildNextTeamOOPRecordPlayer_(playerStats, details) {
  if (!playerStats.has(details.player)) {
    playerStats.set(details.player, {
      player: details.player,
      weightedAppearances: 0,
      singlesCount: 0,
      doublesCount: 0,
      weightedSlotTotal: 0,
      lastSeen: new Date(0)
    });
  }

  const stat = playerStats.get(details.player);
  stat.weightedAppearances += details.weight;
  stat.weightedSlotTotal += details.slot * details.weight;
  if (details.format === 'Doubles') {
    stat.doublesCount += 1;
  } else {
    stat.singlesCount += 1;
  }
  if (details.date > stat.lastSeen) stat.lastSeen = details.date;
}

function buildNextTeamOOPFormatDate_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return date.toLocaleDateString('en-GB', { timeZone: 'Asia/Bangkok' });
}