function buildMyTeamOOP() {
  const ss = getLeagueSpreadsheet_();
  const source = ss.getSheetByName('Last3MatchForm');
  const paramsSheet = ss.getSheetByName('Parameters');
  const gameTemplate = getMyTeamOOPGameTemplateSheet_(ss);
  if (!source) return;
  if (!paramsSheet) return;

  const outName = 'MyTeamOOP';
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear();

  const teamName = String(paramsSheet.getRange('F7').getValue() || '').trim() || 'My Team';

  const headerValues = source.getRange(6, 1, 1, Math.min(source.getLastColumn(), 20)).getValues()[0];
  const headerMap = buildMyTeamOOPHeaderMap_(headerValues);

  if (headerMap.player === -1) {
    out.getRange('A1').setValue('Could not find player column in Last3MatchForm row 6');
    return;
  }

  if (!gameTemplate) {
    out.getRange('A1').setValue('Could not find Game Template sheet');
    return;
  }

  const playerPool = buildMyTeamOOPVisiblePlayerPool_(source, headerMap, gameTemplate);
  if (!playerPool.length) {
    out.getRange('A1').setValue('No matching players found for Game Template K2:K13');
    return;
  }

  const lineup = buildMyTeamOOPSuggestedLineup_(playerPool);
  const refreshStamp = new Date().toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' });

  out.getRange('A1:H1').merge()
    .setValue(`Suggested Order Of Play – ${teamName}`)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#274e13')
    .setFontColor('white');

  out.getRange('A2:H2').merge()
    .setValue('Players sourced from Game Template!K2:K13')
    .setHorizontalAlignment('center')
    .setFontStyle('italic');

  out.getRange('A3:H3').merge()
    .setValue(`Using ${playerPool.length} selected players | Refreshed ${refreshStamp}`)
    .setHorizontalAlignment('center')
    .setFontStyle('italic');

  out.getRange('A4:H4').merge()
    .setValue(playerPool.map(player => player.name).join(' | '))
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setBackground('#d9ead3');

  out.getRange(6, 1, 1, 8).setValues([[
    'Slot',
    'Format',
    'Suggested Player(s)',
    'Suitability',
    'Singles Weight',
    'Doubles Weight',
    'PPG',
    'Notes'
  ]])
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (lineup.length) {
    out.getRange(7, 1, lineup.length, 8).setValues(lineup.map(item => [
      item.slot,
      item.format,
      item.lineup,
      item.score,
      item.singlesWeight,
      item.doublesWeight,
      item.ppg,
      item.note
    ]));
    out.getRange(7, 4, lineup.length, 1).setNumberFormat('0.00');
    out.getRange(7, 5, lineup.length, 2).setNumberFormat('0.00');
    out.getRange(7, 7, lineup.length, 1).setNumberFormat('0.00');
    out.getRange(7, 1, lineup.length, 8).setHorizontalAlignment('center');
    out.getRange(7, 3, lineup.length, 1).setHorizontalAlignment('left');
    out.getRange(7, 8, lineup.length, 1).setHorizontalAlignment('left');
  }

  const summaryHeaderRow = 29;
  out.getRange(summaryHeaderRow, 1, 1, 7).setValues([[
    'Player',
    'Singles Weight',
    'Doubles Weight',
    'PPG',
    'Singles Selected',
    'Doubles Selected',
    'Source Row'
  ]])
    .setFontWeight('bold')
    .setBackground('#fce5cd')
    .setHorizontalAlignment('center');

  const summaryRows = playerPool
    .sort((a, b) => {
      if (b.combinedWeight !== a.combinedWeight) return b.combinedWeight - a.combinedWeight;
      return a.name.localeCompare(b.name);
    })
    .map(player => [
      player.name,
      player.singlesWeight,
      player.doublesWeight,
      player.ppg,
      player.assignedSingles,
      player.assignedDoubles,
      player.sourceRow
    ]);

  if (summaryRows.length) {
    out.getRange(summaryHeaderRow + 1, 1, summaryRows.length, 7).setValues(summaryRows);
    out.getRange(summaryHeaderRow + 1, 2, summaryRows.length, 3).setNumberFormat('0.00');
    out.getRange(summaryHeaderRow + 1, 1, summaryRows.length, 7).setHorizontalAlignment('center');
    out.getRange(summaryHeaderRow + 1, 1, summaryRows.length, 1).setHorizontalAlignment('left');
  }

  out.getRange(1, 1, summaryHeaderRow + Math.max(summaryRows.length, 1), 8)
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
  out.setColumnWidth(7, 80);
  out.setColumnWidth(8, 220);
}

function buildMyTeamOOPHeaderMap_(headers) {
  const findIndex = patterns => headers.findIndex(header => {
    const value = String(header || '').trim().toLowerCase();
    return patterns.some(pattern => pattern.test(value));
  });

  return {
    player: findIndex([/^player$/, /player name/]),
    singlesPlayed: findIndex([/^singles played$/, /singles/]),
    singlesWinPct: findIndex([/^singles win %$/, /singles.*win/]),
    doublesPlayed: findIndex([/^doubles played$/, /doubles played/]),
    doublesWinPct: findIndex([/^doubles win %$/, /doubles.*win/]),
    ppg: findIndex([/^points per game$/, /^ppg$/]),
    points: findIndex([/^points$/])
  };
}

function buildMyTeamOOPVisiblePlayerPool_(sheet, headerMap, gameTemplate) {
  const ktValues = gameTemplate.getRange('K2:K13').getDisplayValues();
  const selectedEntries = ktValues
    .map(row => String(row[0] || '').trim())
    .filter(name => name)
    .map(name => ({ rawName: name, normalizedName: buildMyTeamOOPNormalizeName_(name) }));
  const selectedNames = new Set(selectedEntries.map(entry => entry.normalizedName));

  const rows = sheet.getRange(7, 1, 13, Math.min(sheet.getLastColumn(), 20)).getValues();
  const players = [];
  const matchedNames = new Set();

  rows.forEach((row, index) => {
    const rowNumber = index + 7;
    const name = String(row[headerMap.player] || '').trim();
    if (!name) return;

    const normalizedName = buildMyTeamOOPNormalizeName_(name);
    if (!selectedNames.has(normalizedName)) return;

    matchedNames.add(normalizedName);

    const singlesPlayed = buildMyTeamOOPToNumber_(row[headerMap.singlesPlayed]);
    const singlesWinPct = buildMyTeamOOPToPercent_(row[headerMap.singlesWinPct]);
    const doublesPlayed = buildMyTeamOOPToNumber_(row[headerMap.doublesPlayed]);
    const doublesWinPct = buildMyTeamOOPToPercent_(row[headerMap.doublesWinPct]);
    const ppg = buildMyTeamOOPToNumber_(row[headerMap.ppg]);
    const points = buildMyTeamOOPToNumber_(row[headerMap.points]);
    const singlesWeight = singlesWinPct * 0.6 + Math.min(singlesPlayed, 4) / 4 * 0.2 + Math.min(ppg, 1) * 0.2;
    const doublesWeight = doublesWinPct * 0.5 + Math.min(doublesPlayed, 4) / 4 * 0.3 + Math.min(ppg, 1) * 0.2;

    players.push({
      name,
      sourceRow: rowNumber,
      singlesPlayed,
      singlesWinPct,
      doublesPlayed,
      doublesWinPct,
      ppg,
      points,
      singlesWeight,
      doublesWeight,
      combinedWeight: singlesWeight + doublesWeight,
      assignedSingles: 0,
      assignedDoubles: 0,
      assignedSinglesByBlock: {},
      assignedDoublesByBlock: {}
    });
  });

  selectedEntries.forEach(entry => {
    if (matchedNames.has(entry.normalizedName)) return;

    players.push({
      name: entry.rawName,
      sourceRow: 'Game Template only',
      singlesPlayed: 0,
      singlesWinPct: 0,
      doublesPlayed: 0,
      doublesWinPct: 0,
      ppg: 0,
      points: 0,
      singlesWeight: 0,
      doublesWeight: 0,
      combinedWeight: 0,
      assignedSingles: 0,
      assignedDoubles: 0,
      assignedSinglesByBlock: {},
      assignedDoublesByBlock: {}
    });
  });

  return players;
}

function buildMyTeamOOPSuggestedLineup_(playerPool) {
  const formats = [
    'Singles', 'Singles', 'Singles', 'Singles',
    'Doubles', 'Doubles', 'Doubles', 'Doubles',
    'Singles', 'Singles', 'Singles', 'Singles',
    'Doubles', 'Doubles', 'Doubles', 'Doubles',
    'Singles', 'Singles', 'Singles', 'Singles'
  ];

  const lineup = [];
  const recentPlayers = [];
  const recentPairs = [];

  formats.forEach((format, index) => {
    const slot = index + 1;
    const blockIndex = Math.floor(index / 4);

    if (format === 'Singles') {
      const choice = buildMyTeamOOPChooseSingles_(playerPool, recentPlayers, blockIndex);
      if (!choice) return;

      choice.player.assignedSingles += 1;
      buildMyTeamOOPIncrementSinglesBlock_(choice.player, blockIndex);
      lineup.push({
        slot,
        format,
        lineup: choice.player.name,
        score: choice.score,
        singlesWeight: choice.player.singlesWeight,
        doublesWeight: choice.player.doublesWeight,
        ppg: choice.player.ppg,
        note: `Singles priority; selected ${choice.player.assignedSingles} time(s)`
      });

      recentPlayers.unshift([choice.player.name]);
      recentPlayers.splice(2);
      return;
    }

    const pairChoice = buildMyTeamOOPChooseDoubles_(playerPool, recentPlayers, recentPairs, blockIndex);
    if (!pairChoice) return;

    pairChoice.players.forEach(player => {
      player.assignedDoubles += 1;
      buildMyTeamOOPIncrementDoublesBlock_(player, blockIndex);
    });

    lineup.push({
      slot,
      format,
      lineup: pairChoice.players.map(player => player.name).join(' + '),
      score: pairChoice.score,
      singlesWeight: pairChoice.players.reduce((sum, player) => sum + player.singlesWeight, 0) / pairChoice.players.length,
      doublesWeight: pairChoice.players.reduce((sum, player) => sum + player.doublesWeight, 0) / pairChoice.players.length,
      ppg: pairChoice.players.reduce((sum, player) => sum + player.ppg, 0) / pairChoice.players.length,
      note: `Doubles pairing ${pairChoice.timesUsed + 1} time(s)`
    });

    const pairNames = pairChoice.players.map(player => player.name);
    recentPlayers.unshift(pairNames);
    recentPlayers.splice(2);
    recentPairs.unshift(pairNames.join(' + '));
    recentPairs.splice(2);
  });

  buildMyTeamOOPEnsureSinglesCoverage_(lineup, playerPool);
  return lineup;
}

function buildMyTeamOOPEnsureSinglesCoverage_(lineup, playerPool) {
  const singlesEntries = lineup.filter(item => item.format === 'Singles');
  if (!singlesEntries.length) return;

  const singlesByPlayer = new Map();
  singlesEntries.forEach(item => {
    singlesByPlayer.set(item.lineup, (singlesByPlayer.get(item.lineup) || 0) + 1);
  });

  const playerByName = new Map(playerPool.map(player => [player.name, player]));

  const missingSinglesPlayers = playerPool.filter(player => (singlesByPlayer.get(player.name) || 0) === 0);

  missingSinglesPlayers.forEach(player => {
    const availableReplacementSlots = singlesEntries.filter(item => {
      if ((singlesByPlayer.get(item.lineup) || 0) <= 1) return false;
      const blockIndex = Math.floor((item.slot - 1) / 4);
      return buildMyTeamOOPSinglesCountInBlock_(player, blockIndex) < 1;
    });

    const replacement = availableReplacementSlots
      .sort((a, b) => a.score - b.score || a.slot - b.slot)[0];

    if (!replacement) return;

    const replacedPlayer = playerByName.get(replacement.lineup);
    const replacementBlockIndex = Math.floor((replacement.slot - 1) / 4);

    if (replacedPlayer && replacedPlayer.assignedSinglesByBlock) {
      replacedPlayer.assignedSinglesByBlock[replacementBlockIndex] = Math.max(
        0,
        buildMyTeamOOPSinglesCountInBlock_(replacedPlayer, replacementBlockIndex) - 1
      );
    }

    singlesByPlayer.set(replacement.lineup, (singlesByPlayer.get(replacement.lineup) || 1) - 1);

    replacement.lineup = player.name;
    replacement.score = player.singlesWeight;
    replacement.singlesWeight = player.singlesWeight;
    replacement.doublesWeight = player.doublesWeight;
    replacement.ppg = player.ppg;
    replacement.note = player.sourceRow === 'Game Template only'
      ? 'Selected from Game Template; no Last3 history yet'
      : 'Inserted to ensure every selected player appears in singles';

    buildMyTeamOOPIncrementSinglesBlock_(player, replacementBlockIndex);
    singlesByPlayer.set(player.name, (singlesByPlayer.get(player.name) || 0) + 1);
  });

  const assignedSinglesByName = new Map();
  singlesEntries.forEach(item => {
    assignedSinglesByName.set(item.lineup, (assignedSinglesByName.get(item.lineup) || 0) + 1);
  });

  playerPool.forEach(player => {
    player.assignedSingles = assignedSinglesByName.get(player.name) || 0;
  });
}

function buildMyTeamOOPChooseSingles_(playerPool, recentPlayers, blockIndex) {
  const candidates = playerPool
    .filter(player => buildMyTeamOOPSinglesCountInBlock_(player, blockIndex) < 1)
    .map(player => {
      let score = player.singlesWeight;
      const mostRecent = recentPlayers[0] || [];
      const previous = recentPlayers[1] || [];

      if (mostRecent.includes(player.name)) score -= 0.18;
      if (previous.includes(player.name)) score -= 0.08;
      score -= player.assignedSingles * 0.07;
      if (buildMyTeamOOPTotalAssignments_(player) === 0) score += 0.45;

      return { player, score };
    })
    .sort((a, b) => b.score - a.score || a.player.name.localeCompare(b.player.name));

  return candidates[0] || null;
}

function buildMyTeamOOPChooseDoubles_(playerPool, recentPlayers, recentPairs, blockIndex) {
  const pairs = [];

  for (let i = 0; i < playerPool.length; i++) {
    for (let j = i + 1; j < playerPool.length; j++) {
      const first = playerPool[i];
      const second = playerPool[j];
      if (buildMyTeamOOPDoublesCountInBlock_(first, blockIndex) >= 2) continue;
      if (buildMyTeamOOPDoublesCountInBlock_(second, blockIndex) >= 2) continue;

      const pairName = `${first.name} + ${second.name}`;

      let score = ((first.doublesWeight + second.doublesWeight) / 2) + ((first.ppg + second.ppg) / 2) * 0.15;
      const mostRecent = recentPlayers[0] || [];
      const previous = recentPlayers[1] || [];
      const overlapRecent = [first.name, second.name].filter(name => mostRecent.includes(name)).length;
      const overlapPrevious = [first.name, second.name].filter(name => previous.includes(name)).length;

      score -= overlapRecent * 0.12;
      score -= overlapPrevious * 0.05;
      score -= (first.assignedDoubles + second.assignedDoubles) * 0.04;
      if (recentPairs[0] === pairName) score -= 0.14;
      if (recentPairs[1] === pairName) score -= 0.06;
      if (buildMyTeamOOPTotalAssignments_(first) === 0) score += 0.28;
      if (buildMyTeamOOPTotalAssignments_(second) === 0) score += 0.28;

      pairs.push({
        players: [first, second],
        score,
        pairName,
        timesUsed: [recentPairs[0], recentPairs[1]].filter(name => name === pairName).length
      });
    }
  }

  pairs.sort((a, b) => b.score - a.score || a.pairName.localeCompare(b.pairName));
  return pairs[0] || null;
}

function buildMyTeamOOPToNumber_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function buildMyTeamOOPToPercent_(value) {
  const number = buildMyTeamOOPToNumber_(value);
  if (number > 1) return number / 100;
  return number;
}

function buildMyTeamOOPNormalizeName_(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function buildMyTeamOOPTotalAssignments_(player) {
  return Number(player.assignedSingles || 0) + Number(player.assignedDoubles || 0);
}

function buildMyTeamOOPSinglesCountInBlock_(player, blockIndex) {
  const byBlock = player.assignedSinglesByBlock || {};
  return Number(byBlock[blockIndex] || 0);
}

function buildMyTeamOOPDoublesCountInBlock_(player, blockIndex) {
  const byBlock = player.assignedDoublesByBlock || {};
  return Number(byBlock[blockIndex] || 0);
}

function buildMyTeamOOPIncrementSinglesBlock_(player, blockIndex) {
  if (!player.assignedSinglesByBlock) player.assignedSinglesByBlock = {};
  player.assignedSinglesByBlock[blockIndex] = buildMyTeamOOPSinglesCountInBlock_(player, blockIndex) + 1;
}

function buildMyTeamOOPIncrementDoublesBlock_(player, blockIndex) {
  if (!player.assignedDoublesByBlock) player.assignedDoublesByBlock = {};
  player.assignedDoublesByBlock[blockIndex] = buildMyTeamOOPDoublesCountInBlock_(player, blockIndex) + 1;
}

function getMyTeamOOPGameTemplateSheet_(ss) {
  const exactMatch = ss.getSheetByName('Game Template');
  if (exactMatch) return exactMatch;

  return ss.getSheets().find(sheet =>
    buildMyTeamOOPNormalizeName_(sheet.getName()) === 'game template'
    || buildMyTeamOOPNormalizeName_(sheet.getName()) === 'gametemplate'
  ) || null;
}