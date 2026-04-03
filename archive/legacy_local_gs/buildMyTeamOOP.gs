function buildMyTeamOOP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName('Last3MatchForm');
  const paramsSheet = ss.getSheetByName('Parameters');
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

  const playerPool = buildMyTeamOOPVisiblePlayerPool_(source, headerMap);
  if (!playerPool.length) {
    out.getRange('A1').setValue('No visible filtered players found in Last3MatchForm A7:A19');
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
    .setValue('Players derived from visible filtered rows in Last3MatchForm!A7:A19')
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

function buildMyTeamOOPVisiblePlayerPool_(sheet, headerMap) {
  const rows = sheet.getRange(7, 1, 13, Math.min(sheet.getLastColumn(), 20)).getValues();
  const players = [];

  rows.forEach((row, index) => {
    const rowNumber = index + 7;
    if (sheet.isRowHiddenByFilter(rowNumber) || sheet.isRowHiddenByUser(rowNumber)) return;

    const name = String(row[headerMap.player] || '').trim();
    if (!name) return;

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
      assignedDoubles: 0
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

    if (format === 'Singles') {
      const choice = buildMyTeamOOPChooseSingles_(playerPool, recentPlayers);
      if (!choice) return;

      choice.player.assignedSingles += 1;
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

    const pairChoice = buildMyTeamOOPChooseDoubles_(playerPool, recentPlayers, recentPairs);
    if (!pairChoice) return;

    pairChoice.players.forEach(player => {
      player.assignedDoubles += 1;
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

  return lineup;
}

function buildMyTeamOOPChooseSingles_(playerPool, recentPlayers) {
  const candidates = playerPool
    .map(player => {
      let score = player.singlesWeight;
      const mostRecent = recentPlayers[0] || [];
      const previous = recentPlayers[1] || [];

      if (mostRecent.includes(player.name)) score -= 0.18;
      if (previous.includes(player.name)) score -= 0.08;
      score -= player.assignedSingles * 0.07;

      return { player, score };
    })
    .sort((a, b) => b.score - a.score || a.player.name.localeCompare(b.player.name));

  return candidates[0] || null;
}

function buildMyTeamOOPChooseDoubles_(playerPool, recentPlayers, recentPairs) {
  const pairs = [];

  for (let i = 0; i < playerPool.length; i++) {
    for (let j = i + 1; j < playerPool.length; j++) {
      const first = playerPool[i];
      const second = playerPool[j];
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