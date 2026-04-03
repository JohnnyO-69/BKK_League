function buildTheGame8BSeasonForm() {
  console.log('START buildTheGame8BSeasonForm');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fixturesSheet = ss.getSheetByName('Fixtures');
  const paramsSheet = ss.getSheetByName('Parameters');
  if (!fixturesSheet) {
    console.log('Fixtures sheet not found');
    return;
  }
  if (!paramsSheet) {
    console.log('Parameters sheet not found');
    return;
  }

  const sheetName = 'TheGame8BForm';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();

  const teamName = String(paramsSheet.getRange('F7').getValue() || '').trim();
  const teamId = Number(paramsSheet.getRange('F8').getValue());
  const delayMs = 400;

  console.log('Configured team name:', teamName);
  console.log('Configured team ID:', teamId);

  if (!teamName || Number.isNaN(teamId)) {
    console.log('Missing team name/id in Parameters!F7:F8');
    sheet.getRange('A1').setValue('Missing team name/id in Parameters!F7:F8');
    return;
  }

  const now = new Date();
  const sixWeeksMs = 42 * 24 * 60 * 60 * 1000;

  const fixtureData = fixturesSheet.getDataRange().getValues();
  console.log('Fixture rows including header:', fixtureData.length);

  const header = fixtureData.shift();
  const idxId = header.indexOf('Match ID');
  const idxHome = header.indexOf('Home Team');
  const idxAway = header.indexOf('Away Team');
  const idxDate = header.indexOf('Match Date');

  console.log('Fixture header:', header);
  console.log('Header indexes:', {
    idxId,
    idxHome,
    idxAway,
    idxDate
  });

  if ([idxId, idxHome, idxAway, idxDate].some(i => i < 0)) {
    console.log('Required Fixtures columns are missing');
    return;
  }

  function toDate(v) {
    if (v instanceof Date) return v;
    const d = new Date(String(v || '').trim());
    return isNaN(d) ? new Date(0) : d;
  }

  const matches = fixtureData
    .filter(r => (r[idxHome] === teamName || r[idxAway] === teamName) && r[idxId])
    .map(r => ({ id: r[idxId], date: toDate(r[idxDate]) }))
    .sort((a, b) => a.date - b.date);

  console.log('Matches found for team:', matches.length);
  console.log('First five matches:', matches.slice(0, 5));

  if (!matches.length) {
    console.log(`No matches found for ${teamName}`);
    sheet.getRange('A1').setValue(`No matches found for ${teamName}`);
    return;
  }

  const playerStats = {};

  function record(player, matchDate, won, isDouble) {
    if (!player) return;
    if (!playerStats[player]) {
      console.log('Creating player stats bucket for:', player);
      playerStats[player] = {
        player,
        singlesPlayed: 0,
        singlesWon: 0,
        doublesPlayed: 0,
        doublesWon: 0,
        points: 0,
        lastPlayed: new Date(0)
      };
    }
    const s = playerStats[player];

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

  function fetchJson(url) {
    try {
      console.log('Fetching match details:', url);
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = res.getResponseCode();
      console.log('Match details response code:', code);

      if (code !== 200) {
        console.log('Non-200 match details response body:', res.getContentText().slice(0, 500));
        return null;
      }

      const json = JSON.parse(res.getContentText());
      console.log('Match details payload keys:', Object.keys(json || {}));
      return json;
    } catch (error) {
      console.log('Error fetching/parsing match details:', error && error.message ? error.message : error);
      return null;
    }
  }

  matches.forEach(m => {
    console.log('Processing match:', m);

    const json = fetchJson(`https://api.bkkleague.com/match/details/${m.id}`);
    if (!json) {
      console.log('Skipping match because no JSON was returned:', m.id);
      return;
    }

    if (!Array.isArray(json.data)) {
      console.log('Skipping match because json.data is not an array:', m.id, json);
      return;
    }

    console.log('Frames returned for match:', m.id, json.data.length);

    json.data.forEach(f => {
      const hp = (f.homePlayers || []).map(p => p.nickName);
      const ap = (f.awayPlayers || []).map(p => p.nickName);
      const isDouble = hp.length > 1;

      const homeWon = f.homeWin === 1;
      const homeId = f.homeTeamId || f.home_team_id || f.homeTeamid;
      const awayId = f.awayTeamId || f.away_team_id || f.awayTeamid;

      const weHome = homeId === teamId;
      const weAway = awayId === teamId;

      if (!weHome && !weAway) {
        console.log('Skipping frame because configured team did not match frame teams:', {
          matchId: m.id,
          homeId,
          awayId,
          teamId
        });
        return;
      }

      const ourPlayers = weHome ? hp : ap;
      const won = weHome ? homeWon : !homeWon;

      console.log('Recording frame for our players:', {
        matchId: m.id,
        ourPlayers,
        isDouble,
        won,
        weHome,
        homeId,
        awayId
      });

      ourPlayers.forEach(p => record(p, m.date, won, isDouble));
    });

    Utilities.sleep(delayMs);
  });

  console.log('Player stats keys:', Object.keys(playerStats));
  console.log('Player stats snapshot:', JSON.stringify(playerStats, null, 2));

  const activeRows = Object.values(playerStats)
    .filter(p => now - p.lastPlayed <= sixWeeksMs)
    .map(s => {
      const totalGames = s.singlesPlayed + s.doublesPlayed;
      const ppg = totalGames ? s.points / totalGames : 0;
      const singlesPct = s.singlesPlayed ? (s.singlesWon / s.singlesPlayed) : 0;
      const doublesPct = s.doublesPlayed ? (s.doublesWon / s.doublesPlayed) : 0;

      return [
        s.player,
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

  console.log('Active rows before sort:', activeRows.length);
  console.log('Active row sample:', activeRows.slice(0, 5));

  // Sort by PPG desc → Points desc → Games desc → Name asc
  activeRows.sort((a, b) => {
    if (b[9] !== a[9]) return b[9] - a[9];
    if (b[8] !== a[8]) return b[8] - a[8];
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  });

  console.log('Active rows after sort:', activeRows.length);

  const headers = [
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

  // Write header
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (activeRows.length) {
    console.log('Writing active player rows to output sheet');

    sheet.getRange(2, 1, activeRows.length, headers.length).clearFormat();
    sheet.getRange(2, 1, activeRows.length, headers.length).setValues(activeRows);

    const startRow = 2;

    // Percent formatting
    sheet.getRange(startRow, 5, activeRows.length, 1).setNumberFormat("0.0%");
    sheet.getRange(startRow, 8, activeRows.length, 1).setNumberFormat("0.0%");

    // PPG formatting
    sheet.getRange(startRow, 10, activeRows.length, 1).setNumberFormat("0.00");

    // Center-align B → J
    sheet.getRange(startRow, 2, activeRows.length, 9).setHorizontalAlignment("center");

  } else {
    console.log('No active players in last 6 weeks');
    sheet.getRange(2, 1).setValue("No active players in last 6 weeks");
  }

  // Column widths
  for (let col = 1; col <= headers.length; col++) {
    sheet.setColumnWidth(col, 100);
  }

  // Border around data
  const fullRange = sheet.getRange(1, 1, activeRows.length + 1, headers.length);
  fullRange.setBorder(true, true, true, true, true, true);

  console.log('END buildTheGame8BSeasonForm');
}
