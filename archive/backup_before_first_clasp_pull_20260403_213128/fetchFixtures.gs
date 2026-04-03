function fetchDivision8BBFixtures() {

  console.log("START fetchDivision8BBFixtures");

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const params = ss.getSheetByName('Parameters');
  if (!params) {
    console.log("Parameters sheet not found");
    SpreadsheetApp.getUi().alert('Sheet "Parameters" not found');
    return;
  }

  let sheet = ss.getSheetByName('Fixtures');
  if (!sheet) {
    console.log("Creating Fixtures sheet");
    sheet = ss.insertSheet('Fixtures');
  }

  sheet.clear();

  console.log("Reading team table");

  const teamData = params.getRange('A9:B25').getValues();

  const teamIds = teamData
    .filter(r => r[1])
    .map(r => Number(r[1]))
    .filter(id => !isNaN(id));

  console.log("Team IDs loaded:", teamIds);

  if (teamIds.length === 0) {
    sheet.getRange('A1').setValue('No team IDs found in Parameters A9:B25');
    console.log("No team IDs found");
    return;
  }

  const sources = [
    {
      status: 'Pending',
      url: 'https://bkkleague.com/en/matches/pending'
    },
    {
      status: 'Completed',
      url: 'https://bkkleague.com/en/matches/completed'
    }
  ];

  const warnings = [];
  const matchesById = new Map();

  console.log("Fetching pending and completed match pages");

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(
      sources.map(source => ({
        url: source.url,
        muteHttpExceptions: true
      }))
    );
  } catch (e) {
    const msg = `Network error fetching match pages: ${e.message}`;
    sheet.getRange('A1').setValue(msg);
    console.log(msg);
    return;
  }

  responses.forEach((response, index) => {
    const source = sources[index];
    const code = response.getResponseCode();

    console.log(`${source.status} page response code:`, code);

    if (code !== 200) {
      warnings.push(`${source.status}: HTTP ${code}`);
      return;
    }

    try {
      const pageMatches = extractBkkLeagueMatchesFromHtml_(
        response.getContentText()
      );

      console.log(`${source.status} matches extracted:`, pageMatches.length);

      pageMatches.forEach(match => {
        const homeTeamId = Number(match.home_team_id);
        const awayTeamId = Number(match.away_team_id);

        if (
          teamIds.includes(homeTeamId) ||
          teamIds.includes(awayTeamId)
        ) {
          matchesById.set(String(match.match_id), {
            ...match,
            sourceStatus: source.status
          });
        }
      });
    } catch (e) {
      warnings.push(`${source.status}: ${e.message}`);
      console.log(`Failed to parse ${source.status} page: ${e.message}`);
    }
  });

  const headers = [
    'Match ID',
    'Status',
    'Match Date',
    'Home Team',
    'Away Team',
    'Home Frames',
    'Away Frames',
    'Home Team ID',
    'Away Team ID'
  ];

  const rows = Array.from(matchesById.values()).map(match => [
    match.match_id || '',
    match.sourceStatus || '',
    getBkkLeagueMatchDate_(match),
    match.home_team_name || '',
    match.away_team_name || '',
    match.home_frames ?? '',
    match.away_frames ?? '',
    match.home_team_id || '',
    match.away_team_id || ''
  ]);

  console.log("Matches collected:", rows.length);

  if (rows.length === 0) {
    const msg = warnings.length > 0
      ? `No matches found. ${warnings.join(' | ')}`
      : 'No matches found for configured teams';
    sheet.getRange('A1').setValue(msg);
    console.log("No matches matched configured teams");
    return;
  }

  console.log("Sorting rows");

  rows.sort((a, b) => new Date(a[1]) - new Date(b[1]));

  console.log("Writing data to sheet");

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  console.log("Applying formatting");

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#cfe2f3');

  sheet.autoResizeColumns(1, headers.length);

  sheet.getRange('J1').setValue('Last Fixtures Refresh');
  sheet.getRange('J2').setValue(
    new Date().toLocaleString('en-GB', { timeZone: 'Asia/Bangkok' })
  );

  if (warnings.length > 0) {
    sheet.getRange('J4').setValue('Warnings');
    sheet.getRange('J5').setValue(warnings.join(' | '));
  }

  console.log(`SUCCESS: Loaded ${rows.length} fixtures`);
}

function extractBkkLeagueMatchesFromHtml_(html) {
  const chunkPattern = /self\.__next_f\.push\(\[\d+,"((?:[^"\\]|\\.)*)"\]\)/g;
  const payloadChunks = [];
  let chunkMatch;

  while ((chunkMatch = chunkPattern.exec(html)) !== null) {
    payloadChunks.push(JSON.parse(`"${chunkMatch[1]}"`));
  }

  if (payloadChunks.length === 0) {
    throw new Error('No Next.js payload chunks found');
  }

  const payload = payloadChunks.join('');
  const matchesArrays = [];
  let searchIndex = 0;

  while (searchIndex < payload.length) {
    const keyIndex = payload.indexOf('"matches":[', searchIndex);
    if (keyIndex === -1) {
      break;
    }

    const arrayStart = payload.indexOf('[', keyIndex);
    const arrayText = extractBalancedJsonArray_(payload, arrayStart);
    matchesArrays.push(JSON.parse(arrayText));
    searchIndex = arrayStart + arrayText.length;
  }

  if (matchesArrays.length === 0) {
    throw new Error('No embedded matches arrays found');
  }

  const matchesById = new Map();

  matchesArrays.forEach(matches => {
    matches.forEach(match => {
      if (match && typeof match === 'object' && match.match_id) {
        matchesById.set(String(match.match_id), match);
      }
    });
  });

  return Array.from(matchesById.values());
}

function extractBalancedJsonArray_(text, startIndex) {
  let depth = 0;
  let inString = false;
  let isEscaped = false;

  for (let index = startIndex; index < text.length; index++) {
    const char = text[index];

    if (isEscaped) {
      isEscaped = false;
      continue;
    }

    if (char === '\\' && inString) {
      isEscaped = true;
      continue;
    }

    if (char === '"') {
      inString = !inString;
      continue;
    }

    if (inString) {
      continue;
    }

    if (char === '[') {
      depth += 1;
      continue;
    }

    if (char === ']') {
      depth -= 1;
      if (depth === 0) {
        return text.slice(startIndex, index + 1);
      }
    }
  }

  throw new Error('Could not extract matches array from payload');
}

function getBkkLeagueMatchDate_(match) {
  const rawDate = String(match.match_date || match.date || '');
  return rawDate.replace(/^\$D/, '').split('T')[0].split(' ')[0];
}