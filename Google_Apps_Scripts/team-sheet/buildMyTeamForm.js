function buildMyTeamForm() {
  console.log('START buildMyTeamForm');

  const ss = getLeagueSpreadsheet_();
  const paramsSheet = ss.getSheetByName('Parameters');

  if (!paramsSheet) {
    console.log('Parameters sheet not found');
    return;
  }

  const teamName = String(paramsSheet.getRange('F7').getValue() || '').trim();
  const teamId = Number(paramsSheet.getRange('F8').getValue());

  console.log('Configured team name:', teamName);
  console.log('Configured team ID:', teamId);

  if (!teamName || Number.isNaN(teamId)) {
    console.log('Missing team name/id in Parameters!F7:F8');
    return;
  }

  buildLast3FormForTeam(teamId, teamName, 'Last3MatchForm');
  const frameCache = buildSeasonFormForTeam(teamId, teamName, 'SeasonForm');
  buildMyTeamOOP();
  buildMyTeamPairings(frameCache);

  console.log('END buildMyTeamForm');
}