# BKK League Apps Script Repo

This repository is the source-controlled local home for the Bangkok Pool League Google Apps Script project.

## Scope

This repo manages:

- Apps Script source files
- the Apps Script manifest
- the explicit target spreadsheet binding
- local `clasp` workflow and verification tooling

This repo does not version spreadsheet data or formatting unless scripts recreate those elements.

## Linked Apps Script Projects

- BKK League Data: `1gP5BJz1Lpz3-XLQmiKYIi_y0nARwfpED8zjoyhK6gO1OgBAvtYllamCt`
- Team Sheet: `1CWO4vZaW5FTQ9yRjghI4zg3yryKUt0jgW-y7tT8lLHxffNSA3Hg31vKl`

## Target Spreadsheets

- BKK League Data spreadsheet: `1Kcv1y5bQX8YGxnIIXyKYj5QkcQQO_qBO5Zcvt6lSMAU`
- Team Sheet spreadsheet: `1zz1rk8E_r3dDxkMUR1_30igrPXKVZu7ju0SkRMaYJf4`
- Config source: `projectConfig.js` in each project folder

## Canonical Deployable Source

The deployable Apps Script source lives under:

- `Google_Apps_Scripts/bkk-league-data/`
- `Google_Apps_Scripts/team-sheet/`

Archived or deprecated scripts are kept under `archive/` and are not part of the production deployment set.

## Active Entry Points

The current core flow is built around these functions:

- `fetchDivision8BBFixtures()` in `fetchFixtures.js`
- `fetchAllBKKLeagueData()` in `fetchFixtures.js` as a compatibility alias for legacy triggers
- `buildMyTeamForm()` in `buildMyTeamForm.js`
- `buildNextTeamForm()` in `buildNextTeamForm.js`
- `buildLast3FormForTeam()` in `buildLast3FormForTeam.js`
- `buildSeasonFormForTeam()` in `buildSeasonFormForTeam.js`
- `buildMyTeamOOP()` in `buildMyTeamOOP.js`
- `buildNextTeamOOP()` in `buildNextTeamOOP.js`
- `buildCurrentMatchFormSmart()` in `buildCurrentMatchFormSmart.js`
- `refreshPrediction()` in `refreshPrediction.js`

## Trigger Design

Current edit-trigger behavior is centralized in `onEdit_Trigger.js`.

The active current-match builder is:

- `buildCurrentMatchFormSmart()` in `buildCurrentMatchFormSmart.js`

The trigger file provides:

- a global `onEdit(e)` simple trigger
- `onEdit_Trigger(e)` as the shared dispatcher
- a dispatch to the canonical current-match refresh helper

This avoids conflicting global `onEdit(e)` definitions in Apps Script's shared global namespace.

The legacy hardcoded implementation was archived, and the former V2 implementation was promoted into the canonical `buildCurrentMatchFormSmart()` path.

## Local Workflow

Install dependencies:

```powershell
npm install
```

Verify source-of-truth alignment:

```powershell
npm run sync:check
```

Push Apps Script code:

```powershell
npm run clasp:push
```

Push only one project:

```powershell
npm run clasp:push:data
npm run clasp:push:team
```

Pull Apps Script code:

```powershell
npm run clasp:pull
```

## Git Workflow

The active Git branch is `main`.

Suggested flow:

1. Edit code locally.
2. Run `npm run sync:check`.
3. Run `npm run clasp:push` if Apps Script deployment is intended.
4. Commit and push to GitHub.

## Archive Policy

Files under `archive/` are intentionally preserved for reference and Git history.

They are:

- not production deployable
- not part of the canonical active source set
- available for recovery, comparison, or audit

## Fixtures Refresh Rule

Both script projects keep a local `fetchFixtures.js` because both spreadsheets read their own `Fixtures` tab.

- `Google_Apps_Scripts/bkk-league-data/fetchFixtures.js`
- `Google_Apps_Scripts/team-sheet/fetchFixtures.js`

`npm run verify:source` now fails if those two files drift apart.
