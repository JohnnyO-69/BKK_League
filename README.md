# BKK League Apps Script Repo

This repository is the source-controlled local home for the Bangkok Pool League Google Apps Script project.

## Scope

This repo manages:

- Apps Script source files
- the Apps Script manifest
- the explicit target spreadsheet binding
- local `clasp` workflow and verification tooling

This repo does not version spreadsheet data or formatting unless scripts recreate those elements.

## Linked Apps Script Project

- Script ID: `1CWO4vZaW5FTQ9yRjghI4zg3yryKUt0jgW-y7tT8lLHxffNSA3Hg31vKl`

## Target Spreadsheet

- Spreadsheet ID: `1Kcv1y5bQX8YGxnIIXyKYj5QkcQQO_qBO5Zcvt6lSMAU`
- Config source: `projectConfig.js`

## Canonical Deployable Source

The deployable Apps Script source is the root `*.js` file set plus `appsscript.json`.

Archived or deprecated scripts are kept under `archive/` and are not part of the production deployment set.

## Active Entry Points

The current core flow is built around these functions:

- `fetchDivision8BBFixtures()` in `fetchFixtures.js`
- `buildMyTeamForm()` in `buildMyTeamForm.js`
- `buildNextTeamForm()` in `buildNextTeamForm.js`
- `buildLast3FormForTeam()` in `buildLast3FormForTeam.js`
- `buildSeasonFormForTeam()` in `buildSeasonFormForTeam.js`
- `buildMyTeamOOP()` in `buildMyTeamOOP.js`
- `buildNextTeamOOP()` in `buildNextTeamOOP.js`
- `buildCurrentMatchFormSmartV2()` in `buildCurrentMatchFormSmartV2.js`
- `refreshPrediction()` in `refreshPrediction.js`

## Trigger Design

Current edit-trigger behavior is centralized in `onEdit_Trigger.js`.

The active current-match builder is:

- `buildCurrentMatchFormSmartV2()` in `buildCurrentMatchFormSmartV2.js`

The trigger file provides:

- a global `onEdit(e)` simple trigger
- `onEdit_Trigger(e)` as the shared dispatcher
- a dispatch to the V2 current-match refresh helper

This avoids conflicting global `onEdit(e)` definitions in Apps Script's shared global namespace.

The previous non-V2 implementation was archived because it was hardcoded to a specific team and had no remaining in-repo callers once trigger handling was centralized.

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
