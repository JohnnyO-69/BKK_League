# Google Apps Script Workflow

This workspace root is the source-controlled Apps Script project.

## Linked Script Project

- Script ID: `1CWO4vZaW5FTQ9yRjghI4zg3yryKUt0jgW-y7tT8lLHxffNSA3Hg31vKl`
- Config file: `.clasp.json`
- Manifest: `appsscript.json`

## Target Spreadsheet

- Spreadsheet ID: `1Kcv1y5bQX8YGxnIIXyKYj5QkcQQO_qBO5Zcvt6lSMAU`
- Spreadsheet URL: `https://docs.google.com/spreadsheets/d/1Kcv1y5bQX8YGxnIIXyKYj5QkcQQO_qBO5Zcvt6lSMAU/edit?gid=680648455#gid=680648455`
- Binding file: `projectConfig.js`

All canonical Apps Script files now resolve the workbook through `getLeagueSpreadsheet_()` instead of `SpreadsheetApp.getActiveSpreadsheet()`.

## Managed Files

The canonical deployable files are the root `*.js` Apps Script source files pulled from the linked project, plus the manifest.

Deployable through `clasp` from this workspace:

- root `*.js` files
- `appsscript.json`

Included in the source of truth:

- Apps Script code
- manifest configuration
- explicit target spreadsheet ID in `projectConfig.js`

Not included in the source of truth:

- spreadsheet data values
- sheet formatting
- sheet tabs, filters, protections, charts, and formulas unless they are recreated by script
- manual spreadsheet edits outside script-managed behavior

Excluded from deployment:

- `_repo_analysis/`
- `Google_Apps_Scripts/`
- `archive/`
- Markdown docs
- local tool config not required by Apps Script

## Migration Note

The first remote `clasp pull` showed that the linked Apps Script project stores script files as `.js`, not `.gs`.

Local pre-pull `.gs` files and first-pull backups were archived and should be treated as historical references only, not active source.

## First-Time Setup

Install dependencies:

```powershell
npm install
```

Authenticate with Google:

```powershell
npm run clasp:login
```

Check the linked project status:

```powershell
npm run clasp:status
```

## Daily Workflow

Verify the repo still follows the single-source model:

```powershell
npm run sync:check
```

Push local code to Apps Script:

```powershell
npm run clasp:push
```

Pull remote code from Apps Script:

```powershell
npm run clasp:pull
```

Open the linked Apps Script project in the browser:

```powershell
npm run clasp:open
```

## Recommended Source-Of-Truth Rule

To keep this workspace as the canonical source:

1. Edit code locally in this repository.
2. Edit the root `*.js` Apps Script files, not archived legacy `.gs` copies.
3. Run `npm run sync:check` before pushing.
4. Push changes with `npm run clasp:push`.
5. Avoid editing production code directly in the Apps Script browser editor unless necessary.
6. If an emergency browser edit is made, immediately run `npm run clasp:pull` and commit the result locally.

## Safe Sync Rule

Before pushing:

1. Run `npm run sync:check`.
2. Confirm only intended root `*.js` files and `appsscript.json` are tracked.
3. Push with `npm run clasp:push`.

## Notes

- `clasp` is installed locally in `devDependencies`, so `npx` is not required after `npm install`.
- The workspace root is configured as the `clasp` root directory.
- The current `.claspignore` is intentionally strict so only Apps Script source files deploy.
- `npm run verify:source` fails fast if root-level `*.gs` files, `Copy ...` scripts, or `.gs.js` files reappear.
- `projectConfig.js` is part of the canonical deployable source because it fixes the workbook target explicitly.
