# Deprecated Scripts

These files were archived locally to keep them in the repository and Git history, while excluding them from the deployable Apps Script source set.

## Archived Files

- `buildCurrentMatchFormSmart.js`
  - Archived because `buildCurrentMatchFormSmartV2.js` is the active current-match implementation in the deployable source set.
  - The old version was hardcoded to `The Game 8B` and team ID `2327`, while V2 reads team configuration from `Parameters!F7:F8`.
  - After trigger consolidation, no remaining active deployable callers required the old implementation.

- `buildSeasonForm.js`
  - Archived because the active workflow uses `buildSeasonFormForTeam.js`.
  - No current in-repo callers were found for `buildSeasonForm()`.

- `buildPlayerLast3MatchForm.js`
  - Archived because the active workflow uses `buildLast3FormForTeam.js`.
  - No current in-repo callers were found for `buildPlayerLast3MatchForm()`.

- `buildLast3FormWithHeadings.js`
  - Archived because it duplicated wrapper-style output logic that was not part of the current core automation chain.
  - No current in-repo callers were found from the active deployable path.

- `Last3GamesHeader.js`
  - Archived because it duplicated global helper and wrapper functions already represented elsewhere.
  - Keeping it in production would increase ambiguity in the Apps Script global namespace without a confirmed active use.

## Important Note

These scripts are archived from local deployable source only.

They are not part of the current `clasp` tracked production set unless they are restored to the workspace root and intentionally redeployed.
