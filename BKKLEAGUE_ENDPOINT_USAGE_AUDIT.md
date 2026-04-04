# BKK League Endpoint Usage Audit

Last updated: 2026-04-04

## Purpose

This document answers two practical questions for this repository:

1. Which BKK League URLs are actually used by the current deployable Apps Script code?
2. Do any active scripts appear to use outdated or mismatched BKK League paths?

This is a repo-usage audit, not a backend source-of-truth.

## Conclusion

The active deployable scripts use a small subset of the known BKK League surface:

- one API endpoint pattern on `api.bkkleague.com`
- two public website routes on `bkkleague.com`

No outdated BKK League API paths were found in the active root deployable scripts.

## Active Deployable Usage

### 1. Public Website Routes Used by Production Scripts

These are scraped from the public website and used to build the `Fixtures` sheet.

| Host | Method | Path | Used By | Purpose | Status |
| --- | --- | --- | --- | --- | --- |
| `bkkleague.com` | GET | `/en/matches/pending` | `fetchDivision8BBFixtures()` in `fetchFixtures.js` | Pull pending match listings from the public site | Active |
| `bkkleague.com` | GET | `/en/matches/completed` | `fetchDivision8BBFixtures()` in `fetchFixtures.js` | Pull completed match listings from the public site | Active |

Notes:

- These calls do not use a JSON API endpoint.
- They depend on the public Next.js page payload structure remaining stable.
- This is not an outdated path issue, but it is more brittle than calling a dedicated API.

### 2. API Endpoints Used by Production Scripts

These are called directly on the BKK League API host.

| Host | Method | Path Pattern | Used By | Purpose | Status |
| --- | --- | --- | --- | --- | --- |
| `api.bkkleague.com` | GET | `/match/details/:matchId` | `buildCurrentMatchFormSmart()` in `buildCurrentMatchFormSmart.js` | Load detailed frame data for the current match form | Active |
| `api.bkkleague.com` | GET | `/match/details/:matchId` | `buildLast3FormForTeam()` in `buildLast3FormForTeam.js` | Load frame data for recent-match player stats | Active |
| `api.bkkleague.com` | GET | `/match/details/:matchId` | `buildSeasonFormForTeam()` in `buildSeasonFormForTeam.js` | Load frame data for season-form stats | Active |
| `api.bkkleague.com` | GET | `/match/details/:matchId` | `buildNextTeamOOP()` in `buildNextTeamOOP.js` | Load opponent match details for lineup and slot analysis | Active |

## Active Root Scripts With No Direct BKK League HTTP Calls

These deployable root scripts currently do not fetch BKK League URLs directly:

- `buildMyTeamForm.js`
- `buildMyTeamOOP.js`
- `buildNextTeamForm.js`
- `onEdit_Trigger.js`
- `projectConfig.js`
- `refreshPrediction.js`

They either orchestrate other scripts, read spreadsheet data, or handle local trigger logic.

## Comparison Against API_ENDPOINT_REFERENCE.md

Compared with the broader reference in `API_ENDPOINT_REFERENCE.md`, this Apps Script repo uses only:

- `/match/details/:matchId`
- `/en/matches/pending`
- `/en/matches/completed`

It does not currently call other known BKK League endpoints such as:

- `/matches`
- `/match/:matchId`
- `/match/info/full/:matchId`
- `/frames/:matchId`
- `/scores/live`
- `/league/standings/:seasonId`
- `/teams`
- `/player/:playerId`

## Outdated Path Review

### Finding: No outdated BKK League paths found in active deployable scripts

The current root deployable scripts are consistent with the observed endpoint reference.

No active root script was found using:

- old hostnames
- dead-looking version prefixes
- mismatched mobile-only auth routes
- archived duplicate endpoints when a newer active equivalent is already used

### Important nuance

Although no outdated paths were found, the fixture importer is structurally more fragile than the API-based scripts because it parses public website HTML and embedded Next.js payload data instead of using a stable JSON API contract.

That means the higher operational risk is:

- page structure drift on `bkkleague.com`

not:

- an obviously outdated URL path in this repo

## Archived References

Archived scripts under `archive/` still contain historical references to the same API path pattern, especially:

- `/match/details/:matchId`

Those are not part of the current deployable source set and should not be treated as active production dependencies.

## Practical Summary

- The Git repo does not mirror the whole BKK League website or backend.
- The active Apps Script production code currently depends on only three BKK League routes.
- The active direct API dependency is limited to `/match/details/:matchId`.
- No outdated active endpoint usage was found.
- The main maintenance risk is HTML scraping from the two public match pages.
