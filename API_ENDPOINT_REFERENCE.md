# Bangkok Pool League Endpoint Reference

Last updated: 2026-04-03

## Scope

This document consolidates the endpoint surface that is visible from:

- The public mobile client repo: `khkwan0/bangkok_pool_league.orig`
- The public website: `https://bkkleague.com`

This is not a backend source-of-truth. The backend code is not public in the inspected repo, so this file should be treated as a client-observed API reference.

## Confidence Levels

- Confirmed: directly referenced in client code or public website output
- Inferred: strongly suggested by route patterns or visible behavior, but not directly defined by backend source
- Unknown: response schema or auth rules cannot be fully proven from available sources

## Base Hosts

- Web app: `https://bkkleague.com`
- API/media host: `https://api.bkkleague.com`
- Mobile client code uses `https://<config.domain>/<endpoint>` with Bearer JWT auth for most requests

## Auth Model

Observed from the mobile client:

- JWT is stored in AsyncStorage under `jwt`
- Most HTTP requests send `Authorization: Bearer <jwt>`
- Some auth routes may accept unauthenticated requests despite the generic wrapper

## HTTP API Endpoints

### Account and Auth

| Method | Path | Status | Request Data | Response Data Notes |
| --- | --- | --- | --- | --- |
| GET | `/user` | Confirmed | none | Returns current user object. Observed field: `role_id`. |
| POST | `/user/token` | Confirmed | `{ token }` | Likely saves push notification token. |
| POST | `/login` | Confirmed | `{ email, password }` | On success: `{ status: 'ok', data: { token, user } }`. |
| POST | `/admin/login` | Confirmed | `{ playerId }` | On success: `{ status: 'ok', data: { token, user } }`. |
| GET | `/logout` | Confirmed | none | Unknown schema. |
| POST | `/login/social/:platform` | Confirmed | `{ data }` | On success: `{ status: 'ok', data: { token, user } }`. |
| POST | `/login/register` | Confirmed | `{ email, password1, password2, nickname, firstName, lastName }` | Unknown schema. |
| POST | `/login/recover` | Confirmed | `{ email }` | Used by forgot password flow. |
| POST | `/login/recover/verify` | Confirmed | `{ code, password, passwordConfirm }` | Completes password reset. |
| GET | `/account/delete` | Confirmed | none | Deletes current account or initiates delete workflow. |
| POST | `/account/first_name` | Confirmed | `{ name }` | Updates first name. |
| POST | `/account/last_name` | Confirmed | `{ name }` | Updates last name. |
| POST | `/account/nick_name` | Confirmed | `{ name }` | Updates nickname. |
| POST | `/avatar` | Confirmed | multipart form with `photo` | Returns upload result; exact schema unknown. |

### Seasons, League, Standings

| Method | Path | Status | Request Data | Response Data Notes |
| --- | --- | --- | --- | --- |
| GET | `/season` | Confirmed | none | Current season data. |
| GET | `/v2/season` | Confirmed | none | Versioned season endpoint. |
| GET | `/seasons` | Confirmed | none | List of seasons. |
| POST | `/admin/season/new` | Confirmed | `{ name, shortName, description }` | Creates a season. |
| GET | `/admin/season/activate/:seasonId` | Confirmed | none | State-changing action implemented as GET. |
| POST | `/admin/migrate` | Confirmed | `{ oldSeason, newSeason }` | Migrates season data. |
| GET | `/league/standings/:seasonId` | Confirmed | none | Standings by season. Public website shows division standings data. |
| GET | `/league/season/:seasonId/division/player/stats` | Confirmed | none | Player stats grouped by division. |

### Players

| Method | Path | Status | Request Data | Response Data Notes |
| --- | --- | --- | --- | --- |
| GET | `/player/:playerId` | Confirmed | none | Player profile/info. |
| GET | `/player/raw/:playerId` | Confirmed | none | Raw player info. Likely admin-oriented. |
| GET | `/player/stats/info/:playerId` | Confirmed | none | Stats summary metadata for player. |
| GET | `/players?active_only=:bool` | Confirmed | none | Player list with active filter. |
| GET | `/players/unique` | Confirmed | none | Deduplicated player list. |
| GET | `/players/all` | Confirmed | none | Full player list, likely admin use. |
| POST | `/player` | Confirmed | `{ nickName, firstName, lastName, email, teamId }` | Creates a new player. |
| POST | `/admin/player/attribute` | Confirmed | `{ playerId, key, value }` | Updates arbitrary player attribute. |
| POST | `/player/privilege/grant` | Confirmed | `{ playerId, teamId, level }` | Grants player privilege. |
| POST | `/player/privilege/revoke` | Confirmed | `{ playerId, teamId }` | Revokes privilege. |
| GET | `/users/mergerequest/count` | Confirmed | none | Count of active merge requests. |
| GET | `/mymergerequests` | Confirmed | none | Current user's merge requests. |
| GET | `/admin/mergerequests` | Confirmed | none | Admin merge request list. |
| GET | `/admin/mergerequest/accept/:requestId` | Confirmed | none | Accepts merge request. Implemented as GET. |
| GET | `/admin/mergerequest/deny/:requestId` | Confirmed | none | Denies merge request. Implemented as GET. |
| GET | `/admin/users/merge/:currentId/:toMergeId` | Confirmed | none | Merges users/players. Implemented as GET. |

### Teams, Divisions, Venues

| Method | Path | Status | Request Data | Response Data Notes |
| --- | --- | --- | --- | --- |
| GET | `/teams` | Confirmed | none | Team list. |
| GET | `/teams/:season` | Confirmed | none | Teams by season. |
| GET | `/admin/teams/:season` | Confirmed | none | Team list for admin workflows. |
| GET | `/team/:teamId` | Confirmed | none | Team details. |
| POST | `/admin/team` | Confirmed | `{ name, venue }` | Creates a team. |
| POST | `/admin/team/division` | Confirmed | `{ teamId, divisionId }` | Assigns team to division. |
| GET | `/team/division/:seasonId` | Confirmed | none | Team-division mapping for season. |
| GET | `/divisions/:season` | Confirmed | none | Division list for season. |
| POST | `/team/player` | Confirmed | `{ playerId, teamId }` | Adds player to team. |
| POST | `/team/player/remove` | Confirmed | `{ playerId, teamId }` | Removes player from team. |
| GET | `/playersteam/players?teamid=:teamId&active_only=:bool` | Confirmed | none | Team roster. |
| GET | `/venues` | Confirmed | none | Venue list. |
| GET | `/venues/all` | Confirmed | none | Full venue list. |
| POST | `/venue` | Confirmed | `{ venue }` | Creates or saves venue data. |

### Matches, Frames, Schedules, Scores

| Method | Path | Status | Request Data | Response Data Notes |
| --- | --- | --- | --- | --- |
| GET | `/matches?{query}` | Confirmed | query string | General match list endpoint. |
| GET | `/season/matches?season=:season` | Confirmed | none | Matches by season. |
| GET | `/matches/season/:seasonId` | Confirmed | none | Alternate season match endpoint. |
| GET | `/matches/postponed` | Confirmed | none | Postponed matches. |
| GET | `/v2/matches/completed/season/:season` | Confirmed | none | Completed matches by season. |
| GET | `/match/:matchId` | Confirmed | none | Match info. |
| GET | `/match/details/:matchId` | Confirmed | none | Detailed match payload. |
| GET | `/match/info/full/:matchId` | Confirmed | none | Full match info. Likely richest match payload. |
| GET | `/match/stats/:matchId` | Confirmed | none | Match stats. |
| GET | `/frames/:matchId` | Confirmed | none | Match frames. |
| POST | `/admin/match/completed` | Confirmed | `{ type, matchId, data }` | Updates completed match state. |
| POST | `/admin/match/date` | Confirmed | `{ matchId, newDate }` | Reschedules match. |
| GET | `/scores/live` | Confirmed | none | Live scores for public display and app. |

### Statistics and Reference Data

| Method | Path | Status | Request Data | Response Data Notes |
| --- | --- | --- | --- | --- |
| GET | `/stats?playerid=:playerId` | Confirmed | none | General player stats. |
| GET | `/stats/match?playerid=:playerId` | Confirmed | none | Match performance stats. |
| GET | `/stats/doubles?playerid=:playerId` | Confirmed | none | Doubles stats. |
| GET | `/stats/players/:seasonId?minimum=:n&gameType=:type&singles=1|doubles=1` | Confirmed | none | Filterable player stats table. |
| GET | `/stats/teams/:seasonId` | Confirmed | none | Team stats. |
| GET | `/stats/team/players/internal/:teamId` | Confirmed | none | Internal team player stats. |
| GET | `/game/types` | Confirmed | none | Game types list. |
| GET | `/gametypes` | Confirmed | none | Alternate game type endpoint. |
| GET | `/rules` | Confirmed | none | League rules data. |

## WebSocket Events

Observed Socket.IO events from the mobile client and public TV/live pages.

### Client Emits

| Event | Status | Payload Notes |
| --- | --- | --- |
| `join` | Confirmed | Emits room id, server returns join status via callback. |
| `matchupdate` | Confirmed | Payload shape below. |

Observed `matchupdate` payload fields:

```json
{
  "type": "string",
  "matchId": 0,
  "timestamp": 0,
  "playerId": 0,
  "jwt": "string",
  "nickname": "string",
  "dest": "string",
  "data": {}
}
```

Observed `type` usage:

- `newnote`

### Client Receives

| Event | Status | Notes |
| --- | --- | --- |
| `match_update` | Confirmed | Match state updates. |
| `frame_update` | Confirmed | Frame state updates. |
| `historyupdate` | Confirmed | Match history updates. |
| `historyupdate2` | Confirmed | Alternate history update stream. |
| `match_update2` | Confirmed | Alternate match update stream. |
| `connect` | Confirmed | Standard Socket.IO lifecycle event. |
| `disconnect` | Confirmed | Standard Socket.IO lifecycle event. |
| `reconnect` | Confirmed | Standard Socket.IO lifecycle event. |

## Public Website Routes

These are public-facing web routes on `https://bkkleague.com` rather than documented JSON API endpoints.

### English Routes

- `/en`
- `/en/auth`
- `/en/auth/forgot-password`
- `/en/announcements`
- `/en/tv`
- `/en/matches/pending`
- `/en/matches/completed`
- `/en/matches/postponed`
- `/en/matches/by-team`
- `/en/matches/by-venue`
- `/en/league-standings`

### Thai Routes

- `/th`
- `/th/auth`
- `/th/announcements`
- `/th/tv`
- `/th/matches/pending`
- `/th/matches/completed`
- `/th/matches/postponed`
- `/th/matches/by-team`
- `/th/matches/by-venue`
- `/th/league-standings`

## Publicly Visible Website Data

The public website exposes at least the following categories of data without requiring the backend source code:

- announcements
- season labels
- standings by division
- upcoming matches calendar
- completed matches calendar
- postponed matches listing
- matches by team
- matches by venue
- venue names, addresses, and apparent table counts
- live TV schedule data
- team logos hosted on `api.bkkleague.com`

## Static and Media Endpoints

These are public URL patterns observed from the website.

| Path Pattern | Host | Status | Notes |
| --- | --- | --- | --- |
| `/logos/:file` | `api.bkkleague.com` | Confirmed | Team and venue logos. |
| `/announcements_gallery/:file` | `bkkleague.com` | Confirmed | Announcement images. |
| `/line-logo.svg` | `bkkleague.com` | Confirmed | Login page asset. |

## Protected or Hidden Route Areas

Observed from `robots.txt`:

- `/api/`
- `/dash/`
- `/members/`
- `/auth/`
- `/admin/`

These indicate additional non-public route areas likely exist on the web platform, but specific endpoint paths under these prefixes were not enumerable from the public frontend alone.

## Known Data Shapes

### Login Success

```json
{
  "status": "ok",
  "data": {
    "token": "jwt",
    "user": {}
  }
}
```

### User Object

Known fields observed in code or comments:

```json
{
  "id": 1,
  "email": "user@example.com",
  "firstName": "First",
  "lastName": "Last",
  "role_id": 9
}
```

Additional likely user fields:

- `nickname`
- team or privilege-related fields

### Venue Data Visible On Website

Observed website-visible venue fields:

```json
{
  "name": "The Sportsman",
  "address": "Sukhumvit Soi 13",
  "matchCount": 78,
  "tableCount": 3,
  "logoUrl": "https://api.bkkleague.com/logos/sportsman.png"
}
```

### Match Data Visible On Website

Observed website-visible match fields:

```json
{
  "date": "2026-04-06",
  "division": "8B B",
  "homeTeam": "Breaking Bad",
  "awayTeam": "C.O. Jagerbombers",
  "venueName": "Breakers Sports Bar",
  "venueAddress": "Sukhumvit Soi 13",
  "score": "14 - 6"
}
```

Not all pages expose all fields at once.

## Notable Design Findings

- Several state-changing admin operations use `GET` instead of `POST`, `PUT`, or `DELETE`
- There are duplicate or versioned endpoints for similar resources, for example:
  - `/season` and `/v2/season`
  - `/game/types` and `/gametypes`
  - `/season/matches?season=:id` and `/matches/season/:id`
- The backend likely serves both JSON API responses and static/media assets from `api.bkkleague.com`
- The website and mobile app likely share at least part of the same backend data model

## Gaps

The following remain unknown without backend access:

- full response schemas for most endpoints
- validation rules
- exact auth requirements per endpoint
- hidden `/api/*` routes used by the website frontend
- admin dashboard-only endpoints
- whether any GraphQL or server actions exist behind the web frontend

## Suggested Maintenance

When new endpoints are discovered, record them with:

1. method
2. path
3. auth requirement
4. request shape
5. response shape
6. source of evidence

Add a `Source` column if this file is later expanded into a stricter audit inventory.