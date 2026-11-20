# openteams-server

The reference bring-your-own server for **OpenTeams**, an open-source team workspace. One small
Node process does three jobs:

- `ws /sync/<room>`: a y-websocket compatible Yjs document and awareness server, persisted to disk;
- `ws /signal/<room>`: a JSON relay for WebRTC call signaling;
- `POST|GET|DELETE /files/<scope>/<name>`, `GET /files/<scope>`: file storage for attachments (25 MB cap), scoped
  by workspace and by private chat, with short-lived signed download links;
- `GET /metrics`: Prometheus metrics, only when `TEAMS_METRICS_TOKEN` is set.

It understands neither chat nor Office formats: it moves bytes, so any server that speaks the same
three contracts works with the OpenTeams bindings just as well.

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme) &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; **[OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme)** &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

## Run

Needs Node 22.3 or later.

```bash
openssl rand -hex 32 > session-secret    # once; keep it outside the repository
TEAMS_USERS_FILE=users.json TEAMS_SESSION_SECRET="$(cat session-secret)" npx openteams-server   # PORT=8787 HOST=127.0.0.1 by default
npx openteams-server --insecure-dev     # local development only: no accounts, ids are claims
```

**The server is secure by default.** With no `TEAMS_USERS_FILE` (or `users`), no `authenticate`
hook and no `authorizeSync` hook it refuses to start, because every client could otherwise claim any
user id. `--insecure-dev` (option `insecureDev: true`) opts out, logs a loud warning and is meant
for your own machine. A shared `TEAMS_TOKEN` alone does not count as authentication. A server with
only `authorizeSync` starts, but offers sync alone: without a way to check who someone is, it
answers `/auth/session`, `/files` and `/signal` with 404 (see Per-room permissions).

Flags: `--insecure-dev`, `--insecure` (see `TEAMS_INSECURE`) and `--trust-proxy[=N]`. `SIGTERM` and
`SIGINT` save every room before the process exits.

| Variable                                                     | Meaning                                                                                                                                                                                                                                               |
| ------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`                                               | Listen address. Bind `0.0.0.0` only behind TLS.                                                                                                                                                                                                       |
| `TEAMS_TOKEN`                                                | If set, sockets take it in the `Sec-WebSocket-Protocol` header (`?token=` only as a fallback) and file requests as `Authorization: Bearer`.                                                                                                           |
| `TEAMS_USERS_FILE`, `TEAMS_SESSION_SECRET`                   | Path to a JSON object of user id to credential, or to `{ credential, role }` (enables accounts); the secret that signs sessions (at least 32 characters with 8 distinct ones, or unset for a random one per process; a shorter one stops the server). |
| `TEAMS_ENFORCE_AUTHORSHIP`                                   | `0` turns off the authorship checks for signed-in users (default on).                                                                                                                                                                                 |
| `TEAMS_MAX_MESSAGE_CHARS`, `TEAMS_MAX_REACTIONS_PER_MESSAGE` | Longest message text (default 32768) and most reactions on one message (default 200) the authorship check accepts.                                                                                                                                    |
| `TEAMS_ADMIN_TOKEN`, `TEAMS_AUTH_RATE_*`                     | Enables `POST /auth/revoke`; the `/auth/session` rate limit (see Hardening).                                                                                                                                                                          |
| `TEAMS_ORIGINS`                                              | Comma-separated allowed `Origin`s for sockets and files. Empty: any origin when a token is set; without a token only loopback origins and clients that send no `Origin`.                                                                              |
| `TEAMS_TRUST_PROXY`                                          | `1` (or a hop count, or `--trust-proxy`): take the client address for rate limits from `X-Forwarded-For`. Off by default.                                                                                                                             |
| `TEAMS_INSECURE`                                             | `1` (or `--insecure`) allows a non-loopback `HOST` with no token. Otherwise the server refuses to start.                                                                                                                                              |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP`      | Open sockets overall (default 500) and per remote address (20).                                                                                                                                                                                       |
| `TEAMS_MAX_ROOMS`                                            | Sync rooms and call rooms (default 1000 each).                                                                                                                                                                                                        |
| `TEAMS_MAX_DOC_MB`                                           | Largest persisted room document in MB (default 64); updates past it are rejected.                                                                                                                                                                     |
| `TEAMS_MAX_UPDATE_STRUCTS`                                   | Most structs one update may hold (default one per 4 bytes of the update), counted before the update is decoded; a socket sending more is closed.                                                                                                      |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`                  | Messages per second (default 100) and burst (200) for each socket; a socket over it is closed.                                                                                                                                                        |
| `TEAMS_WORKSPACE_QUOTA_MB`                                   | Files per workspace in MB (default 1024; a private chat counts for its workspace); uploads past it get 507.                                                                                                                                           |
| `TEAMS_USER_QUOTA_MB`                                        | Files one person uploads, in MB, across every workspace and chat (default 1024); uploads past it get 507.                                                                                                                                             |
| `TEAMS_FILES_QUOTA_MB`                                       | Ceiling on the whole files directory in MB (default 20480); uploads past it get 507.                                                                                                                                                                  |
| `TEAMS_ROOMS_QUOTA_MB`, `TEAMS_MAX_STORED_ROOMS`             | Ceiling on every stored room together in MB (default 10240) and on how many rooms are stored (10000); an update past them closes its socket with `1009 server storage full`.                                                                          |
| `TEAMS_MAX_AWARENESS_KB`, `TEAMS_MAX_ROOM_AWARENESS_KB`      | Largest presence state of one client (default 64) and all presence of one room (4096), in KB; a socket going past either is closed with `1009 presence too large`.                                                                                    |
| `TEAMS_AWARENESS_RATE_KB`                                    | Presence bytes per second each socket may send, in KB (default 256, burst four times that); a socket over it is closed.                                                                                                                               |
| `TEAMS_MAX_REFUSALS_PER_MIN`                                 | Costly refusals per client per minute (default 30), see Limits below; past it the client's sync and file requests get 429 for the rest of the minute.                                                                                                 |
| `TEAMS_COMPACT_INTERVAL_MIN`                                 | Minutes between room snapshot compactions (default 360, `0` turns the timer off). A room is also compacted when it loads.                                                                                                                             |
| `TEAMS_METRICS_TOKEN`                                        | Turns on `GET /metrics`, served only to `Authorization: Bearer <this token>`. Off without it.                                                                                                                                                         |
| `TEAMS_LOG_FORMAT`, `--log-format=json`                      | `text` (default) or `json`: one JSON object per line with `time`, `level` and `msg`.                                                                                                                                                                  |
| `TEAMS_DATA`                                                 | Data directory (default `./data`): room snapshots and files.                                                                                                                                                                                          |

The same limits are options of `createTeamsServer` (`maxConnections`, `maxConnectionsPerIp`,
`maxRooms`, `maxDocBytes`, `maxUpdateStructs`, `maxSyncBytes` (largest sync message: 32 MiB, or a 64th of
the heap when smaller, so a 512 MiB heap takes 8 MiB messages), `messageRate`, `messageBurst`, `workspaceQuotaBytes`,
`userQuotaBytes`, `filesQuotaBytes`, `roomsQuotaBytes`, `maxStoredRooms`, `maxAwarenessBytes`, `maxRoomAwarenessBytes`,
`awarenessByteRate`, `maxRefusals`, `metricsToken`, `compactIntervalMs`, `sweepIntervalMs`, `insecure`, `insecureDev`, `trustProxy`).

### Limits

- **Presence** (Yjs awareness) is capped per client state and per room, and by bytes per second per
  socket, because the server keeps it in memory and sends all of it to every joiner.
- **Storage** cannot be grown by choosing new names. Files count against the workspace, against the
  person who uploaded them (in every scope together) and against the whole directory. With
  accounts an upload goes only to a workspace that exists (a sync room that holds something) or to a
  private chat, else 404. Stored rooms have a ceiling in bytes and in number, and a room nobody wrote
  to is never written to disk.
- **Refusals that cost work** are counted per client (the signed-in user, or the address without
  accounts): a chat join or chat file request refused after the server had to read that room from
  disk, and an update refused by the checks. A client with too many in a minute gets 429 for sync
  sockets and chat file requests until the minute is over. Refusals answered from memory are not
  counted, so a person who is refused while waiting to be added to a group is not turned away.
- **Checking costs the update, not the room.** Each loaded room keeps one copy of its document
  that updates are tried on, so the server holds up to two copies of each loaded room in memory.

Or embed it:

```js
import { createTeamsServer } from 'openteams-server';

const server = createTeamsServer({ users: { alice: 'long-random-secret' } });
await server.listen(8787, '127.0.0.1');
```

Put it behind TLS (browsers refuse `ws://` from an `https://` page), and add a TURN server for calls
across NATs. The full deployment guide is
[docs/deploy.md](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/docs/deploy.md).

## Authorship and moderators

With accounts, every sync message from a socket is checked against the room document before it is
applied (`validateTeamsSyncMessage` of ooxml-core, bundled into `validator.generated.mjs`). A
signed-in user may only post messages and reactions as themselves, edit or delete their own
messages, create channels and tabs as themselves, rename, archive or delete only the channels and
tabs they created, and pin as themselves (in a private chat, only members pin). The chat directory
is checked too: only a chat's members change its row, its kind and creator never change, and a 1:1
row never changes at all. A violating message
is not applied and the socket is closed with 1008 (`update rejected: forged-author`), so the client
resyncs. A user with the `moderator` role may additionally hide or remove anybody's messages and
reactions and change anybody's channels, tabs and meetings, remove anybody from a private chat and
remove anybody's roster entry, but never write in another person's name. Give
the role in the users file:

```json
{ "ann": "scrypt$...", "root": { "credential": "scrypt$...", "role": "moderator" } }
```

or return `{ id, role: 'moderator' }` from `authenticate`. The role rides in the signed session
(`POST /auth/session` also answers it as `role`), so changing a role takes effect when the person
signs in again; revoke their sessions to force it. Disable with `enforceAuthorship: false` or
`TEAMS_ENFORCE_AUTHORSHIP=0`. Without accounts nothing is enforced, because the user is a claim.
Meetings follow the channel rule (only the organiser edits or cancels), a roster entry is writable
only under your own id, a chat member is added only by a member and removed only by themselves, and
message text and reactions per message are capped (`maxMessageChars`, `maxReactionsPerMessage`).
Yjs client ids and who may post where are not checked; the full list is in the core's
`docs/teams-area.md`. A close with `1008 update rejected` is shown to the person, and the client
backs off (5 s, 10 s, ...) and stops after three in a row instead of reconnecting in a loop.

On every socket, with or without accounts, an update is applied whole or not at all. It is tried on
a copy of the room first, so one that cannot be checked or applied changes nothing, reaches nobody
else, and closes its sender's socket with `1008 update rejected: invalid-update`. The copy is kept
with the room and follows it, so trying an update costs time in proportion to the update; a refused
update makes the server build the copy again, which is why refused updates are counted (see Limits).

## Accounts and sessions

The server needs a way to check a credential (see above); without one it only starts with
`--insecure-dev`, where user ids are claims and the server says so at startup. To bind ids to real
people:

```js
import { createTeamsServer } from 'openteams-server';

// A closed team: user id -> credential (or TEAMS_USERS_FILE=users.json, a JSON object)
createTeamsServer({
	users: { alice: 'long-random-secret' },
	sessionSecret: process.env.TEAMS_SESSION_SECRET,
});

// Or plug in your identity provider. Return { id }, a string id, or null to refuse.
createTeamsServer({
	authenticate: async ({ user, credential, headers }) =>
		(await myIdp.check(headers, credential)) ? { id: user } : null,
});
```

Clients call `POST /auth/session` with `{ user, credential }` (core: `requestTeamsSession`) and get
a signed session (HMAC-SHA-256, 12 hours by default, `sessionTtlMs`). Sockets then pass
it as the subprotocol `openteams.session.<session>`; the server takes the user id from it and ignores the claimed `?user=`. With accounts
configured, sockets without a valid session get 401, and an open socket is closed with
`1008 session expired` when its session expires. Set `TEAMS_SESSION_SECRET` (or
`sessionSecret`) to keep sessions across restarts and across processes; without it a random secret
is used per process. Anyone who knows the secret can sign a session for any user and role, so the
server refuses to start with one shorter than 32 characters or with fewer than 8 distinct
characters; generate it (`openssl rand -hex 32`). The server never logs credentials, sessions or the secret.

### Hardening

- **Hashed credentials.** A `users` value is either plain text or an scrypt hash
  (`scrypt$N$r$p$salt$hash`, node:crypto, random salt, constant-time check). Make one with the
  bundled command, which reads the password from a hidden prompt or from standard input (never from
  the command line):

  ```bash
  npx openteams-hash-password            # prompts, prints the hash
  printf '%s' "$PASSWORD" | npx openteams-hash-password
  ```

  Plain text is still accepted so existing files keep working, but the server logs a startup
  warning with the number (not the names or values) of plain text credentials. An unknown user id
  costs one hash check, so timing does not reveal which ids exist. Hashing protects the file at
  rest; the credential still travels to the server in the request body, so use TLS.

- **Rate limiting `/auth/session`.** A sliding window counts every attempt per remote address and
  per user id; over the limit the server answers `429` with `Retry-After` (seconds). Defaults: 60 s
  window, 30 attempts per address, 10 per user. Configure with `sessionRate: { windowMs, perIp,
  perUser }` or `TEAMS_AUTH_RATE_WINDOW_MS`, `TEAMS_AUTH_RATE_PER_IP`, `TEAMS_AUTH_RATE_PER_USER`.
  The address is the socket peer: behind a reverse proxy every client shares the proxy's address, so
  also rate limit at the proxy. The per-user limit stops guessing but also lets someone who hammers a
  known id keep that person from signing in for the rest of the window (a lockout, not a bypass).
  State is in memory and per process.

- **Checking a session.** `POST /auth/check` with `{ session }` answers `204` while the session is
  valid and `401` when it expired, was revoked or is forged; the Teams UI uses it to tell a rejected
  session (sign in again) from a server that is merely unreachable.

- **Revoking sessions.** Each user has a session generation that tokens carry. `POST /auth/signout`
  with `{ session }` (and the workspace token as `Authorization: Bearer`, if set) bumps it, which
  invalidates every session of that user, on all devices, and closes their open sockets. An
  operator can do the same with `server.revokeSessions(user)` or, when `TEAMS_ADMIN_TOKEN`
  (`adminToken`) is set, `POST /auth/revoke` with `{ user }` and `Authorization: Bearer <admin
  token>`; without an admin token that route is a 404. Generations are saved to
  `session-revocations.json` in the data directory, so a revocation survives a restart as long as
  `TEAMS_SESSION_SECRET` is fixed (with a random per-process secret, restarting already ends every
  session).

Remaining limits: the `users` table is a file you must protect, there is no password reset or
account management, a revocation applies to the whole user rather than one device, and sessions
are valid until they expire or are revoked. Plug an identity provider into `authenticate` for anything
larger than a small closed team.

**`--insecure-dev` is not secure.** With no `users` and no `authenticate`, `/auth/session` signs any
claimed id, so it adds nothing over open access. The server only runs that way when asked to.

## Per-room permissions

Hosts that keep their own permission model (a review queue, a CRM record, a project) decide per
sync socket with `authorizeSync`. It runs after the token, origin and session checks, may be
async, and returns `'write'`, `'read'`, `{ access, principal }` or a falsy value to refuse (403).
A read-only socket receives the document and shares presence, but the server drops its document
updates. The principal (by default the session user, or the claimed `user`) names the socket for
`closeSync`:

```js
import { createServer } from 'node:http';
import { createTeamsServer } from 'openteams-server';

const teams = createTeamsServer({
	// Verify your own short-lived ticket; never trust a query parameter on its own.
	authorizeSync: async ({ room, url }) => {
		const ticket = await myTickets.verify(url.searchParams.get('ticket'));
		if (!ticket || ticket.room !== room) return null;
		return { access: ticket.canEdit ? 'write' : 'read', principal: ticket.account };
	},
});

// Serve the sockets from your own origin instead of teams.listen():
const app = createServer(myHandler);
app.on('upgrade', (req, socket, head) => {
	if (req.url?.startsWith('/sync/')) teams.handleUpgrade(req, socket, head);
});

// After you change someone's access, drop their sockets so they reconnect and are asked again.
teams.closeSync('deck-42', 'ben'); // close code 4403 by default
```

`syncPrincipals(room)` lists who is connected and `isSyncRoomEmpty(room)` tells whether a room
has content yet (to decide which participant seeds it). The hook covers sync rooms only. A server
with `authorizeSync` and no `users` or `authenticate` cannot tell who a request is from, so it
offers no sessions, file storage or call signaling: `/auth/session`, `/files/...` and `/signal/...`
answer 404 (`insecureDev: true` opens them, for development). Add `authenticate` to offer them. A
verdict applies for the socket's lifetime: changing access does not affect open sockets until you
close them. `demos/pptx/demo-host-review` in the ooxml repository
uses all of this for collaborative PowerPoint review.

## Security model

- **One shared token opens every room and every channel file.** Per-room permissions exist only
  when you pass `authorizeSync` (above). Treat the token like a team password.
- **All members of a room trust each other** (the workspace room, and each private chat's room). The server only moves Yjs bytes, so any member can edit
  or delete other people's messages and forge author ids and roles.
- **Private chats have their own rooms.** A room named `chat-<16 hex>-<chat id>` admits a socket
  only when its user is `true` in the room's `teams:chat-members` map, and sockets of people removed
  from the map are closed. Nobody chooses the members of a new room: a 1:1 room (`dm-<32 hex>`) is
  created for the two people its id was derived from (the opener passes them in
  `?members=<id>,<id>`), and a group room (`g-<32 hex>-<24 hex>`) for the one person its id is bound
  to, who then adds the others through validated updates. A room that has held anything is never
  created again, even with an empty member list, and a name that differs from a chat room only by
  letter case is refused. Without accounts the `user` is a claim, so anyone holding the token can
  claim another person's id; with accounts (above) it comes from the signed session.
- **The token and session travel in the `Sec-WebSocket-Protocol` header**
  (`openteams, openteams.session.<session>, openteams.token.<token>`; the server echoes only
  `openteams`), which keeps them out of URLs and access logs. The shipped client sends them that
  way, and falls back to the query string (`?token=`, `?session=`) only for a credential that
  contains characters a subprotocol cannot carry; the server accepts both. Run the server behind
  TLS. Files never take the token in a URL: they use the `Authorization` header or a signed link.
- **Authorship is enforced for signed-in users** (see Authorship below).
- **Protocol version.** Clients announce `proto=<n>`. A client whose major version the server does
  not speak is accepted at the handshake and closed with 1008 and a reason such as `unsupported
  protocol version 1: this server speaks 2`. A missing `proto` is accepted (older clients) and
  logged.
- **Defaults are local.** Without a token the server refuses a non-loopback `HOST` (unless
  `TEAMS_INSECURE=1`) and accepts only loopback origins, so a website you visit cannot drive a server
  on your own machine. Set `TEAMS_TOKEN` and `TEAMS_ORIGINS` for any shared deployment.
- **Use short-lived TURN credentials**; static ones are visible to everyone who opens the app.
- **Limits** on connections, rooms, document size, message rate and file storage bound abuse, but do
  not replace a reverse proxy's own rate limiting.
- A socket can only speak for the awareness ids it announced itself. A call signaling socket speaks
  for one peer id: with accounts, the signed-in user's id (a signal under any other id closes the
  socket with `1008 peer id is not the signed-in user`, and a second socket of the same user in one
  call is closed with `1008 peer id in use`); without accounts, the first id it uses. An addressed
  call signal is delivered to that peer only.
- **A private chat's call is for its members.** Its signaling room is `call-` and the chat's room
  name (`call-chat-<16 hex>-<chat id>`). With accounts the server admits only the chat's members
  (403 otherwise); without them it closes a socket whose first peer id is not a member. Someone who
  leaves the chat is dropped from its call (`1008 not a member`). The form older clients used for a
  current chat id (`call-<workspace>-<chat id>`) and letter-case variants of a chat's call room are
  refused (403).
- An upload never overwrites a file: a name already taken is stored under a new one, returned in
  the response `url`.

## Files, compaction and metrics

- **File permissions.** A file lives under a scope: the workspace id (channel files, readable by every
  signed-in user of the server) or a private chat's room name `chat-<hash>-<chat id>` (only that chat's
  members may list, read, upload to, delete from or link it; the server reads the room's
  `teams:chat-members` map as it does for the sync room, and a room nobody created has no members).
  With accounts every file request needs the signed session in `X-Teams-Session`; without them the claimed
  `X-Teams-User` stands in. A signed download link is issued only after that check and opens its one
  file for five minutes.
- **Listing and deleting.** `GET /files/<scope>` answers `{ files: [{ name, size, modified, uploader }] }`;
  `DELETE /files/<scope>/<name>` is allowed to the uploader or a moderator. Uploader records live in
  `file-meta/=<name>.json` (files from older versions have none: only a moderator may delete them).
- **Member lists.** Beside a private chat room's snapshot the server keeps its member list
  (`rooms/=<name>.members.json`, with the snapshot's size), so it can refuse a join or a file request
  for a room that is not loaded without reading the room. A missing list, or one whose size does not
  match the snapshot, is rebuilt from the snapshot once.
- **Stored names.** Rooms, file scopes and uploader records are stored under `=` and the lower-case
  base32 of their exact name, so names that differ only by case stay apart on every disk. The first
  start of this version renames what an older one stored; an older server cannot read the result.
  An entry named after a chat id only this version issues (`dm-` and 32 hex digits, or a new `g-`
  id, in any spelling) is not migrated: it moves to a `.quarantine` folder beside it, is logged, and
  is never read.
- **Cleanup and quotas.** Uploads stream to a hidden `.part` file that failed uploads remove; start-up removes
  every leftover and a timer (`sweepIntervalMs`) those older than an hour. Quotas apply per workspace
  (`TEAMS_WORKSPACE_QUOTA_MB`), per uploader across every scope (`TEAMS_USER_QUOTA_MB`) and to the whole
  directory (`TEAMS_FILES_QUOTA_MB`); see Limits.
- **Compaction.** A room's snapshot is re-encoded as one Yjs update on a fresh document when the room loads and every
  `TEAMS_COMPACT_INTERVAL_MIN` minutes, only if it is smaller and describes the same state. Snapshots are
  written to a temporary file, flushed and renamed, so a crash keeps the old or the new one.
- **Metrics.** `GET /metrics` (with `TEAMS_METRICS_TOKEN`) exposes `teams_rooms_open`, `teams_connections{kind}`,
  `teams_stored_bytes{kind}`, `teams_updates_rejected_total{code}`, `teams_auth_failures_total{kind}`,
  `teams_rate_limited_total{scope}` and `teams_room_compactions_total`. Logs are plain text, or JSON lines with
  `--log-format=json`.

## Honest limits

This is not Microsoft Teams and is not affiliated with Microsoft. There is **no end-to-end
encryption**: this server stores and can read every chat message and file. The token is a shared
secret that keeps strangers out of a small team's server. Per-user authentication exists only when
you enable accounts; it is a small session layer, not a full identity system.

Licence: Apache-2.0.

## OneDrive host integration

This package does not acquire Microsoft credentials or synchronize OneDrive automatically.
An authorized application can use `ooxml-core/embedding/onedrive` to retrieve or persist Office
bytes and provide them to its normal document/tool workflow. See the
[package reference and conflict recovery guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md)
and [live acceptance requirements](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md).
OpenTeams server attachments remain independent copies; no simulation service is shipped.
