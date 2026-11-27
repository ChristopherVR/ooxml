# openteams-server

**The reference server for OpenTeams: authenticated Yjs sync, a WebRTC signaling relay and scoped file storage in one small Node process that you run.**

[![npm](https://img.shields.io/npm/v/openteams-server.svg)](https://www.npmjs.com/package/openteams-server)
[![license](https://img.shields.io/npm/l/openteams-server.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE)
[![status: beta](https://img.shields.io/badge/status-beta-orange.svg)](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)

[Server guide](https://christophervr.github.io/ooxml/teams/server) · [Deploying](https://christophervr.github.io/ooxml/teams/deploy) · [Live demo](https://christophervr.github.io/ooxml/teams/demo/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/server/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

OpenTeams is an open-source, bring-your-own-server team workspace, in **Beta**. This package is the
server its framework bindings talk to. It understands neither chat nor Office formats: it moves
bytes, checks who sent them and stores them, so any server that speaks the same three contracts
(y-websocket sync, a JSON signaling relay, optional file storage) works just as well. It is not
Microsoft Teams, and there is no end-to-end encryption: the server stores and can read every
message and file.

| Route                                                                | Purpose                                                                                    |
| -------------------------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| `ws /sync/<room>`                                                    | y-websocket compatible Yjs document and presence, persisted to disk, every update checked. |
| `ws /signal/<room>`                                                  | JSON relay for WebRTC call signaling, bound to the signed-in user.                         |
| `POST\|GET\|DELETE /files/<scope>/<name>`, `GET /files/<scope>`      | Attachments (25 MB each) scoped by workspace and by private chat, with quotas.             |
| `POST /files/link/<scope>/<name>`                                    | A signed download link that opens one file for five minutes.                               |
| `POST /auth/session`, `/auth/check`, `/auth/signout`, `/auth/revoke` | Signed sessions: sign in, check, sign out everywhere, operator revocation.                 |
| `GET /health`                                                        | Liveness, room and call counts, and whether accounts are required.                         |
| `GET /metrics`                                                       | Prometheus metrics, only when `TEAMS_METRICS_TOKEN` is set.                                |

## Install and run

Needs Node.js 22.3 or newer. The package installs two commands: `openteams-server` and
`openteams-hash-password`.

On your own machine, with no accounts (user ids are claims, and the server says so loudly):

```bash
npx openteams-server --insecure-dev   # 127.0.0.1:8787
```

Point the app at `ws://127.0.0.1:8787/sync` and `ws://127.0.0.1:8787/signal`.

For a team, create accounts, a session secret and an origin allowlist, and put the server behind
TLS:

```bash
npx openteams-hash-password               # prompts for a password, prints an scrypt hash
openssl rand -hex 32 > session-secret     # once; keep it outside the repository
# users.json: { "ada": "scrypt$...", "root": { "credential": "scrypt$...", "role": "moderator" } }
TEAMS_USERS_FILE=users.json \
TEAMS_SESSION_SECRET="$(cat session-secret)" \
TEAMS_ORIGINS=https://app.example.com \
TEAMS_DATA=/var/lib/openteams \
npx openteams-server
```

Or embed it in your own Node process:

```js
import { createTeamsServer } from 'openteams-server';

const server = createTeamsServer({
	users: { ada: 'scrypt$...' },
	sessionSecret: process.env.TEAMS_SESSION_SECRET,
	origins: ['https://app.example.com'],
	dataDir: '/var/lib/openteams',
	log: console.log,
});
await server.listen(8787, '127.0.0.1');
// server.http is the node:http server; server.close() flushes every room to disk.
```

**Secure by default.** With no `TEAMS_USERS_FILE` (or `users`), no `authenticate` hook and no
`authorizeSync` hook the server refuses to start, because every client could claim any user id.
`--insecure-dev` (`insecureDev: true`) opts out with a loud warning, for your own machine only. A
shared `TEAMS_TOKEN` alone does not count as authentication. Separately, without `TEAMS_TOKEN` the
server refuses to listen on a non-loopback `HOST` unless `TEAMS_INSECURE=1` (`--insecure`), even
with accounts, so keep it on loopback behind your proxy or set a token.

Use the server and the OpenTeams bindings of the same major version: they speak the same Teams
protocol, and the server closes clients that speak another (`1008 unsupported protocol version`).
The first start of a new major version on an older data directory renames what it stored, and an
older server cannot read the result: back the directory up first.

## Configuration

Flags: `--insecure-dev`, `--insecure`, `--trust-proxy[=N]` and `--log-format=json`. `SIGTERM` and
`SIGINT` save every room before the process exits. Environment variables:

| Variable                                                                           | Default                       | Meaning                                                                                                                                                      |
| ---------------------------------------------------------------------------------- | ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `PORT`, `HOST`                                                                     | `8787`, `127.0.0.1`           | Listen address. Bind `0.0.0.0` only behind TLS.                                                                                                              |
| `TEAMS_DATA`                                                                       | `./data`                      | Data directory: room snapshots, files, uploader records, chat member lists, session revocations.                                                             |
| `TEAMS_USERS_FILE`                                                                 | none                          | JSON object of user id to credential, or to `{ credential, role }`. Enables accounts.                                                                        |
| `TEAMS_SESSION_SECRET`                                                             | random per process            | Signs sessions. At least 32 characters with 8 distinct ones, or the server refuses to start; set it so sessions survive restarts.                            |
| `TEAMS_TOKEN`                                                                      | none                          | Shared token: sockets take it in `Sec-WebSocket-Protocol` (`?token=` as a fallback), file requests as `Authorization: Bearer`.                               |
| `TEAMS_ORIGINS`                                                                    | none                          | Comma-separated allowed `Origin`s for sockets and files. Empty: any origin with a token; without one, only loopback origins and clients that send none.      |
| `TEAMS_INSECURE`                                                                   | off                           | `1` (or `--insecure`) allows a non-loopback `HOST` without a token.                                                                                          |
| `TEAMS_TRUST_PROXY`                                                                | off                           | `1` or a hop count (or `--trust-proxy`): take the client address from `X-Forwarded-For`. Only behind a proxy that sets it, or clients can forge it.          |
| `TEAMS_ENFORCE_AUTHORSHIP`                                                         | on                            | `0` turns off the authorship checks for signed-in users.                                                                                                     |
| `TEAMS_MAX_MESSAGE_CHARS`, `TEAMS_MAX_REACTIONS_PER_MESSAGE`                       | 32768, 200                    | Longest message text and most reactions on one message the authorship check accepts.                                                                         |
| `TEAMS_ADMIN_TOKEN`                                                                | none                          | Enables `POST /auth/revoke`.                                                                                                                                 |
| `TEAMS_AUTH_RATE_WINDOW_MS`, `TEAMS_AUTH_RATE_PER_IP`, `TEAMS_AUTH_RATE_PER_USER`  | 60000, 30, 10                 | `/auth/session` rate limit: window, attempts per address, attempts per user id. Over it: `429` with `Retry-After`.                                           |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP`                            | 500, 20                       | Open sockets overall and per remote address.                                                                                                                 |
| `TEAMS_MAX_ROOMS`                                                                  | 1000                          | Loaded sync rooms, and call rooms, each.                                                                                                                     |
| `TEAMS_MAX_DOC_MB`                                                                 | 64                            | Largest persisted room document; updates past it are rejected.                                                                                               |
| `TEAMS_MAX_UPDATE_STRUCTS`                                                         | one per 4 bytes of the update | Most structs one update may hold, counted before it is decoded; a socket sending more is closed.                                                             |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`                                        | 100, 200                      | Messages per second and burst per socket; a socket over it is closed.                                                                                        |
| `TEAMS_WORKSPACE_QUOTA_MB`, `TEAMS_USER_QUOTA_MB`, `TEAMS_FILES_QUOTA_MB`          | 1024, 1024, 20480             | Files per workspace (its private chats included), per uploader across every scope, and for the whole directory. Uploads past them get `507`.                 |
| `TEAMS_ROOMS_QUOTA_MB`, `TEAMS_MAX_STORED_ROOMS`                                   | 10240, 10000                  | All stored rooms together, and how many are stored; an update past them closes its socket with `1009 server storage full`.                                   |
| `TEAMS_MAX_AWARENESS_KB`, `TEAMS_MAX_ROOM_AWARENESS_KB`, `TEAMS_AWARENESS_RATE_KB` | 64, 4096, 256                 | Presence of one client, of one room, and presence KB per second per socket (burst four times that); past them: `1009 presence too large` or a closed socket. |
| `TEAMS_MAX_REFUSALS_PER_MIN`                                                       | 30                            | Costly refusals per client per minute (see [Limits](#limits)); past it, sync and chat file requests get `429` for the rest of the minute.                    |
| `TEAMS_COMPACT_INTERVAL_MIN`                                                       | 360                           | Minutes between room snapshot compactions (`0` turns the timer off); rooms are also compacted when they load.                                                |
| `TEAMS_METRICS_TOKEN`                                                              | none                          | Turns on `GET /metrics`, served only to `Authorization: Bearer <token>`.                                                                                     |
| `TEAMS_LOG_FORMAT`                                                                 | `text`                        | `json` (or `--log-format=json`): one `{ time, level, msg }` object per line.                                                                                 |

`createTeamsServer(options)` takes the same settings as options: `token`, `origins`, `dataDir`,
`users`, `sessionSecret`, `sessionTtlMs` (default 12 hours), `adminToken`,
`sessionRate: { windowMs, perIp, perUser }`, `enforceAuthorship`, `maxMessageChars`,
`maxReactionsPerMessage`, `maxConnections`, `maxConnectionsPerIp`, `maxRooms`, `maxDocBytes`,
`maxUpdateStructs`, `messageRate`, `messageBurst`, `workspaceQuotaBytes`, `userQuotaBytes`,
`filesQuotaBytes`, `roomsQuotaBytes`, `maxStoredRooms`, `maxAwarenessBytes`,
`maxRoomAwarenessBytes`, `awarenessByteRate`, `maxRefusals`, `metricsToken`, `compactIntervalMs`,
`sweepIntervalMs`, `insecure`, `insecureDev`, `trustProxy` and `log`. Options only: `authenticate`,
`authorizeSync`, `maxFileBytes` (25 MiB), `maxSignalBytes` (256 KiB) and `maxSyncBytes` (32 MiB,
or a 64th of the heap when smaller: size the heap with `NODE_OPTIONS=--max-old-space-size=...` or
set it for large documents).

### Accounts and sessions

Give the server `users` (or `TEAMS_USERS_FILE`), or plug in your identity provider:

```js
createTeamsServer({
	// Return { id }, { id, role: 'moderator' }, a string id, or null to refuse.
	authenticate: async ({ user, credential, headers }) =>
		(await myIdp.check(headers, credential)) ? { id: user } : null,
});
```

- Clients call `POST /auth/session` with `{ user, credential }` and get a signed session
  (HMAC-SHA-256, `sessionTtlMs`). Sockets pass it as the subprotocol `openteams.session.<session>`
  and file requests as `X-Teams-Session`; the server takes the user id from it. Sockets without a
  valid session get 401, and an open socket closes with `1008 session expired` when it expires.
- Credentials are plain text or scrypt hashes (`scrypt$N$r$p$salt$hash`) from
  `openteams-hash-password`, which reads the password from a hidden prompt or standard input
  (`printf '%s' "$PASSWORD" | npx openteams-hash-password`). Plain text works but logs a warning
  with their count. Unknown ids cost one hash check, so timing does not reveal which exist.
- `POST /auth/check` with `{ session }` answers 204 or 401. `POST /auth/signout` with `{ session }`
  revokes every session of that user on all devices and closes their sockets; operators do the
  same with `server.revokeSessions(user)` or `POST /auth/revoke` `{ user }` with
  `Authorization: Bearer <TEAMS_ADMIN_TOKEN>` (404 without an admin token). Revocations are saved to
  `session-revocations.json` and survive a restart with a fixed `TEAMS_SESSION_SECRET`.
- The `/auth/session` rate limit is in memory, per process and per socket peer: behind a proxy,
  rate limit there too (or enable `TEAMS_TRUST_PROXY`). The per-user limit can lock a known id out
  for the rest of the window.

There is no password reset or account management, and a revocation covers the whole user. For
anything larger than a small closed team, use `authenticate`. Under `--insecure-dev`,
`/auth/session` signs any claimed id.

### Per-room permissions

Hosts with their own permission model decide per sync socket with `authorizeSync`. It runs after
the token, origin and session checks, may be async, and returns `'write'`, `'read'`,
`{ access, principal }` or a falsy value (403). A read-only socket receives the document and shares
presence, but its updates are dropped.

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

// After changing someone's access, drop their sockets so they reconnect and are asked again.
teams.closeSync('deck-42', 'ben'); // close code 4403 by default
```

`syncPrincipals(room)` lists who is connected and `isSyncRoomEmpty(room)` tells whether a room has
content yet. A verdict lasts for the socket's lifetime. The hook covers sync only: a server with
`authorizeSync` and no `users` or `authenticate` cannot tell who a request is from, so
`/auth/session`, `/files/...` and `/signal/...` answer 404 (`insecureDev: true` opens them for
development). `demos/pptx/demo-host-review` in the ooxml repository uses all of this.

### Limits

- **Presence** is capped per client, per room and in bytes per second per socket, because the
  server keeps it in memory and sends all of it to every joiner.
- **Storage** cannot be grown by choosing new names: files count against the workspace, the
  uploader and the whole directory; with accounts an upload goes only to a workspace that exists or
  to a private chat (else 404); stored rooms have byte and count ceilings, and a room nobody wrote
  to is never stored.
- **Refusals that cost work** (a chat join or chat file request refused after reading the room from
  disk, an update refused by the checks) are counted per signed-in user, or per address without
  accounts. Refusals answered from memory are not counted.
- **Checking costs the update, not the room**: each loaded room keeps one checking copy of its
  document, so memory holds up to two copies of each loaded room. Updates are applied whole or
  not at all; one that cannot be checked or applied closes its sender with
  `1008 update rejected: invalid-update`.

## Security model

### Authorship and moderators

With accounts, every sync message is checked against the room before it is applied (the
`validateTeamsSyncMessage` check from `ooxml-core/teams`, bundled in `validator.generated.mjs`). A
signed-in user may only post messages and reactions as themselves, edit or delete their own
messages, create channels and tabs as themselves, change only the channels, tabs and meetings they
created, write only their own roster entry, and pin as themselves (in a private chat, members
only). Only a chat's members change its directory row; its kind and creator never change, and a 1:1
row never changes. A member is added only by a member and removed only by themselves. A violating
message is not applied and the socket closes with `1008 update rejected: forged-author`; the client
shows it and backs off (5 s, 10 s, ...), stopping after three in a row.

A `moderator` (`"role": "moderator"` in the users file, or from `authenticate`) may also hide or
remove anybody's messages and reactions, change anybody's channels, tabs and meetings, remove
anybody from a private chat or the roster, but never write in another person's name. The role
rides in the signed session, so a change takes effect at the next sign-in; revoke sessions to force
it. `enforceAuthorship: false` or `TEAMS_ENFORCE_AUTHORSHIP=0` disables the checks. Yjs client ids
and who may post where are not checked; the full list is in
[teams-area.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-area.md).

### What the server guarantees, and what it does not

- **Without accounts, all members of a room trust each other**: user ids are claims, so anyone
  holding the token can edit others' messages and forge authors and roles. One shared token opens
  every room and channel file; per-room permissions exist only with `authorizeSync`.
- **Private chats have their own rooms** (`chat-<16 hex>-<chat id>`), admitting only users marked
  in the room's `teams:chat-members` map and closing sockets of people removed from it. A 1:1 room
  (`dm-<32 hex>`) is created for the two people its id was derived from, a group room
  (`g-<32 hex>-<24 hex>`) for the one person its id is bound to; a room that held anything is never
  created again, and letter-case variants are refused.
- **Calls are bound to identities.** A signaling socket speaks for one peer id: with accounts the
  signed-in user's (`1008 peer id is not the signed-in user`, `1008 peer id in use`), without them
  the first id it uses. A private chat's call (`call-chat-<16 hex>-<chat id>`) admits only its
  members (403) and drops someone who leaves (`1008 not a member`). Addressed signals reach only
  their peer.
- **Credentials stay out of URLs**: sockets carry them in `Sec-WebSocket-Protocol`
  (`openteams, openteams.session.<session>, openteams.token.<token>`; the server echoes only
  `openteams`), with `?token=` / `?session=` only as a fallback. Files use the `Authorization` or
  `X-Teams-Session` header, or a signed link. With accounts, a file in a private chat's scope is
  reachable only by its members; channel files are readable by every signed-in user.
- **Files never overwrite**: a taken name is stored under a new one, returned in `url`. Deleting is
  for the uploader or a moderator.
- **Defaults are local**: without a token only loopback hosts and origins are accepted, so a website
  you visit cannot drive a server on your machine.
- **Not covered**: no end-to-end encryption, no TLS of its own, one process with rooms in memory
  (no horizontal scaling), no virus scanning or file versioning. Use short-lived TURN credentials,
  and keep a reverse proxy's own rate limiting: the limits above bound abuse, they do not replace
  it. The server never logs credentials, sessions or the session secret.

## Deploying

Run it behind a reverse proxy that terminates TLS and upgrades WebSockets (browsers refuse `ws://`
from an `https://` page), and add a TURN server for calls across strict networks. The
[deployment guide](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/docs/deploy.md)
covers TLS with a Caddy example, TURN, a Docker recipe (not yet built and tested), backup and
restore, compaction, metrics, upgrading and a production checklist;
[bring your own server](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/docs/server.md)
documents every endpoint and the three contracts.

## Privacy and network use

The server makes no outbound requests: it only answers the clients that connect to it. It stores
room snapshots, uploaded files, uploader records, chat member lists and session revocations in
`TEAMS_DATA`, and keeps presence and rate-limit state in memory. It can read every message and
file it stores. Call audio and video do not pass through it; call signaling, which carries each
participant's network addresses, does. See
[PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#teams-ooxml-coreteams-ooxml-uiteams-the-openteams--packages).

## Documentation

- [Server guide](https://christophervr.github.io/ooxml/teams/server),
  [deploying](https://christophervr.github.io/ooxml/teams/deploy) and
  [limitations](https://christophervr.github.io/ooxml/teams/limitations).
- Clients: [`openteams-react-viewer`](https://www.npmjs.com/package/openteams-react-viewer),
  [`openteams-vue-viewer`](https://www.npmjs.com/package/openteams-vue-viewer),
  [`openteams-angular-viewer`](https://www.npmjs.com/package/openteams-angular-viewer),
  [`openteams-svelte-viewer`](https://www.npmjs.com/package/openteams-svelte-viewer),
  [`openteams-solid-viewer`](https://www.npmjs.com/package/openteams-solid-viewer) and
  [`openteams-vanilla-viewer`](https://www.npmjs.com/package/openteams-vanilla-viewer).
- Build your own server on [`ooxml-core/teams`](https://www.npmjs.com/package/ooxml-core)
  ([teams-area.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-area.md)).
- Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE); see
[NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/NOTICE).
