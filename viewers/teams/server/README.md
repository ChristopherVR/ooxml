# openteams-server

The reference bring-your-own server for **OpenTeams**, an open-source team workspace. One small
Node process does three jobs:

- `ws /sync/<room>`: a y-websocket compatible Yjs document and awareness server, persisted to disk;
- `ws /signal/<room>`: a JSON relay for WebRTC call signaling;
- `POST|GET /files/<workspace>/<name>`: plain file storage for attachments (25 MB cap), with
  short-lived signed download links when a token is set.

It understands neither chat nor Office formats: it moves bytes, so any server that speaks the same
three contracts works with the OpenTeams bindings just as well.

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme) &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; **[OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme)** &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

## Run

Needs Node 22 or later.

```bash
TEAMS_USERS_FILE=users.json TEAMS_SESSION_SECRET=change-me npx openteams-server   # PORT=8787 HOST=127.0.0.1 by default
npx openteams-server --insecure-dev     # local development only: no accounts, ids are claims
```

**The server is secure by default.** With no `TEAMS_USERS_FILE` (or `users`), no `authenticate`
hook and no `authorizeSync` hook it refuses to start, because every client could otherwise claim any
user id. `--insecure-dev` (option `insecureDev: true`) opts out, logs a loud warning and is meant
for your own machine. A shared `TEAMS_TOKEN` alone does not count as authentication.

Flags: `--insecure-dev`, `--insecure` (see `TEAMS_INSECURE`) and `--trust-proxy[=N]`. `SIGTERM` and
`SIGINT` save every room before the process exits.

| Variable                                                     | Meaning                                                                                                                                                                  |
| ------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `PORT`, `HOST`                                               | Listen address. Bind `0.0.0.0` only behind TLS.                                                                                                                          |
| `TEAMS_TOKEN`                                                | If set, sockets take it as `?token=` and file requests as `Authorization: Bearer`.                                                                                       |
| `TEAMS_USERS_FILE`, `TEAMS_SESSION_SECRET`                   | Path to a JSON object of user id to credential, or to `{ credential, role }` (enables accounts); the secret that signs sessions.                                         |
| `TEAMS_ENFORCE_AUTHORSHIP`                                   | `0` turns off the authorship checks for signed-in users (default on).                                                                                                    |
| `TEAMS_MAX_MESSAGE_CHARS`, `TEAMS_MAX_REACTIONS_PER_MESSAGE` | Longest message text (default 32768) and most reactions on one message (default 200) the authorship check accepts.                                                       |
| `TEAMS_ADMIN_TOKEN`, `TEAMS_AUTH_RATE_*`                     | Enables `POST /auth/revoke`; the `/auth/session` rate limit (see Hardening).                                                                                             |
| `TEAMS_ORIGINS`                                              | Comma-separated allowed `Origin`s for sockets and files. Empty: any origin when a token is set; without a token only loopback origins and clients that send no `Origin`. |
| `TEAMS_TRUST_PROXY`                                          | `1` (or a hop count, or `--trust-proxy`): take the client address for rate limits from `X-Forwarded-For`. Off by default.                                                |
| `TEAMS_INSECURE`                                             | `1` (or `--insecure`) allows a non-loopback `HOST` with no token. Otherwise the server refuses to start.                                                                 |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP`      | Open sockets overall (default 500) and per remote address (20).                                                                                                          |
| `TEAMS_MAX_ROOMS`                                            | Sync rooms and call rooms (default 1000 each).                                                                                                                           |
| `TEAMS_MAX_DOC_MB`                                           | Largest persisted room document in MB (default 64); updates past it are rejected.                                                                                        |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`                  | Messages per second (default 100) and burst (200) for each socket; a socket over it is closed.                                                                           |
| `TEAMS_FILES_QUOTA_MB`                                       | Total size of the files directory in MB (default 1024); uploads past it get 507.                                                                                         |
| `TEAMS_DATA`                                                 | Data directory (default `./data`): room snapshots and files.                                                                                                             |

The same limits are options of `createTeamsServer` (`maxConnections`, `maxConnectionsPerIp`,
`maxRooms`, `maxDocBytes`, `messageRate`, `messageBurst`, `filesQuotaBytes`, `insecure`,
`insecureDev`, `trustProxy`).

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
tabs they created, and pin as themselves (in a private chat, only members pin). A violating message
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
configured, sockets without a valid session get 401. Set `TEAMS_SESSION_SECRET` (or
`sessionSecret`) to keep sessions across restarts and across processes; without it a random secret
is used per process. The server never logs credentials, sessions or the secret.

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
has content yet (to decide which participant seeds it). The hook covers sync rooms only; signaling
rooms still use the token. A verdict applies for the socket's lifetime: changing access does not
affect open sockets until you close them. `demos/pptx/demo-host-review` in the ooxml repository
uses all of this for collaborative PowerPoint review.

## Security model

- **One shared token opens every room and every file.** Per-room permissions exist only
  when you pass `authorizeSync` (above). Treat the token like a team password.
- **All members of a room trust each other** (the workspace room, and each private chat's room). The server only moves Yjs bytes, so any member can edit
  or delete other people's messages and forge author ids and roles.
- **Private chats have their own rooms.** A room named `chat-<16 hex>-<chat id>` admits a socket
  only when its `?user=<id>` is listed in the room's `teams:chat-members` map; the first socket of a
  new room lists the members in `?members=<id>,<id>` (a 1:1 `dm-` room must list exactly the two
  people its id was derived from), and sockets of people removed from the map are closed. This keeps
  workspace members who are not on the list out of the room. Without accounts the `user` is a claim,
  so anyone holding the token can claim another person's id; with accounts (above) it comes from
  the signed session.
- **The token and session travel in the `Sec-WebSocket-Protocol` header**
  (`openteams, openteams.session.<session>, openteams.token.<token>`; the server echoes only
  `openteams`), which keeps them out of URLs and access logs. The shipped client sends them that
  way, and falls back to the query string (`?token=`, `?session=`) only for a credential that
  contains characters a subprotocol cannot carry; the server accepts both. Run the server behind
  TLS. Files never take the token in a URL: they use the `Authorization` header or a signed link.
- **Authorship is enforced for signed-in users** (see Authorship below).
- **Protocol version.** Clients announce `proto=<n>`. A client whose major version the server does
  not speak is accepted at the handshake and closed with 1008 and a reason such as `unsupported
  protocol version 2: this server speaks 1`. A missing `proto` is accepted (older clients) and
  logged.
- **Defaults are local.** Without a token the server refuses a non-loopback `HOST` (unless
  `TEAMS_INSECURE=1`) and accepts only loopback origins, so a website you visit cannot drive a server
  on your own machine. Set `TEAMS_TOKEN` and `TEAMS_ORIGINS` for any shared deployment.
- **Use short-lived TURN credentials**; static ones are visible to everyone who opens the app.
- **Limits** on connections, rooms, document size, message rate and file storage bound abuse, but do
  not replace a reverse proxy's own rate limiting.
- A socket can only speak for the awareness ids and the call peer id it announced itself, and an
  addressed call signal is delivered to that peer only. An upload never overwrites a file: a name
  already taken is stored under a new one, returned in the response `url`.

## Honest limits

This is not Microsoft Teams and is not affiliated with Microsoft. There is **no end-to-end
encryption**: this server stores and can read every chat message and file. The token is a shared
secret that keeps strangers out of a small team's server. Per-user authentication exists only when
you enable accounts; it is a small session layer, not a full identity system.

Licence: Apache-2.0.
