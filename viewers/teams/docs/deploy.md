# Running your own server

## The reference server

```bash
cd server && node index.mjs          # PORT=8787 HOST=127.0.0.1 by default
```

| Variable                                                | Meaning                                                                                                                                                                                                             |
| ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `PORT`, `HOST`                                          | Listen address. Bind `0.0.0.0` only behind TLS.                                                                                                                                                                     |
| `TEAMS_TOKEN`                                           | If set, sockets take it as `?token=` and file requests as `Authorization: Bearer`.                                                                                                                                  |
| `TEAMS_ORIGINS`                                         | Comma-separated allowed `Origin`s for sockets and files. Empty: any origin when a token is set; without a token only loopback origins and clients that send no `Origin`.                                            |
| `TEAMS_DATA`                                            | Data directory (default `./data`): `rooms/<room>.bin` snapshots and `files/`.                                                                                                                                       |
| `TEAMS_USERS_FILE`, `TEAMS_SESSION_SECRET`              | Path to a JSON object of user id to credential, or `{ credential, role }` (enables accounts); the secret that signs sessions.                                                                                       |
| `TEAMS_ENFORCE_AUTHORSHIP`                              | `0` turns off the authorship checks for signed-in users (default on).                                                                                                                                               |
| `TEAMS_ADMIN_TOKEN`, `TEAMS_AUTH_RATE_*`                | Enables `POST /auth/revoke`; the `/auth/session` rate limit (window, per address, per user).                                                                                                                        |
| `TEAMS_TRUST_PROXY`, `--trust-proxy[=N]`                | Key the auth rate limiter and per-address connection cap on `X-Forwarded-For` (the Nth trusted proxy's entry, default 1). Off by default; enable only behind a proxy that sets the header, or clients can forge it. |
| `--insecure-dev`                                        | Start with no accounts and no `authenticate` hook (option `insecureDev`). Without it the server refuses to start unless `TEAMS_USERS_FILE` is set. Development only.                                                |
| `TEAMS_INSECURE`                                        | `1` (or `--insecure`) allows a non-loopback `HOST` with no token. Otherwise the server refuses to start.                                                                                                            |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP` | Open sockets overall (default 500) and per remote address (20).                                                                                                                                                     |
| `TEAMS_MAX_ROOMS`                                       | Sync rooms and call rooms (default 1000 each).                                                                                                                                                                      |
| `TEAMS_MAX_DOC_MB`                                      | Largest persisted room document in MB (default 64); updates past it are rejected.                                                                                                                                   |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`             | Messages per second (default 100) and burst (200) for each socket; a socket over it is closed.                                                                                                                      |
| `TEAMS_FILES_QUOTA_MB`                                  | Optional ceiling on the whole files directory in MB (off by default); uploads past it get 507.                                                                                                                      |
| `TEAMS_WORKSPACE_QUOTA_MB`                              | Files per workspace in MB (default 1024; a private chat counts for its workspace); uploads past it get 507.                                                                                                         |
| `TEAMS_COMPACT_INTERVAL_MIN`                            | Minutes between room snapshot compactions (default 360, `0` turns the timer off). A room is also compacted when it loads.                                                                                           |
| `TEAMS_METRICS_TOKEN`                                   | Turns on `GET /metrics` (Prometheus text), served only to `Authorization: Bearer <this token>`. Off without it.                                                                                                     |
| `TEAMS_LOG_FORMAT`, `--log-format=json`                 | `text` (default) or `json`: one JSON object per line with `time`, `level` and `msg`.                                                                                                                                |

It serves `ws /sync/<room>` (y-websocket compatible, one persisted `Y.Doc` per room),
`ws /signal/<room>` (a JSON relay) and `POST|GET /files/<workspace>/<name>` (25 MB cap, names
sanitised, downloads sent as attachments with `Referrer-Policy: no-referrer`). With a token, a
browser opens files through `POST /files/link/<workspace>/<name>` (token in the header), which
returns a link signed for that one file and valid for five minutes. `GET /health` reports room and call counts.

**Put it behind TLS.** Browsers refuse `ws://` from an `https://` page, and a token or call
signaling over plain HTTP is readable on the network. Any reverse proxy that upgrades WebSockets
works (nginx, Caddy, Traefik). Caddy example:

```
teams.example.com {
	reverse_proxy 127.0.0.1:8787
}
```

Then point the app at `wss://teams.example.com/sync` and `wss://teams.example.com/signal`.

**The token is a shared secret** stored in the browser: it keeps strangers out of a small team's
server, it is not per-user authentication. For real accounts, give the server `users` (or
`TEAMS_USERS_FILE`) or an `authenticate` function for your identity provider, and set
`TEAMS_SESSION_SECRET` so sessions survive restarts; see "Accounts and sessions" in the server's
README. Keep the same three contracts.

## Security model

- **One shared token opens every room and every file.** Per-user identity exists only when you
  enable accounts, and per-room permissions only when you pass `authorizeSync`. Treat the token
  like a team password.
- **All members of a room trust each other.** The server only moves Yjs bytes, so any member can edit
  or delete other people's messages and forge author ids and roles. Do not invite people you would
  not trust with the whole room.
- **The token and session travel in the `Sec-WebSocket-Protocol` header** (`openteams.session.<session>`,
  `openteams.token.<token>`, next to a bare `openteams`), which stays out of URLs and logs. The
  shipped client falls back to the query string only for a credential a subprotocol cannot carry.
  Run the server behind TLS. Files never take the token in a URL: they use the `Authorization`
  header or a signed link.
- **Authorship is enforced** for signed-in users: nobody can post as, edit or delete another
  person's messages, or change channels and tabs they did not create, unless the account has
  `"role": "moderator"` in the users file. `TEAMS_ENFORCE_AUTHORSHIP=0` turns it off.
- **Secure by default.** Without accounts, an `authenticate` hook or `authorizeSync` the server
  refuses to start; `--insecure-dev` is the explicit, warned opt-out. A shared token alone is not
  authentication.
- **Defaults are local.** Without a token the server refuses a non-loopback `HOST` (unless
  `TEAMS_INSECURE=1`) and accepts only loopback origins, so a website you visit cannot drive a server
  on your own machine. Set `TEAMS_TOKEN` and `TEAMS_ORIGINS` for any shared deployment.
- **Use short-lived TURN credentials** (see below); static ones are visible to everyone who opens the app.
- **Limits** on connections, rooms, document size, message rate and file storage (the variables
  above) bound abuse, but they do not replace a reverse proxy's own rate limiting.
- A socket can only speak for the awareness ids and the call peer id it announced itself, and an
  addressed call signal is delivered to that peer only. An upload never overwrites a file: a name
  already taken is stored under a new one, returned in the response `url`.

## Calls across NATs: TURN

Calls are peer to peer. Public STUN works on open networks; behind symmetric NAT or strict
firewalls you need a TURN relay, which also carries media only as ciphertext (DTLS-SRTP). A minimal
coturn (`turnserver.conf`):

```
listening-port=3478
tls-listening-port=5349
realm=turn.example.com
fingerprint
lt-cred-mech
user=teams:CHANGE_ME
no-multicast-peers
cert=/etc/letsencrypt/live/turn.example.com/fullchain.pem
pkey=/etc/letsencrypt/live/turn.example.com/privkey.pem
```

and in the app's server settings, one ICE server per line:

```
stun:turn.example.com:3478
turn:turn.example.com:3478,turns:turn.example.com:5349 teams CHANGE_ME
```

Static TURN credentials are visible to anyone who can open the app. For anything public, issue
short-lived credentials (coturn `use-auth-secret` plus a small endpoint that signs them) and pass
them in the `iceServers` you give the element.

## Docker

```dockerfile
FROM node:22-alpine
WORKDIR /app
COPY server/package.json ./
RUN npm install --omit=dev
COPY server/bin.mjs server/index.mjs server/auth.mjs server/passwords.mjs server/guards.mjs   server/rate-limit.mjs server/chat-access.mjs server/validator.generated.mjs server/files.mjs server/metrics.mjs server/room-store.mjs ./
ENV HOST=0.0.0.0 TEAMS_DATA=/data TEAMS_USERS_FILE=/config/users.json
RUN mkdir /data && chown node:node /data
VOLUME /data
EXPOSE 8787
USER node
CMD ["node", "bin.mjs"]
```

Build it from `viewers/teams` (`docker build -t openteams-server .` with the file saved there). It
copies every module the entry point imports (they are the `files` of `server/package.json`; copying
only `index.mjs` fails with a missing module) and installs the four runtime dependencies. The
server needs no other file: the data directory is created under `TEAMS_DATA`, and an accounts file,
if you use one, is mounted and named by `TEAMS_USERS_FILE`. Add `server/hash-password.mjs` if you
want `openteams-hash-password` inside the image. The recipe was checked by running those files with
`npm install --omit=dev` on Node, not by building the image.

Run it with an accounts file mounted and `TEAMS_SESSION_SECRET` and `TEAMS_ORIGINS` set, behind your TLS proxy:

```bash
docker run -d --name openteams -p 127.0.0.1:8787:8787 \
  -v openteams-data:/data -v "$PWD/users.json:/config/users.json:ro" \
  -e TEAMS_SESSION_SECRET=change-me -e TEAMS_ORIGINS=https://app.example.com \
  openteams-server
```

Without the mounted accounts file the container exits with the "refusing to start" message; do not
add `--insecure-dev` to a container that is reachable from a network. Add `--trust-proxy` to the
command (`CMD ["node", "bin.mjs", "--trust-proxy"]`) only if a reverse proxy sets `X-Forwarded-For`.
`docker stop` sends `SIGTERM`, which the server answers by saving every room first (Docker waits ten
seconds by default, which is ample). The image creates `/data` owned by `node`; with a bind
mount, `chown` the host directory to uid 1000.

## Operations

### Backup and restore

Everything the server keeps is under `TEAMS_DATA` (default `./data`):

| Path                       | Holds                                                                                   |
| -------------------------- | --------------------------------------------------------------------------------------- |
| `rooms/<room>.bin`         | One Yjs snapshot per room: all chat, channels, calendars and tab data.                  |
| `files/<workspace>/`       | Uploaded files.                                                                         |
| `file-meta/<scope>.json`   | Who uploaded each file (decides who may delete it). Back it up with `files/`.           |
| `session-revocations.json` | Session revocation generations (lose it and revoked sessions revive until they expire). |

Also keep what is not in the data directory: your accounts file (`TEAMS_USERS_FILE`) and
`TEAMS_SESSION_SECRET`. Changing or losing the secret signs everyone out, which is harmless.

To back up, copy the whole data directory. Room snapshots are written to a temporary file and
renamed, so a copy never sees a half-written snapshot; a copy taken while the server runs may miss
changes from the last half second (writes are debounced by 500 ms) and may capture a file that is
newer than its room's neighbours. For a consistent copy, stop the server first (`SIGTERM` flushes every
room before exit), copy, and start it again; or snapshot the volume at the filesystem level.

To restore, stop the server, replace the data directory with the backup (keeping `rooms/`, `files/`
and `session-revocations.json` together), restore the accounts file and secret, and start it. Clients
that were connected reconnect and merge: Yjs keeps whatever they hold that the backup lacks, so a
restore does not discard work still open in a browser.

### Files, compaction and metrics

- **File permissions.** A file lives under a scope: the workspace id for channel files (any signed-in
  user of the server) or a private chat's room name (`chat-<hash>-<chat id>`), which only that
  chat's members may list, read, upload to or link. The server checks the chat's member list the same
  way it does for the chat's sync room. With accounts, every file request must carry the signed session
  in `X-Teams-Session` (the client does this); a signed download link is issued only after that check and
  opens its one file for five minutes. With `--insecure-dev` the claimed `X-Teams-User` stands in, like
  every claim there. Channel files are not restricted per channel: all users of the server can read them.
- **Listing and deleting.** `GET /files/<scope>` answers `{ files: [{ name, size, modified, uploader }] }`.
  `DELETE /files/<scope>/<name>` is allowed to the uploader or a moderator (`204`, else `403`). Files
  uploaded before this version have no recorded uploader, so only a moderator can delete them.
- **Cleanup.** An upload streams to a hidden `.part` file that a failed or aborted upload removes. At start-up
  the server deletes every leftover `.part` file (and stale room temp files), and every 15 minutes
  any older than an hour. Files that were uploaded but never attached to a message stay until someone deletes them.
- **Quotas.** The limit is per workspace (`TEAMS_WORKSPACE_QUOTA_MB`, default 1024), counting the
  workspace's channel files and its private chats' files; deleting frees the space.
  `TEAMS_FILES_QUOTA_MB` is an optional ceiling on the whole directory.
- **Compaction.** When a room loads, and every `TEAMS_COMPACT_INTERVAL_MIN` minutes for every room, the server
  re-encodes the room as a single Yjs update on a fresh document and writes it only if it is smaller and
  describes exactly the same state. The write goes to a temporary file, is flushed to disk and renamed over
  the snapshot, so a crash leaves the old or the new snapshot, never a torn one; interrupted temp files are
  removed at start-up. Backups made before or after a compaction both restore.
- **Metrics.** Set `TEAMS_METRICS_TOKEN` and scrape `GET /metrics` with `Authorization: Bearer <token>`: `teams_rooms_open`,
  `teams_connections{kind}`, `teams_stored_bytes{kind="rooms|files"}`, `teams_updates_rejected_total{code}`,
  `teams_auth_failures_total{kind}`, `teams_rate_limited_total{scope}` and `teams_room_compactions_total`. Without
  the token the endpoint answers 404. Keep it off the public internet like any admin endpoint.
- **Logs.** `--log-format=json` (or `TEAMS_LOG_FORMAT=json`) writes one JSON object per line: `{ "time", "level", "msg" }`.

### Upgrading

1. Read the server's `CHANGELOG.md` for the version range you cross; breaking changes are marked.
2. Back up the data directory.
3. Stop the server (`SIGTERM`), install the new version (`npx openteams-server@<version>`, or rebuild the
   image), and start it. Room snapshots are plain Yjs updates and older ones are read as they are.
4. Check `GET /health`, then open a room.

Upgrading from a version before the secure default: a server that started without accounts now
refuses to. Add accounts (`TEAMS_USERS_FILE`), an `authenticate` hook or `authorizeSync`, or pass
`--insecure-dev` if (and only if) it is a local development server. A shared `TEAMS_TOKEN` on its own no
longer starts the server.

### Production hardening checklist

- [ ] Accounts configured (`TEAMS_USERS_FILE` with hashed credentials from `openteams-hash-password`, or an `authenticate` hook); `--insecure-dev` is not used.
- [ ] `TEAMS_SESSION_SECRET` set to a long random value kept outside the repository, so sessions survive restarts.
- [ ] TLS terminates in front of the server; the app uses `wss://`. The server itself listens on loopback or a private network.
- [ ] `TEAMS_ORIGINS` lists the exact origins of your app.
- [ ] `--trust-proxy` set only if the proxy sets `X-Forwarded-For` and clients cannot reach the server except through it; otherwise leave it off (all clients then share one rate-limit key behind the proxy, which is safe but coarse).
- [ ] Moderator accounts (`"role": "moderator"` in the users file) are the people you intend; remember a role is read at sign-in.
- [ ] `TEAMS_ADMIN_TOKEN` set if you need `POST /auth/revoke`, and kept secret.
- [ ] The data directory is on a persistent volume, writable only by the server's user, and backed up.
- [ ] Limits reviewed (`TEAMS_MAX_*`, `TEAMS_WORKSPACE_QUOTA_MB`) and the proxy applies its own rate limiting.
- [ ] Short-lived TURN credentials, not a static secret in the page.
- [ ] The process is supervised and stopped with `SIGTERM` so rooms are saved.
- [ ] You have read [Limitations](./limitations.md): no end-to-end encryption, room members trust each other.

## Using a different server

Anything that honours the contracts in the core's `docs/teams-area.md` works: a stock
`y-websocket` server for sync, any WebSocket fan-out for signaling, any object store for files (give
`<teams-app>` an `uploadFile` function, or set `uploadFile` on the element, and attachments go
there instead).
