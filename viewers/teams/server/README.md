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
npx openteams-server                    # PORT=8787 HOST=127.0.0.1 by default
TEAMS_TOKEN=change-me TEAMS_ORIGINS=https://app.example.com TEAMS_DATA=/var/lib/openteams npx openteams-server
```

| Variable                                                | Meaning                                                                                                                                                                  |
| ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `PORT`, `HOST`                                          | Listen address. Bind `0.0.0.0` only behind TLS.                                                                                                                          |
| `TEAMS_TOKEN`                                           | If set, sockets take it as `?token=` and file requests as `Authorization: Bearer`.                                                                                       |
| `TEAMS_ORIGINS`                                         | Comma-separated allowed `Origin`s for sockets and files. Empty: any origin when a token is set; without a token only loopback origins and clients that send no `Origin`. |
| `TEAMS_INSECURE`                                        | `1` (or `--insecure`) allows a non-loopback `HOST` with no token. Otherwise the server refuses to start.                                                                 |
| `TEAMS_MAX_CONNECTIONS`, `TEAMS_MAX_CONNECTIONS_PER_IP` | Open sockets overall (default 500) and per remote address (20).                                                                                                          |
| `TEAMS_MAX_ROOMS`                                       | Sync rooms and call rooms (default 1000 each).                                                                                                                           |
| `TEAMS_MAX_DOC_MB`                                      | Largest persisted room document in MB (default 64); updates past it are rejected.                                                                                        |
| `TEAMS_MESSAGE_RATE`, `TEAMS_MESSAGE_BURST`             | Messages per second (default 100) and burst (200) for each socket; a socket over it is closed.                                                                           |
| `TEAMS_FILES_QUOTA_MB`                                  | Total size of the files directory in MB (default 1024); uploads past it get 507.                                                                                         |
| `TEAMS_DATA`                                            | Data directory (default `./data`): room snapshots and files.                                                                                                             |

The same limits are options of `createTeamsServer` (`maxConnections`, `maxConnectionsPerIp`,
`maxRooms`, `maxDocBytes`, `messageRate`, `messageBurst`, `filesQuotaBytes`, `insecure`).

Or embed it:

```js
import { createTeamsServer } from 'openteams-server';

const server = createTeamsServer({ token: process.env.TEAMS_TOKEN });
await server.listen(8787, '127.0.0.1');
```

Put it behind TLS (browsers refuse `ws://` from an `https://` page), and add a TURN server for calls
across NATs. The full deployment guide is
[docs/deploy.md](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/docs/deploy.md).

## Security model

- **One shared token opens every room and every file.** There are no per-user or per-room
  permissions. Treat the token like a team password.
- **All members of a room trust each other.** The server only moves Yjs bytes, so any member can edit
  or delete other people's messages and forge author ids and roles.
- **The token travels in the WebSocket query string** (`?token=`), which proxies and access logs can
  record and browsers can keep in history. Run the server behind TLS and do not log query strings.
  Files never take the token in a URL: they use the `Authorization` header or a signed link.
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
secret that keeps strangers out of a small team's server, not per-user authentication; put your own
auth in front for real accounts.

Licence: Apache-2.0.
