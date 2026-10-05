# openteams-server

The reference bring-your-own server for **OpenTeams**, an open-source team workspace. One small
Node process does three jobs:

- `ws /sync/<room>`: a y-websocket compatible Yjs document and awareness server, persisted to disk;
- `ws /signal/<room>`: a JSON relay for WebRTC call signaling;
- `POST|GET /files/<workspace>/<name>`: plain file storage for attachments (25 MB cap), with
  short-lived signed download links when a token is set.

It understands neither chat nor Office formats: it moves bytes, so any server that speaks the same
three contracts works with the OpenTeams bindings just as well.

## Run

Needs Node 22 or later.

```bash
npx openteams-server                    # PORT=8787 HOST=127.0.0.1 by default
TEAMS_TOKEN=change-me TEAMS_ORIGINS=https://app.example.com TEAMS_DATA=/var/lib/openteams npx openteams-server
```

| Variable        | Meaning                                                                            |
| --------------- | ---------------------------------------------------------------------------------- |
| `PORT`, `HOST`  | Listen address. Bind `0.0.0.0` only behind TLS.                                    |
| `TEAMS_TOKEN`   | If set, sockets take it as `?token=` and file requests as `Authorization: Bearer`. |
| `TEAMS_ORIGINS` | Comma-separated allowed `Origin`s for sockets and files. Empty allows any.         |
| `TEAMS_DATA`    | Data directory (default `./data`): room snapshots and files.                       |

Or embed it:

```js
import { createTeamsServer } from 'openteams-server';

const server = createTeamsServer({ token: process.env.TEAMS_TOKEN });
await server.listen(8787, '127.0.0.1');
```

Put it behind TLS (browsers refuse `ws://` from an `https://` page), and add a TURN server for calls
across NATs. The full deployment guide is
[docs/deploy.md](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/docs/deploy.md).

## Honest limits

This is not Microsoft Teams and is not affiliated with Microsoft. There is **no end-to-end
encryption**: this server stores and can read every chat message and file. The token is a shared
secret that keeps strangers out of a small team's server, not per-user authentication; put your own
auth in front for real accounts.

Licence: Apache-2.0.
