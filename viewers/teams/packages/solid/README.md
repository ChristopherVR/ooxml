# openteams-solid-viewer

SolidJS binding for **OpenTeams**, an open-source, bring-your-own-server team workspace: channels
and chat, presence, meetings with video and screen share (WebRTC), and Office files shared in the
conversation. It gives you the complete UI **and** a raw reactive primitive.

[Source](https://github.com/ChristopherVR/teams-viewer) | [Server](https://www.npmjs.com/package/openteams-server)

## Install

```bash
npm install openteams-solid-viewer
```

`solid-js` (1.9) is a peer dependency. The `<teams-app>` element is bundled in; `ooxml-core`,
`ooxml-ui` and `lit` are installed as regular dependencies.

## Use

```tsx
import { Teams, createTeamsClient } from 'openteams-solid-viewer';

// The full UI: Teams() returns the <teams-app> element (no JSX transform needed).
<div style={{ height: '100vh' }}>{Teams({ workspaceId: 'acme', userName: 'Ada' })}</div>;

// Or signals over the raw client.
const { client, state } = createTeamsClient(() => ({ workspaceId: 'acme', user: { name: 'Ada' } }));
```

Pass `config` (`{ mode: 'server', syncUrl, signalingUrl }`) to use
[`openteams-server`](https://www.npmjs.com/package/openteams-server) or any compatible server;
without it the workspace runs in local mode (tabs of one browser).

## Honest limits

This is not Microsoft Teams and is not affiliated with Microsoft. There is **no end-to-end
encryption**: your server can read chat. Mesh calls scale to about a dozen people, and
authorization is the server's job.

Licence: Apache-2.0.
