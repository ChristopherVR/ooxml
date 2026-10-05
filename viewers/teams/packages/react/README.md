# openteams-react-viewer

React binding for **OpenTeams**, an open-source, bring-your-own-server team workspace: channels
and chat, presence, meetings with video and screen share (WebRTC), and Office files shared in the
conversation. It gives you the complete UI as a component **and** raw hooks to build your own.

[Source](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams) | [Server](https://www.npmjs.com/package/openteams-server)

## Install

```bash
npm install openteams-react-viewer
```

`react` (18 or later) is a peer dependency. The `<teams-app>` element is bundled in; the workspace
logic comes from [`ooxml-core`](https://www.npmjs.com/package/ooxml-core) (`ooxml-core/teams`) and
the visual primitives from [`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui), both installed as
regular dependencies, together with `lit`.

## Use

The full UI (give it a height):

```tsx
import { Teams } from 'openteams-react-viewer';

<Teams
	workspaceId="acme"
	userName="Ada"
	config={{
		mode: 'server',
		syncUrl: 'wss://teams.example.com/sync',
		signalingUrl: 'wss://teams.example.com/signal',
	}}
	onOpenFile={(detail) => window.open(detail.url)}
/>;
```

Or your own markup over the raw client (state in, plain actions out):

```tsx
import { useTeams } from 'openteams-react-viewer';

const { client, state } = useTeams({
	workspaceId: 'acme',
	user: { id: 'ada', name: 'Ada' },
	config,
});
state?.channels.map((c) => <button onClick={() => client?.select(c.id)}>{c.name}</button>);
```

Without `config` the workspace runs in local mode (tabs of one browser share state over
`BroadcastChannel`). For more than one machine, run
[`openteams-server`](https://www.npmjs.com/package/openteams-server) or any server that speaks the
same three small contracts.

## Honest limits

This is not Microsoft Teams and is not affiliated with Microsoft. There is **no end-to-end
encryption**: your server can read chat. Mesh calls scale to about a dozen people, and
authorization is the server's job. Use one OpenTeams binding per page: each registers the same
custom elements.

Licence: Apache-2.0.
