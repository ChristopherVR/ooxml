# openteams-svelte-viewer

Svelte 5 binding for **OpenTeams**, an open-source, bring-your-own-server team workspace: channels
and chat, presence, meetings with video and screen share (WebRTC), and Office files shared in the
conversation. It gives you the complete UI as a component **and** a raw store.

[Source](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams) | [Server](https://www.npmjs.com/package/openteams-server)

## Install

```bash
npm install openteams-svelte-viewer
```

`svelte` (5) is a peer dependency. The package ships `Teams.svelte` as source (your Svelte
toolchain compiles it) next to a bundled JavaScript runtime that includes the `<teams-app>`
element; `ooxml-core`, `ooxml-ui` and `lit` are installed as regular dependencies.

## Use

```svelte
<script lang="ts">
	import Teams from 'openteams-svelte-viewer';
	import { teamsStore } from 'openteams-svelte-viewer/runtime';

	// Raw store: subscribe for state, call actions on client().
	const teams = teamsStore({ workspaceId: 'acme', user: { name: 'Ada' } });
</script>

<div style="height: 100vh">
	<Teams workspaceId="acme" userName="Ada" />
</div>
<p>{$teams?.channels.length ?? 0} channels</p>
```

The root export is the component; the store and the helpers are under
`openteams-svelte-viewer/runtime`. Pass `config` (`{ mode: 'server', syncUrl, signalingUrl }`) to
use [`openteams-server`](https://www.npmjs.com/package/openteams-server) or any compatible server.

## Honest limits

This is not Microsoft Teams and is not affiliated with Microsoft. There is **no end-to-end
encryption**: your server can read chat. Mesh calls scale to about a dozen people, and
authorization is the server's job.

Licence: Apache-2.0.
