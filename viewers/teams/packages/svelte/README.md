# openteams-svelte-viewer

Svelte 5 binding for **OpenTeams**, an open-source, bring-your-own-server team workspace: channels
and chat, presence, meetings with video and screen share (WebRTC), and Office files shared in the
conversation. It gives you the complete UI as a component **and** a raw store.

[Source](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams) | [Server](https://www.npmjs.com/package/openteams-server)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme) &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; **[OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme)** &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

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

## OneDrive integration

Use the optional shared `ooxml-core/embedding/onedrive` connection with host-supplied Microsoft
sign-in. See the [Office package and integration reference](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#sharing-and-openteams)
for this product's load/save API, native storage wiring, sharing permissions and conflict recovery.
The [testing guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)
separates test-local transports from live Microsoft acceptance. No simulation service is shipped;
OneDrive storage does not provide Microsoft Office coauthoring or automatic document merging.

## Privacy and network use

Chat, presence and session data go only to the `syncUrl` and `signalingUrl` you configure; local mode stays in the browser. Calls use Google's public STUN server (`stun:stun.l.google.com:19302`) unless you set `iceServers` to your own STUN/TURN servers (or `[]`), and website tabs load the URL a member added. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).
