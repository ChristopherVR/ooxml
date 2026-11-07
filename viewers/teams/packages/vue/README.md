# openteams-vue-viewer

Vue 3 binding for **OpenTeams**, an open-source, bring-your-own-server team workspace: channels
and chat, presence, meetings with video and screen share (WebRTC), and Office files shared in the
conversation. It gives you the complete UI as a component **and** a raw composable.

[Source](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams) | [Server](https://www.npmjs.com/package/openteams-server)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme) &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; **[OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme)** &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

## Install

```bash
npm install openteams-vue-viewer
```

`vue` (3.4 or later) is a peer dependency. The `<teams-app>` element is bundled in; `ooxml-core`,
`ooxml-ui` and `lit` are installed as regular dependencies.

## Use

```vue
<script setup lang="ts">
import { Teams, useTeams } from 'openteams-vue-viewer';
const config = {
	mode: 'server',
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
};
// Raw composable: shallow refs over the core client.
const { client, state } = useTeams(() => ({ workspaceId: 'acme', user: { name: 'Ada' }, config }));
</script>

<template>
	<Teams workspace-id="acme" user-name="Ada" :config="config" style="height: 100vh" />
</template>
```

Without `config` the workspace runs in local mode (tabs of one browser). For more than one machine,
run [`openteams-server`](https://www.npmjs.com/package/openteams-server) or any compatible server.

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
