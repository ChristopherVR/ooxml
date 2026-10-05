# openteams-vue-viewer

Vue 3 binding for **OpenTeams**, an open-source, bring-your-own-server team workspace: channels
and chat, presence, meetings with video and screen share (WebRTC), and Office files shared in the
conversation. It gives you the complete UI as a component **and** a raw composable.

[Source](https://github.com/ChristopherVR/teams-viewer) | [Server](https://www.npmjs.com/package/openteams-server)

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
const config = { mode: 'server', syncUrl: 'wss://teams.example.com/sync', signalingUrl: 'wss://teams.example.com/signal' };
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
