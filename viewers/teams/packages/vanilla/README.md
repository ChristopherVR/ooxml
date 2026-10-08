# openteams-vanilla-viewer

The framework-free binding for **OpenTeams**, an open-source, bring-your-own-server team workspace:
channels and chat, presence, meetings with video and screen share (WebRTC), and Office files shared
in the conversation. It contains the `<teams-app>` web component and `mountTeams()`.

[Source](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams) | [Server](https://www.npmjs.com/package/openteams-server)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme) &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; **[OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme)** &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

## Install

```bash
npm install openteams-vanilla-viewer
```

The `<teams-app>` element is bundled in; `ooxml-core`, `ooxml-ui` and `lit` are installed as
regular dependencies. Use a bundler (Vite, esbuild, webpack) that resolves bare imports.

## Use

```ts
import { mountTeams } from 'openteams-vanilla-viewer';

const teams = mountTeams(document.querySelector('#teams')!, {
	workspaceId: 'acme',
	userName: 'Ada',
	config: {
		mode: 'server',
		syncUrl: 'wss://teams.example.com/sync',
		signalingUrl: 'wss://teams.example.com/signal',
	},
});
teams.update({ workspaceId: 'acme', userName: 'Ada Lovelace' });
// teams.destroy();
```

For plain markup, call `defineTeamsApp()` once and write `<teams-app workspace-id="acme"
user-name="Ada"></teams-app>`. For your own UI, `createTeams({ workspaceId, user, config })` returns
the raw client (`getState()`, `subscribe()`, plain actions). Give the container a height. Without `config` the workspace runs in local mode (tabs of one
browser). For more than one machine, run
[`openteams-server`](https://www.npmjs.com/package/openteams-server) or any compatible server.

## Honest limits

This is not Microsoft Teams and is not affiliated with Microsoft. There is **no end-to-end
encryption**: your server can read chat. Mesh calls scale to about a dozen people, and
authorization is the server's job.

Licence: Apache-2.0.
