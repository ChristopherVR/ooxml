# Vanilla JavaScript

`openteams-vanilla-viewer` needs no framework. It exports `mountTeams`, the `<teams-app>` element
(`TeamsApp`, registered by `defineTeamsApp()`) and the raw `createTeams`.

```bash
npm install openteams-vanilla-viewer
```

## Mount it

```ts
import { mountTeams } from 'openteams-vanilla-viewer';

const props = {
	workspaceId: 'acme',
	userName: 'Ada',
	config: {
		mode: 'server' as const,
		syncUrl: 'wss://teams.example.com/sync',
		signalingUrl: 'wss://teams.example.com/signal',
		iceServers: [{ urls: 'stun:stun.example.com:3478' }],
	},
	onOpenFile: (detail) => console.log(detail.attachment.kind, detail.url),
};
const teams = mountTeams(document.getElementById('teams')!, props);

teams.update({ ...props, userName: 'Ada Lovelace' }); // only changed values are applied
teams.element.client; // the core client behind the element
teams.destroy();
```

## Plain markup

```ts
import { defineTeamsApp } from 'openteams-vanilla-viewer';

defineTeamsApp();
const app = document.querySelector('teams-app')!;
app.config = { mode: 'local', iceServers: [] };
app.addEventListener('teams-open-file', (event) => console.log(event.detail.url));
```

Attributes: `workspace-id`, `user-name`, `user-id` and `server-config` (JSON). Properties:
`config`, `uploadFile`, `openers`. Events: `teams-ready`, `teams-open-file` (cancelable) and
`teams-config-change`.

## The raw client

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState()));
teams.send({ text: 'hello' });
```

The [vanilla demo](/demo/){target="_self"} is `demos/vanilla`: one `<teams-app>` mounted with
`mountTeams` and a few lines of setup.
