# Vanilla JavaScript

`openteams-vanilla-viewer` kommt ohne Framework aus. Es exportiert `mountTeams`, das Element `<teams-app>`
(`TeamsApp`, registriert von `defineTeamsApp()`) und das rohe `createTeams`.

```bash
npm install openteams-vanilla-viewer
```

## Einbinden {#mount-it}

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

## Reines Markup {#plain-markup}

```ts
import { defineTeamsApp } from 'openteams-vanilla-viewer';

defineTeamsApp();
const app = document.querySelector('teams-app')!;
app.config = { mode: 'local', iceServers: [] };
app.addEventListener('teams-open-file', (event) => console.log(event.detail.url));
```

Attribute: `workspace-id`, `user-name`, `user-id` und `server-config` (JSON). Eigenschaften:
`config`, `uploadFile`, `openers`, `embeds`, `locale`. Ereignisse: `teams-ready`, `teams-open-file` (abbrechbar) und
`teams-config-change`.

## Der rohe Client {#the-raw-client}

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState()));
teams.send({ text: 'hello' });
```

Die [Vanilla-Demo](/demo/){target="_self"} ist `demos/teams/vanilla` im Root des Repositorys: ein `<teams-app>`, eingebunden
mit `mountTeams` und wenigen Zeilen Einrichtung.
