# JavaScript natif

`openteams-vanilla-viewer` n’a besoin d’aucun framework. Il exporte `mountTeams`, l’élément
`<teams-app>` (`TeamsApp`, enregistré par `defineTeamsApp()`) et `createTeams` brut.

```bash
npm install openteams-vanilla-viewer
```

## Le monter {#mount-it}

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

## Balisage simple {#plain-markup}

```ts
import { defineTeamsApp } from 'openteams-vanilla-viewer';

defineTeamsApp();
const app = document.querySelector('teams-app')!;
app.config = { mode: 'local', iceServers: [] };
app.addEventListener('teams-open-file', (event) => console.log(event.detail.url));
```

Attributs : `workspace-id`, `user-name`, `user-id` et `server-config` (JSON). Propriétés :
`config`, `uploadFile`, `openers`, `embeds`, `locale`. Événements : `teams-ready`, `teams-open-file`
(annulable) et `teams-config-change`.

## Le client brut {#the-raw-client}

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState()));
teams.send({ text: 'hello' });
```

La [démo vanilla](/demo/){target="_self"} correspond à `demos/teams/vanilla` à la racine du dépôt :
un `<teams-app>` monté avec `mountTeams` et quelques lignes de configuration.
