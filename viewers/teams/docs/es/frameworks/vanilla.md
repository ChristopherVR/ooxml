# JavaScript puro

`openteams-vanilla-viewer` no necesita ningún framework. Exporta `mountTeams`, el elemento
`<teams-app>` (`TeamsApp`, registrado por `defineTeamsApp()`) y `createTeams` en su versión sin
procesar.

```bash
npm install openteams-vanilla-viewer
```

## Móntalo {#mount-it}

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

## Marcado simple {#plain-markup}

```ts
import { defineTeamsApp } from 'openteams-vanilla-viewer';

defineTeamsApp();
const app = document.querySelector('teams-app')!;
app.config = { mode: 'local', iceServers: [] };
app.addEventListener('teams-open-file', (event) => console.log(event.detail.url));
```

Atributos: `workspace-id`, `user-name`, `user-id` y `server-config` (JSON). Propiedades:
`config`, `uploadFile`, `openers`, `embeds`, `locale`. Eventos: `teams-ready`, `teams-open-file`
(cancelable) y `teams-config-change`.

## El cliente sin procesar {#the-raw-client}

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState()));
teams.send({ text: 'hello' });
```

La [demo vanilla](/demo/){target="_self"} es `demos/teams/vanilla` en la raíz del repositorio: un
`<teams-app>` montado con `mountTeams` y unas pocas líneas de configuración.
