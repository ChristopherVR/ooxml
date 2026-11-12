# Svelte

`openteams-svelte-viewer` (peer: `svelte` 5) incluye `Teams.svelte` como código fuente (la
exportación por defecto), junto a un runtime empaquetado, `openteams-svelte-viewer/runtime`, que
contiene el store sin procesar y las utilidades.

```bash
npm install openteams-svelte-viewer svelte
```

## El componente {#the-component}

```svelte
<script>
	import Teams from 'openteams-svelte-viewer';

	const config = {
		mode: 'server',
		syncUrl: 'wss://teams.example.com/sync',
		signalingUrl: 'wss://teams.example.com/signal',
		iceServers: [{ urls: 'stun:stun.example.com:3478' }],
	};
</script>

<Teams workspaceId="acme" userName="Ada" {config} onOpenFile={(detail) => console.log(detail.url)} />
```

Acepta las mismas props que todos los bindings (`workspaceId`, `userName`, `userId`, `config`,
`uploadFile`, `openers`, `embeds`, `locale`, `onReady`, `onOpenFile`, `onConfigChange`).

## El store {#the-store}

```svelte
<script>
	import { teamsStore } from 'openteams-svelte-viewer/runtime';

	const teams = teamsStore({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
</script>

{#each $teams?.channels ?? [] as channel}
	<button onclick={() => teams.client()?.select(channel.id)}># {channel.name}</button>
{/each}
```

El cliente vive mientras el store tenga suscriptores; sus acciones están en `teams.client()`.
