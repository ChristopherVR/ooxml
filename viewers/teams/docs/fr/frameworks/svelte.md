# Svelte

`openteams-svelte-viewer` (dépendance homologue : `svelte` 5) fournit `Teams.svelte` sous forme de
source (l’export par défaut), à côté d’un environnement d’exécution compilé,
`openteams-svelte-viewer/runtime`, qui contient le store brut et les helpers.

```bash
npm install openteams-svelte-viewer svelte
```

## Le composant {#the-component}

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

Il accepte les mêmes propriétés que toutes les liaisons (`workspaceId`, `userName`, `userId`,
`config`, `uploadFile`, `openers`, `embeds`, `locale`, `onReady`, `onOpenFile`, `onConfigChange`).

## Le store {#the-store}

```svelte
<script>
	import { teamsStore } from 'openteams-svelte-viewer/runtime';

	const teams = teamsStore({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
</script>

{#each $teams?.channels ?? [] as channel}
	<button onclick={() => teams.client()?.select(channel.id)}># {channel.name}</button>
{/each}
```

Le client vit tant que le store a des abonnés ; ses actions se trouvent sur `teams.client()`.
