# Svelte

`openteams-svelte-viewer` (Peer: `svelte` 5) liefert `Teams.svelte` als Quelltext (der Default-Export) zusammen mit
einer gebündelten Laufzeit, `openteams-svelte-viewer/runtime`, die den rohen Store und die Hilfsfunktionen enthält.

```bash
npm install openteams-svelte-viewer svelte
```

## Die Komponente {#the-component}

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

Sie nimmt dieselben Props wie jedes Binding entgegen (`workspaceId`, `userName`, `userId`, `config`,
`uploadFile`, `openers`, `embeds`, `locale`, `onReady`, `onOpenFile`, `onConfigChange`).

## Der Store {#the-store}

```svelte
<script>
	import { teamsStore } from 'openteams-svelte-viewer/runtime';

	const teams = teamsStore({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
</script>

{#each $teams?.channels ?? [] as channel}
	<button onclick={() => teams.client()?.select(channel.id)}># {channel.name}</button>
{/each}
```

Der Client lebt, solange der Store Abonnenten hat. Seine Aktionen stehen unter `teams.client()`.
