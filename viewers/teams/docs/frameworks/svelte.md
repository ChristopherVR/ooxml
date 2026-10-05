# Svelte

`openteams-svelte-viewer` (peer: `svelte` 5) ships `Teams.svelte` as source (the default export)
next to a bundled runtime, `openteams-svelte-viewer/runtime`, that holds the raw store and the
helpers.

```bash
npm install openteams-svelte-viewer svelte
```

## The component

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

It takes the same props as every binding (`workspaceId`, `userName`, `userId`, `config`,
`uploadFile`, `openers`, `onReady`, `onOpenFile`, `onConfigChange`).

## The store

```svelte
<script>
	import { teamsStore } from 'openteams-svelte-viewer/runtime';

	const teams = teamsStore({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
</script>

{#each $teams?.channels ?? [] as channel}
	<button onclick={() => teams.client()?.select(channel.id)}># {channel.name}</button>
{/each}
```

The client lives as long as the store has subscribers; its actions are on `teams.client()`.
