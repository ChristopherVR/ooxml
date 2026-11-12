# Svelte

`openteams-svelte-viewer`（对等依赖：`svelte` 5）以源码形式提供 `Teams.svelte`（默认导出），并附带一个已打包的运行时 `openteams-svelte-viewer/runtime`，其中包含原始存储和辅助函数。

```bash
npm install openteams-svelte-viewer svelte
```

## 组件 {#the-component}

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

它接受与所有绑定相同的属性（`workspaceId`、`userName`、`userId`、`config`、`uploadFile`、`openers`、`embeds`、`locale`、`onReady`、`onOpenFile`、`onConfigChange`）。

## 存储 {#the-store}

```svelte
<script>
	import { teamsStore } from 'openteams-svelte-viewer/runtime';

	const teams = teamsStore({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
</script>

{#each $teams?.channels ?? [] as channel}
	<button onclick={() => teams.client()?.select(channel.id)}># {channel.name}</button>
{/each}
```

只要存储有订阅者，客户端就会保持存活；其动作位于 `teams.client()` 上。
