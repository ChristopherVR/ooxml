# Vue 3

`openteams-vue-viewer` (peer: `vue` 3.4 or later) gives you the full UI as `<Teams>` and the raw
client as a composable.

```bash
npm install openteams-vue-viewer vue
```

## The component

```vue
<script setup lang="ts">
import { Teams } from 'openteams-vue-viewer';

const config = {
	mode: 'server' as const,
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
};
</script>

<template>
	<Teams
		workspace-id="acme"
		user-name="Ada"
		:config="config"
		@open-file="(detail) => console.log(detail.url)"
	/>
</template>
```

Events: `ready`, `open-file` (with the detail and the original event, so you can call
`preventDefault()`), `config-change`. The component exposes the element as `element`.

## The composable

```ts
import { useTeams } from 'openteams-vue-viewer';

const { client, state } = useTeams(() => ({ workspaceId: 'acme', user, config }));
// state.value?.channels, client.value?.send({ text: 'hi' })
```

`options` may be a value, a ref or a getter; the client is recreated when it changes and destroyed
with the scope. `client` and `state` are shallow refs.
