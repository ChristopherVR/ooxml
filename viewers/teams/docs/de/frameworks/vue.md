# Vue 3

`openteams-vue-viewer` (Peer: `vue` ab Version 3.4) stellt die vollständige Oberfläche als `<Teams>` bereit und den
rohen Client als Composable.

```bash
npm install openteams-vue-viewer vue
```

## Die Komponente {#the-component}

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

Ereignisse: `ready`, `open-file` (mit dem Detail und dem ursprünglichen Ereignis, damit Sie `preventDefault()` aufrufen
können), `config-change`. Die Komponente stellt das Element als `element` bereit.

## Das Composable {#the-composable}

```ts
import { useTeams } from 'openteams-vue-viewer';

const { client, state } = useTeams(() => ({ workspaceId: 'acme', user, config }));
// state.value?.channels, client.value?.send({ text: 'hi' })
```

`options` kann ein Wert, ein Ref oder ein Getter sein. Der Client wird neu erzeugt, wenn sich dieser ändert, und mit
dem Scope zerstört. `client` und `state` sind flache Refs.
