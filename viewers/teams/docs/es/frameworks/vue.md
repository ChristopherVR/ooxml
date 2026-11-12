# Vue 3

`openteams-vue-viewer` (peer: `vue` 3.4 o posterior) te ofrece la interfaz completa como `<Teams>` y
el cliente sin procesar como composable.

```bash
npm install openteams-vue-viewer vue
```

## El componente {#the-component}

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

Eventos: `ready`, `open-file` (con el detalle y el evento original, para que puedas llamar a
`preventDefault()`), `config-change`. El componente expone el elemento como `element`.

## El composable {#the-composable}

```ts
import { useTeams } from 'openteams-vue-viewer';

const { client, state } = useTeams(() => ({ workspaceId: 'acme', user, config }));
// state.value?.channels, client.value?.send({ text: 'hi' })
```

`options` puede ser un valor, una ref o un getter; el cliente se recrea cuando cambia y se destruye
con su ámbito. `client` y `state` son refs superficiales.
