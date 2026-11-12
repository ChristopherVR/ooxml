# Vue 3

`openteams-vue-viewer` (dépendance homologue : `vue` 3.4 ou plus récent) vous fournit l’interface
complète sous la forme `<Teams>` et le client brut sous la forme d’un composable.

```bash
npm install openteams-vue-viewer vue
```

## Le composant {#the-component}

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

Événements : `ready`, `open-file` (avec le détail et l’événement d’origine, pour que vous puissiez
appeler `preventDefault()`), `config-change`. Le composant expose l’élément sous le nom `element`.

## Le composable {#the-composable}

```ts
import { useTeams } from 'openteams-vue-viewer';

const { client, state } = useTeams(() => ({ workspaceId: 'acme', user, config }));
// state.value?.channels, client.value?.send({ text: 'hi' })
```

`options` peut être une valeur, une ref ou un accesseur ; le client est recréé lorsqu’elle change,
et détruit avec la portée. `client` et `state` sont des refs superficielles.
