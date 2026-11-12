# Vue 3

`openteams-vue-viewer`（对等依赖：`vue` 3.4 或更高版本）以 `<Teams>` 的形式提供完整界面，并以组合式函数（composable）的形式提供原始客户端。

```bash
npm install openteams-vue-viewer vue
```

## 组件 {#the-component}

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

事件：`ready`、`open-file`（附带详情和原始事件，因此你可以调用 `preventDefault()`）、`config-change`。该组件通过 `element` 暴露此元素。

## 组合式函数 {#the-composable}

```ts
import { useTeams } from 'openteams-vue-viewer';

const { client, state } = useTeams(() => ({ workspaceId: 'acme', user, config }));
// state.value?.channels, client.value?.send({ text: 'hi' })
```

`options` 可以是一个值、一个 ref 或一个 getter；当它变化时客户端会被重新创建，并随其作用域一同销毁。`client` 和 `state` 是浅层 ref。
