# 原生 JavaScript

`openteams-vanilla-viewer` 不需要任何框架。它导出 `mountTeams`、`<teams-app>` 元素（`TeamsApp`，由 `defineTeamsApp()` 注册）以及原始的 `createTeams`。

```bash
npm install openteams-vanilla-viewer
```

## 挂载 {#mount-it}

```ts
import { mountTeams } from 'openteams-vanilla-viewer';

const props = {
	workspaceId: 'acme',
	userName: 'Ada',
	config: {
		mode: 'server' as const,
		syncUrl: 'wss://teams.example.com/sync',
		signalingUrl: 'wss://teams.example.com/signal',
		iceServers: [{ urls: 'stun:stun.example.com:3478' }],
	},
	onOpenFile: (detail) => console.log(detail.attachment.kind, detail.url),
};
const teams = mountTeams(document.getElementById('teams')!, props);

teams.update({ ...props, userName: 'Ada Lovelace' }); // only changed values are applied
teams.element.client; // the core client behind the element
teams.destroy();
```

## 纯标记 {#plain-markup}

```ts
import { defineTeamsApp } from 'openteams-vanilla-viewer';

defineTeamsApp();
const app = document.querySelector('teams-app')!;
app.config = { mode: 'local', iceServers: [] };
app.addEventListener('teams-open-file', (event) => console.log(event.detail.url));
```

特性：`workspace-id`、`user-name`、`user-id` 和 `server-config`（JSON）。属性：`config`、`uploadFile`、`openers`、`embeds`、`locale`。事件：`teams-ready`、`teams-open-file`（可取消）和 `teams-config-change`。

## 原始客户端 {#the-raw-client}

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState()));
teams.send({ text: 'hello' });
```

[原生 JavaScript 演示](/demo/){target="_self"}即 `demos/teams/vanilla`（位于仓库根目录）：一个由 `mountTeams` 挂载的 `<teams-app>`，只需几行设置代码。
