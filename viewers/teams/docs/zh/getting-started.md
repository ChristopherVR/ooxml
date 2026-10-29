# 快速开始

OpenTeams 是一个开源、自带服务器的团队协作空间：频道与聊天、在线状态、基于 WebRTC 的视频和屏幕共享会议，以及在对话中共享的 Office 文件。你只需为所用框架安装一个绑定，将其指向一台服务器（或不使用服务器直接开始），即可获得完整的界面，也可以通过原始 hook 构建自己的界面。

::: warning 早期实现
这不是 Microsoft Teams，与 Microsoft 没有任何关联。**不提供端到端加密**：运行服务器的人可以读取每一条消息和每一个文件。在依赖它之前，请先阅读[局限性](/zh/limitations)。
:::

## 1. 安装绑定

每个绑定都是自包含的：`<teams-app>` 元素已打包在内，[`ooxml-core`](https://www.npmjs.com/package/ooxml-core)（逻辑）和 [`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui)（可视化基础组件）作为常规依赖一并安装。

::: code-group

```bash [React]
npm install openteams-react-viewer react
```

```bash [Vue]
npm install openteams-vue-viewer vue
```

```bash [Angular]
npm install openteams-angular-viewer @angular/core
```

```bash [Svelte]
npm install openteams-svelte-viewer svelte
```

```bash [Solid]
npm install openteams-solid-viewer solid-js
```

```bash [Vanilla]
npm install openteams-vanilla-viewer
```

:::

| 包                         | 适用于                  | 对等依赖        |
| -------------------------- | ----------------------- | --------------- |
| `openteams-react-viewer`   | React 18+               | `react`         |
| `openteams-vue-viewer`     | Vue 3.4+                | `vue`           |
| `openteams-angular-viewer` | Angular 17+（独立组件） | `@angular/core` |
| `openteams-svelte-viewer`  | Svelte 5                | `svelte`        |
| `openteams-solid-viewer`   | SolidJS 1.9             | `solid-js`      |
| `openteams-vanilla-viewer` | 无框架（`<teams-app>`） | 无              |
| `openteams-server`         | 参考服务器（Node 22+）  | 无              |

每个页面只使用一个绑定：它们都会注册相同的自定义元素。这些包以 `openteams-*` 命名，是因为 npm 上的 `openteams` 这个名称属于一个无关的项目。

## 2. 挂载

React：

```tsx
import { Teams } from 'openteams-react-viewer';

const config = {
	mode: 'server' as const,
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
};

export function Workspace() {
	return <Teams workspaceId="acme" userName="Ada" config={config} />;
}
```

原生 JavaScript：

```ts
import { mountTeams } from 'openteams-vanilla-viewer';

// `config` as in the React example above.
const teams = mountTeams(document.getElementById('teams')!, {
	workspaceId: 'acme',
	userName: 'Ada',
	config,
	onOpenFile: (detail) => console.log('open', detail.attachment.kind, detail.url),
});
teams.update({ workspaceId: 'acme', userName: 'Ada Lovelace', config });
teams.destroy();
```

或者在调用 `openteams-vanilla-viewer` 的 `defineTeamsApp()` 之后直接使用标记：

```html
<teams-app
	workspace-id="acme"
	user-name="Ada"
	server-config='{"mode":"server","syncUrl":"wss://teams.example.com/sync","signalingUrl":"wss://teams.example.com/signal","iceServers":[{"urls":"stun:stun.example.com:3478"}]}'
></teams-app>
```

所有绑定接受相同的属性：`workspaceId`（共享房间：字母、数字、`-` 和 `_`）、`userName`、`userId`（稳定的 id；默认使用此浏览器中记住的 id）、`config`、`uploadFile`（附件的去向）、`openers`（按 Office 类型区分的处理程序），以及三个事件：就绪、打开文件（对其调用 `preventDefault()` 可自行打开文件）和配置变更（React、Solid、Svelte 和原生 JavaScript 中为 `onReady`、`onOpenFile`、`onConfigChange`；Vue 中为 `@ready`、`@open-file`、`@config-change`；Angular 中为 `(ready)`、`(openFile)`、`(configChange)`）。各绑定的详情请参阅[框架指南](/frameworks/react)。

省略 `config` 时，该元素会提供一个设置对话框（保存在此浏览器中），或以本地模式运行。

## 3. 选择状态存放的位置

| 模式     | `config`                                                        | 行为                                                                                             |
| -------- | --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `local`  | `{ mode: 'local', iceServers: [...] }`                          | 无服务器。同一浏览器的标签页通过 `BroadcastChannel` 共享聊天、在线状态和通话。文件仅按名称共享。 |
| `server` | `{ mode: 'server', syncUrl, signalingUrl, iceServers, token? }` | 由你的服务器同步共享文档、中继通话信令，并（可选）存储文件。                                     |

[在线演示](/demos)使用的是本地模式。只要涉及多个浏览器，就需要运行一台服务器：

```bash
npx openteams-server --insecure-dev     # 仅用于本地开发：127.0.0.1:8787
```

然后使用 `ws://127.0.0.1:8787/sync` 和 `ws://127.0.0.1:8787/signal`（置于 TLS 之后则使用 `wss://`）。[服务器指南](/zh/server)介绍了协议、环境变量以及参考服务器不做的事情；[部署](/zh/deploy)介绍了 TLS、令牌、TURN 和 Docker。

## 4. 或构建你自己的界面

每个绑定还提供基于核心客户端的原始 hook：状态输入，纯动作输出。

```tsx
import { useTeams } from 'openteams-react-viewer';

const { client, state } = useTeams({
	workspaceId: 'acme',
	user: { id: 'u1', name: 'Ada' },
	config,
});
state?.channels.map((c) => <button onClick={() => client!.select(c.id)}>{c.name}</button>);
```

`state` 是一个不可变的 `TeamsState` 快照（连接状态、带未读计数的频道、所选频道的消息和文件、成员、输入状态、搜索结果、通话）。动作包括 `select`、`createChannel`、`send`、`startReply`、`startEdit`、`deleteMessage`、`toggleReaction`、`search`、`openCall`、`joinCall`、`leaveCall`、`toggleMic`、`toggleCamera`、`toggleScreenShare` 和 `toggleHand`。React 演示并排展示了这两种方式。

## 从源码运行

需要 Node 22+ 和 Bun。本文件夹是 [`ooxml`](https://github.com/ChristopherVR/ooxml) 仓库的 `viewers/teams`：先在仓库根目录运行 `bun install`，再在此处运行下面的命令。演示（根目录下的 `demos/teams/*`）和测试使用同一棵树中 `ooxml-core` 和 `ooxml-ui` 的源码。

```bash
bun install        # 在仓库根目录
bun run dev        # 在 viewers/teams 中：服务器 :8787，vanilla 演示 :5173，React 演示 :5174
```

打开 `http://127.0.0.1:5173/?name=Ada`，并在另一个标签页打开 `...?name=Bob`：两个人处于同一个工作空间。`?local=1` 表示不使用服务器运行。`TEAMS_USE_DIST=1` 会让演示使用已构建的包（需先执行 `bun run build:packages`），而不是源码。`TEAMS_DEMOS=vanilla,vue` 用于选择 `bun run dev` 启动的演示。摄像头和麦克风需要 `localhost` 或 HTTPS。
