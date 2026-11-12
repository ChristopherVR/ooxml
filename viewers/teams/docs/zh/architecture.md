# 架构

OpenTeams 的拆分方式与基于 [`ooxml`](https://github.com/ChristopherVR/ooxml) 构建的其他产品（Word、Excel、PowerPoint）相同：逻辑和可视化基础组件都位于该仓库中，包括 `<teams-app>` 元素本身。本文件夹（即 `ooxml` 仓库的 `viewers/teams`）保存框架绑定、参考服务器和文档；演示位于根目录的 `demos/teams/`，浏览器测试位于 `e2e/teams/`。

| 层级           | 位置                                                               | 负责内容                                                                        |
| -------------- | ------------------------------------------------------------------ | ------------------------------------------------------------------------------- |
| 逻辑           | `ooxml-core/teams`（`ooxml` 仓库中的 `src/core/teams`）            | 聊天 CRDT、在线状态、通话、信令、服务器配置以及 `createTeamsClient` 存储        |
| 协作           | `ooxml-core/collab`                                                | Yjs 会话和传输层（WebSocket、`BroadcastChannel`），与格式无关                   |
| 可视化基础组件 | `ooxml-ui/teams`（`src/ui/src/teams`）                             | 头像、应用侧栏、频道列表、聊天列表、编辑器、会议前界面、通话网格和控件          |
| 应用           | `ooxml-ui/teams`（`src/ui/src/teams/app`）                         | `<teams-app>`：在存储之上组合各基础组件、设置对话框、原始存储辅助函数           |
| 绑定层         | `ooxml-ui/teams`（`src/ui/src/teams/app/bind.ts`）                 | `bindTeams`、`pickTeamsProps` 和 `TEAMS_CONTRACT`：各绑定使用的属性和事件连接   |
| Web 组件       | 此处的 `packages/web-component`（私有的 `teams-viewer`，从不发布） | 重新导出 `ooxml-ui/teams`，使各绑定只需导入一个名称；内联进每个绑定             |
| 绑定           | 此处的 `packages/{react,vue,angular,svelte,solid,vanilla}`         | 生命周期适配器，加上各自的原始 hook，以 `openteams-<framework>-viewer` 形式发布 |
| 参考服务器     | 此处的 `server/`                                                   | Yjs 同步、信令中继、文件存储，以 `openteams-server` 形式发布                    |

该区域在核心中的说明见 [`docs/teams-area.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-area.md)。

## 客户端：一份状态，简单的动作 {#the-client-one-state-plain-actions}

界面所需的一切都来自由 `createTeamsClient` 创建的同一个客户端（各绑定将其封装为 `createTeams`，并额外提供一个带保护的 `localStorage`，用于离线快照）：

```ts
import { createTeams } from 'openteams-vanilla-viewer';

const teams = createTeams({ workspaceId: 'acme', user: { id: 'u1', name: 'Ada' }, config });
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
```

`getState()` 返回一个不可变的 `TeamsState` 快照。每个按钮都是客户端上的一个动作。用户应当看到的错误通过 `on('notice', ...)` 送达。界面从不直接接触 Yjs、`RTCPeerConnection` 或服务器：**状态输入，动作输出**。

## `<teams-app>` {#} {#} {#} {#} {#}

这是一个 Lit 元素，拥有一个模板和一个真实的 CSS 文件（`teams-app.ts`、`teams-app.css`，位于 `src/ui/src/teams/app`）。它决定屏幕上显示什么，并把用户的操作转发为客户端动作；颜色、间距和圆角来自 `ooxml-ui` 的令牌（`var(--office-...)`）。对端发来的任何内容都以文本形式渲染（即 `textContent` 语义），而不会作为标记解析，核心也会重新校验从共享文档中读回的所有内容。

## 绑定 {#bindings}

绑定是一个生命周期适配器：它创建 `<teams-app>`，通过 `bindTeams` 转发属性（包括 class；各框架分别使用 `class` 或 `className`），并重新发出事件，这些逻辑全部位于 `ooxml-ui/teams`（`src/ui/src/teams/app/bind.ts`，基于共享的 `ooxml-ui/bindings` 工具包）。除 Angular 之外，`onOpenFile` 在所有绑定中都是 `(detail, event)`；Angular 的单值输出会发出 `{ detail, event }`。每个绑定还通过 `createTeams` 暴露原始 hook：

| 绑定            | 组件                       | 原始 hook                                                |
| --------------- | -------------------------- | -------------------------------------------------------- |
| React           | `<Teams />`                | `useTeams(options)`（`useTeamsClient`、`useTeamsState`） |
| Vue 3           | `<Teams />`                | `useTeams(() => options)` 返回浅层 ref                   |
| Solid           | `Teams(props)`             | `createTeamsClient(() => options)` 返回信号              |
| Svelte 5        | `Teams.svelte`（默认导出） | 来自 `/runtime` 的 `teamsStore(options)`（可读存储）     |
| Angular         | `<teams-workspace>`        | `TeamsService`（信号）                                   |
| 原生 JavaScript | `mountTeams(el, props)`    | `createTeams(options)`                                   |

行为上的错误只需在 Web 组件或核心中修复一次。hook 或绑定连接上的错误则需要在同一次修改中，在每个绑定里都修复。

构建时，`scripts/build-packages.mjs` 会把私有 Web 组件（包括其 CSS）内联进每个框架的打包产物，并将其声明文件展平，因此发布的压缩包只导入其清单中声明的内容（`ooxml-core`、`ooxml-ui`、`lit` 以及框架的对等依赖）。

## 聊天、在线状态和通话 {#chat-presence-and-calls}

- **聊天。** 频道保存在一个 `Y.Map` 中，每个频道的消息各保存在一个 `Y.Map` 中，表情回应保存在以消息、表情和用户为键的扁平映射中，因此并发的表情回应不会相互覆盖。在客户端，只有作者可以编辑和删除消息（删除是软墓碑）。
- **在线状态**由 Yjs 感知（awareness）提供：在线的人、正在输入、状态以及状态消息。它只是参考信息，且未经身份验证。对端还会收到旧版的三值 `availability`。
- **置顶**消息每个频道使用一个 Yjs 映射；任何成员都可以置顶或取消置顶。
- **通话**是使用完美协商（perfect-negotiation）模式的全网状对等连接；静音、摄像头和屏幕共享通过切换轨道实现，无需重新协商。网状结构在大约十二人以内是可靠的（`maxParticipants`，默认 12）。
- **文件**是链接，而不是内联的字节：通过 `uploadFile` 或服务器的 `/files` 端点上传。Word、Excel、PowerPoint 和 Visio 文件通过文件名识别，因此宿主可以用对应的查看器打开它们（即 `openers` 属性或打开文件事件）。

## 传输层 {#transports}

在 `local` 模式下，客户端通过 `BroadcastChannel` 同步和传递信令，因此只有同一浏览器中的标签页才能看到彼此；GitHub Pages 上的演示即以此方式运行。在 `server` 模式下，同步使用 y-websocket 线格式，信令使用 JSON 中继；参见[自带服务器](/zh/server)。
