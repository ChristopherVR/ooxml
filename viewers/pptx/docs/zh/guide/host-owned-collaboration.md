---
title: 宿主自有协作
description: 提供现有的 Yjs 文档、awareness 和同步状态，而无需让查看器拥有传输层。
---

# 宿主自有协作 {#host-owned-collaboration}

当你的应用已经拥有身份验证、Yjs provider、离线存储或共享连接时，请使用 `collaboration.externalSession`。React、Vue、Angular、Svelte 和 Vanilla 都接受相同的约定。查看器负责其幻灯片 schema 和在线状态（presence）；宿主负责连接、重连、授权和资源生命周期。

```ts
import type { CollaborationConfig, ExternalCollaborationSession } from 'pptx-react-viewer';

// 这些资源和 store 属于你的应用，而不是某次查看器挂载。
const session: ExternalCollaborationSession = {
	doc: hostDocument,
	awareness: hostAwareness,
	getSnapshot: () => connectionStore.getSnapshot(),
	subscribe: (listener) => connectionStore.subscribe(listener),
};

const collaboration: CollaborationConfig = {
	roomId: 'presentation-123',
	serverUrl: '', // 共享配置要求提供，但外部会话不会使用。
	userName: 'Alice',
	sessionIntent: 'join',
	externalSession: session,
};
```

请从你所用框架对应的包中导入这些类型，并将 `collaboration` 传给其常规的查看器属性或选项。文档必须是真正的 `Y.Doc`；awareness 必须属于该文档。请保持 session 对象稳定，并在 `status` 或 `synced` 变化时通知订阅者。订阅会返回一个取消订阅函数。

宿主与查看器必须解析到同一个 Yjs 运行时。查看器不会把 Yjs 打包进自身，而是将其声明为可选的 peer 依赖。无论使用内置协作还是宿主自有协作，都请在应用中安装 `yjs`：

```sh
npm install yjs
```

不使用协作的普通查看和编辑不需要此 peer 依赖。宿主还需自行安装所选的 provider 和 awareness 实现。如果依赖树中存在多个 Yjs 版本，请进行去重：来自不同运行时副本的类型无法安全地混用。

## 使用共享协作约定 {#using-the-shared-collaboration-contract}

Word、Excel 和 Visio 接收来自 `ooxml-core/collab` 的同一种 `collaboration` 选项形状（`CollaborationOptions`：`roomId`、`serverUrl`、`provider`、`session`、`user`、`intent`）。PowerPoint 在其 `collaboration` 属性中除自有的 `CollaborationConfig` 外，也接受相同的形状，因此嵌入多个产品的宿主可以向它们传入同一个会话：

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({ roomId, provider: transportProvider({ transport }), user });
// <PowerPointViewer collaboration={{ session, user: { name: 'Ada' }, intent: 'join' }} />
```

- `session` 是借用的：它会成为 `externalSession`，查看器绝不会销毁它。会话由宿主创建、连接和销毁。`roomId` 默认为该会话的房间。
- `user.name`、`color`、`avatar` 和 `role` 分别映射到 `userName`、`userColor`、`userAvatar` 和 `role`；`intent` 映射到 `sessionIntent`。当两者同时设置时，显式的旧版字段优先。
- 每个绑定在读取该属性的位置都会运行同一个 `resolveCollaborationConfig`（来自 `ooxml-ui/pptx`），因此行为一致。旧版 `CollaborationConfig` 会原样返回。

没有 `session` 时，同样的选项也涵盖宿主不想为其构建会话的情形：

```ts
// 你自己的传输层：查看器围绕它打开并销毁会话。
collaboration = { roomId, provider: transportProvider({ transport }), user: { name: 'Ada' } };
// 无服务器：此浏览器的标签页和窗口通过 BroadcastChannel 共享。
collaboration = { roomId, user: { name: 'Ada' } };
```

传输方式按共享的顺序选择：`session`（借用，绝不销毁）> `provider` > `serverUrl` > BroadcastChannel。对于 `provider` 和 BroadcastChannel 房间，查看器会调用 `openCollaborationSession`（仅在此时才导入 `ooxml-core/collab` 运行时，并共用同一份 Yjs 副本），通过宿主 `session` 所走的同一条借用会话路径呈现该会话，并在停止、卸载或配置变化时销毁它。会话在同步后即可写入，或者对于网状网络中的单独节点，在宽限期之后可写入；`intent: 'join'` 仍会等待房间中的演示文稿，`role: 'viewer'` 保持只读。会话打开期间状态报告为 `connecting`，房间名称无效或浏览器不支持 BroadcastChannel 时报告为 `error`。

旧版 `CollaborationConfig` 保留其传输方式：`serverUrl` 选择 y-websocket，`transport: 'webrtc'` 或空白的 `serverUrl` 选择无服务器的 y-webrtc（它也会同步同一浏览器的标签页）。只有既未指定 `serverUrl` 也未指定 `transport` 的共享约定输入才会使用 BroadcastChannel 房间。查看器的 presence 与会话的身份心跳共用 awareness 的 `presence` 字段；查看器会在心跳之间保留其幻灯片和光标载荷，因此远程光标会持续显示。

## 报告共享状态 {#reporting-the-shared-state}

每个绑定报告的 `CollaborationState` 都与 Word、Excel 和 Visio 相同（`active`、`roomId`、`status`、`synced`、`people`，其中本地用户标记为 `self`；类型来自 `ooxml-core/collab`，各包也会重新导出）。它在会话启动或停止时（关闭状态，`status: 'off'`）以及状态、同步标志或人员变化时触发，对宿主自有和查看器自有的会话一视同仁，并且绝不会重复发送未变化的状态：

| 绑定    | 通知方式                                                 |
| ------- | -------------------------------------------------------- |
| React   | `onCollaborationChange` 属性                             |
| Vue     | `collaboration-change` 事件                              |
| Angular | `collaborationChange` 输出                               |
| Svelte  | `oncollaborationchange` 回调属性                         |
| Vanilla | `onCollaborationChange` 选项和 `getCollaborationState()` |

Vanilla 保留 `onCollaborationStatus` 和 `getCollaborationStatus()`，用于获取原始连接状态。

## 就绪状态由宿主决定 {#readiness-is-the-hosts-decision}

`getSnapshot()` 返回 `{ status, synced }`。`status` 为 `connecting`、`connected`、`disconnected` 或 `error`。`synced` 表示权威文档已可用，查看器可以发布编辑。它与网络状态相互独立：

- 在加载或填充共享文档期间保持 `synced: false`。不会有任何超时代替宿主启用写入。
- 初始同步完成后设置 `synced: true`。房间中已有的幻灯片会先被采用，之后才可能发布任何本地引导演示文稿。
- 如果编辑可以在宿主的 Yjs provider 中排队，则在离线期间保持为 true。
- 设为 false 会使编辑器变为只读，并在新一轮同步期间暂停持久化回写。采用权威房间状态后，编辑会恢复。

已接受的内联文本和手势预览会同步写入借用的文档，因此之后的就绪状态变化不会丢弃已排队的最后一次按键。宿主 provider 仍可对网络流量进行批处理。就绪状态被撤销时，进行中的文本编辑会关闭；恢复时会采用权威房间状态，而不是重放过期的本地草稿。

对被允许为空房间写入初始内容的参与者使用 `sessionIntent: 'create'`，对启动时的演示文稿不得写入初始内容的参与者使用 `'join'`。显式的“文件 > 打开”仍然是有意的替换。房间的创建需由宿主协调；这不是服务器端授权，也不是选举协议。

## 所有权与持久化 {#ownership-and-persistence}

查看器绝不会创建第二个 provider、断开宿主的 provider，或销毁所提供的文档或 awareness。卸载时，它会移除自己的监听器，并仅恢复它仍然拥有的 presence 字段。无关的 awareness 字段和宿主之后的写入都会保留。每个 awareness 实例请只使用一个活动的幻灯片 presence 写入者。

宿主可以直接持久化 Yjs 更新。或者，指定一名参与者使用 `role: 'owner'` 和 `onWriteBack(bytes)` 来接收经防抖的 PPTX 快照，其源 PPTX 和保存选项与编辑器相同。如果会话离开或宿主撤销就绪状态，待处理和进行中的快照都会被取消。写入者的选举仍由宿主负责。可以从 `onWriteBack` 返回一个 promise 来确认已存储：查看器会在 `collaborationState.writeBack` 中报告 `saving`、`saved` 或 `failed`（仅限 owner），它会串行化写入（同一时间只有一个在进行，之后的更改合并为一次后续写入），并按 `writeBackRetry`（`{ attempts, baseDelayMs?, maxDelayMs? }`）的设定重试被拒绝的写入；直接返回则视为立即确认。待处理的回写会在 `pagehide` 以及页面变为隐藏时刷新。页面卸载时浏览器不会等待异步工作，因此需要网络往返的写入在关闭时仍可能丢失。如果宿主的 `write` 必须在关闭后继续完成，应使用 `navigator.sendBeacon(url, bytes)` 或 `fetch(url, { method: 'PUT', body: bytes, keepalive: true })` 发送字节，浏览器会在页面消失后完成发送（keepalive 请求体上限约为 64 KiB，因此大型文档仍可能无法送达）。
处于活动状态且 `role: 'viewer'` 的会话会使编辑器变为只读并阻止发布，对内置和宿主自有协作均是如此。服务器必须独立地强制执行访问控制。

## 集成边界 {#integration-boundaries}

此选项改变的是会话的所有权，而不是现有的协作数据模型。查看器拥有 `pptx:slides` 及其嵌套的 Yjs 类型。请为所有参与者提供相同的源 PPTX，以便共用包资源、主题和媒体。其他应用数据可以占用单独的 Yjs 键。不要在宿主端独立协调完整的幻灯片快照。

请使用有代表性的演示文稿检查宿主的消息和持久化限制，包括初始文档更新以及编辑后的 `Y.encodeStateAsUpdate(doc)`。压缩后的 PPTX 文件大小并不等于其 Yjs 快照大小：共享 schema 还携带解析后的 XML、形状/组元数据和资源载荷。仅将图片移出文档可能并不足够。此会话 API 不会增加传输分块，也不会改变宿主的限制。

它不会增加评论后端集成、新的仅本地撤销实现，也不会同步每一个演示文稿级字段。在将这些工作流暴露到协作产品中之前，请单独验证。在原生 PowerPoint 中重新打开同样是与两个浏览器对等端收敛不同的另一项保真度检查。

## 试用演示 {#trying-the-demos}

启动演示中继（`bun run collab`）和一个框架演示，然后打开：

```text
/?externalSession=1&room=example&sample=1&server=ws://localhost:1234
```

第二位参与者使用相同的房间，但不带 `sample=1`。宿主栏提供就绪状态和编辑器挂载控件，同时保留文档和 provider。演示会禁用 BroadcastChannel，因此同步走 WebSocket 中继。

添加 `&headless=1` 即可在五个演示中的任意一个里试用自定义界面。这些示例使用绑定的公共构建块，而不会挂载隐藏的完整查看器。其宿主自有的“保存”按钮调用公共序列化路径；缩放选择器可让你在不同缩放比例下检查远程光标和选区轮廓。

各自的组合方式请参阅 [React 钩子](/zh/react/hooks)、[Vue 协作](/zh/vue/collaboration)、[Angular 协作](/zh/angular/collaboration)、[Svelte 协作](/zh/svelte/collaboration)和 [Vanilla 自定义外壳](/zh/vanilla/api#custom-shell-collaboration)。

关于围绕此选项构建的完整应用（账户、经权限检查的中继、实时访问变更、所有者版本和受控的 AI 路由），请参阅[在你的应用内审阅](/zh/guide/review-workflow)。
