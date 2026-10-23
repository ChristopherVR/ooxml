# 协作 {#collaboration}

“文件 > 共享”会在同一浏览器的多个窗口之间启动实时会话。宿主应用也可以通过 `collaboration` 属性跨设备共享，该属性与其他所有 Office 编辑器接受相同的选项。

::: warning 部分支持
此版本没有托管服务。使用“文件 > 共享”时，任何内容都不会离开你的浏览器。跨设备时需要自备服务器（兼容 y-websocket 的 URL）、提供程序或 Yjs 会话；页面和形状按条目合并，其余内容整体共享（见下文）。
:::

## 嵌入：`collaboration` 属性 {#embedding-the-collaboration-property}

`<visio-viewer>` 实现了 `ooxml-core/collab` 的共享嵌入约定（`CollaborationOptions`、`CollaborationState`、`CollaborativeEditor`），与 Word、Excel 和 PowerPoint 使用的相同，因此嵌入多个产品的宿主只需集成一次。

```ts
viewer.collaboration = {
	roomId: 'floor-plan', // 1 to 128 letters, digits, hyphens or underscores
	serverUrl: 'wss://example.com/collab', // optional: a y-websocket compatible server
	user: { name: 'Ada', color: '#2563eb' }, // optional; `role: 'viewer'` makes this window read-only
	intent: 'join', // optional: never seed an empty room with the drawing already open
};
viewer.addEventListener('collaboration-change', (event) => console.log(event.detail));
viewer.collaborationState; // { active, roomId, status, synced, people }
viewer.reconnectCollaboration();
viewer.stopCollaboration(); // or: viewer.collaboration = null
```

- **传输**：优先使用 `session`（宿主自有的 `CollabSession`，查看器绝不会销毁它），其次 `provider`，再次 `webrtc`，再次 `serverUrl`，最后是此浏览器窗口之间的 `BroadcastChannel`。
- **声明式**：只要设置了 `collaboration`，已打开的图纸就会被共享，之后加载的图纸同样会被共享。设为 `null` 即离开房间。查看器自己打开的会话在离开时会被销毁，宿主的 `session` 则不会。
- **意图（`intent`）**：使用 `join` 时，空房间不会用本地图纸初始化，以免迟到的加入者覆盖尚未加载完成的房间。之后加载或编辑的图纸仍会发布。默认值为 `create`。
- **角色**：`role: 'viewer'` 会让查看器在房间内保持只读（离开时恢复 `canEdit`）。角色仅作提示，必须由服务器强制执行。
- **事件**：共享开始或结束，以及连接状态、首次同步或参与者发生变化时，`collaboration-change` 会携带一个 `CollaborationState`。加入失败会作为 `document-error` 报告。“文件 > 共享”使用同一个控制器，因此同样会触发该事件。
- **绑定**：每个绑定都转发 `collaboration` 属性（React、Vue、Solid、Svelte、vanilla 的 `mountViewer`/`update`、Angular 的 `[collaboration]`），并转发 `collaboration-change`（`events['collaboration-change']`、Vue 的 `@collaboration-change`、Angular 的 `(collaborationChange)`）。卸载时会离开房间。
- **类型**：`VisioCollaborationOptions`（共享选项加上 `onWriteBack`、`writeBackDebounceMs` 和 `writeBackRetry`）与 `VisioCollaborationState` 从 `ooxml-ui/visio` 导出。

## 服务器与点对点：令牌和 WebRTC {#server-and-peer-to-peer-auth-token-and-webrtc}

```ts
// A y-websocket compatible server. The token is sent as the `token` query parameter:
// wss://example.com/collab/floor-plan?token=<authToken>
viewer.collaboration = {
	roomId: 'floor-plan',
	serverUrl: 'wss://example.com/collab',
	authToken: await getToken(),
};

// Peer to peer over y-webrtc. Pass the `WebrtcProvider` class from your own y-webrtc install;
// the viewer does not bundle it. The room password defaults to `authToken`.
import { WebrtcProvider } from 'y-webrtc';
viewer.collaboration = {
	roomId: 'floor-plan',
	authToken: 'room password',
	webrtc: { WebrtcProvider, signaling: ['wss://signal.example.com'] },
};
```

查看器会把这些选项原样交给 `openCollaborationSession`。令牌只用于向你运行的服务器认证（服务器必须校验它），WebRTC 密码只是把没有密码的对等端挡在房间之外；二者都不会让角色变得比提示更强。点对点房间从不确认首次同步，因此由会话的宽限期决定此窗口何时可以写入。

## 实时光标与选区 {#live-cursors-and-selections}

同一页面上的协作者会在画布上看到彼此的指针（以此人的颜色显示其姓名的光标）和所选形状（以其颜色显示的虚线轮廓）。它们随图纸的缩放和平移而移动、缩放，并在此人离开、切换页面、将指针移出页面、清除选区或超时失效时消失。

每个窗口发布 `{ page, x?, y?, selection? }`：页面 ID、以页面绘图单位（英寸）表示的指针位置，以及最多 64 个所选形状 ID。`sanitizeVisioPresence`（来自 `ooxml-core/visio/ui`）会校验每个传入的载荷。查看器会把它应用于自己打开的会话。**宿主自有的 `session` 应使用它创建**（并把 `VISIO_INITIAL_PRESENCE` 作为 `initialPresence`），否则其对等端不携带载荷，也就不会绘制光标和选区（不会出错）。

限制：光标和选区只在双方都打开的页面上显示；选区轮廓依据形状已保存的外框，而不是其渲染后的几何形状。

## 所有者回写 {#owner-write-back}

希望保存共享图纸的宿主可以设置 `onWriteBack`，并可选设置 `writeBackDebounceMs`（默认 5000）：

```ts
viewer.collaboration = {
	roomId: 'floor-plan',
	user: { name: 'Ada', role: 'owner' },
	onWriteBack: (bytes) => upload(bytes), // VSDX bytes
	writeBackDebounceMs: 5000,
};
```

只有角色为 `owner` 的窗口才会写入，因此协作者不会争着保存。图纸在防抖时间内保持静止后（无论是本地还是远程更改之后）会触发该回调；若有待保存内容，所有者离开房间时还会再触发一次。对 `collaborator` 和 `viewer` 永不触发。从 `onWriteBack` 返回 Promise 可确认已存储：同一时间只有一次写入在进行，期间的更改合并为一次后续写入，设置了 `writeBackRetry`（`{ attempts, baseDelayMs?, maxDelayMs? }`）时被拒绝的写入会重试。`collaborationState.writeBack`（也出现在 `collaboration-change` 中）报告 `{ status: 'idle' | 'pending' | 'saving' | 'saved' | 'failed', lastSavedAt?, error? }`，仅所有者会报告。待保存内容会在离开房间、`pagehide` 以及页面隐藏时提交；页面关闭时异步写入仍可能丢失。需要在关闭后仍能完成的写入，请使用 `navigator.sendBeacon` 或 `fetch(..., { keepalive: true })`（约 64 KiB 上限）。

## 目前已有的功能 {#what-exists-today}

- 基于 `ooxml-core/collab`（Yjs）并通过 `BroadcastChannel` 构建的实时会话，因此同一浏览器的窗口和标签页可以共享同一份图纸；宿主可通过 `collaboration` 跨设备共享。
- 页面、形状（细到单元格、行和节）、连接线、母版和图片按条目合并（`ooxml-core/visio/collab`）；包中其余部分（样式、主题、数据、文档设置）整体共享，每个部分以最后一次写入为准。协作运行时仅在开始共享时才加载。
- 对不同形状的并发编辑都会保留；对同一形状的两次编辑在所有窗口中得到同一个值；被删除的形状或页面胜过对它的并发编辑。
- 撤销和重做（Ctrl+Z、功能区）只作用于此窗口自己的编辑，绝不会撤销协作者的更改。远程更改在你的撤销历史之外应用，历史从加入时开始。
- 协作者通过共享的 `ooxml-ui` 在线状态元素显示。
- 会话名称由 1 到 128 个字母、数字、连字符或下划线组成。加入一个房间已有图纸的会话时，窗口会采用该图纸；否则它会初始化房间（除非宿主要求 `intent: 'join'`）。
- 面板会话保留其原有的 BroadcastChannel 名称（`ooxml-collab:visio-<room>`），因此旧版本构建的窗口仍可相互找到。

## 实时演示：两个窗口，任意框架 {#live-demo-two-windows-any-framework}

共享会话只是一个名称。下面的两个窗格是由不同框架绑定构建的独立演示环境，它们在加载时都会加入会话 `docs-collab`：窗口 A 是原生演示环境并打开示例，窗口 B 是 React 演示环境并从房间接收图纸。在一个窗格中编辑形状文本，另一个窗格会随之更新；撤销会还原远程步骤。

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer window A (vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer window B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### 任意搭配两个框架 {#pair-any-two-frameworks}

每个演示环境都接受 `?share=<session>`，它会填入“文件 > 共享”并启动共享。只有会话名称起作用，因此可以任意搭配：主页的[实时演示](/zh/#live-demo)为每个窗口提供了框架选择器；在两个标签页中，你也可以打开例如 `/demo-vue/?sample=1&share=my-room` 和 `/demo-angular/?share=my-room`。你还可以在任意两个窗口中手动操作：文件，然后共享，再输入相同的会话名称。

### 为什么“文件 > 共享”仅限同一浏览器 {#why-it-is-same-browser-only}

该面板使用 `BroadcastChannel`，它只连接同一来源、同一浏览器配置文件中的窗口、标签页和框架，与所用框架无关，除此之外别无其他。两台不同的机器无法以这种方式相遇；为此，宿主需要用 `serverUrl`、`provider` 或 `session` 设置 `collaboration`，而此版本不附带任何服务器。

## 支持与不支持的功能 {#what-is-and-is-not-supported}

| 支持                                                              | 不支持                                                         |
| ----------------------------------------------------------------- | -------------------------------------------------------------- |
| 基于 Yjs、通过 BroadcastChannel 的同浏览器会话                    | 托管服务器：请自备 y-websocket 服务器、提供程序或会话          |
| 按形状、按单元格和按页面合并，删除优先，撤销仅限本地              | 单元格内部的合并：同一单元格的两次编辑只保留一个值             |
| 使用共享在线状态元素显示协作者                                    | 批注；角色仅作提示（必须由服务器强制执行）                     |
| 以“文件 > 共享”和 `collaboration` 属性作为入口                    | 跨不同页面的光标（只绘制共享的页面）                           |
| 同一页面上协作者的实时光标和选区                                  | 未使用 `sanitizeVisioPresence` 创建的宿主 `session` 的在线状态 |
| 所有者回写：防抖、Promise 确认、重试、状态                        | 保证异步写入在页面关闭时完成                                   |
| 用于服务器的 `authToken` 和 `webrtc`（你自己的 `WebrtcProvider`） | 自带 y-webrtc 或信令服务器；令牌校验（由你的服务器负责）       |

### 形状级合并的限制 {#limits-of-shape-level-merging}

- 两个人编辑同一形状的不同属性（移动与文本或填充更改、不同的单元格、行或节），双方的更改都会保留。两人编辑同一单元格或同一文本时，所有窗口最终得到同一个值。完整的 `Text` 或 `Data` 子节点各自作为一个整体。母版内部的编辑以同样的方式合并（按形状和单元格），且一个人删除的母版会胜过其内部的并发编辑。
- 页面部件按文本拆分，因此单元格会与编辑器写入时完全一致地往返（元素之间的空白会被丢弃）。无法拆分的页面部件（没有可用的 `Shapes` 列表）会退回为把它的形状作为页面外壳的一部分共享，最后写入者胜出。
- 母版和媒体按项合并：两个人同时添加的母版都会保留；图片按内容区分，因此相同的图片只存储一次，而同名的不同图片都会保留（其中一个会被重命名）。样式、主题、文档设置、数据和内容类型部件仍按部件合并，最后写入者胜出。窗口状态和缩略图保持本地。
- 两个窗口同时添加相同 ID 的形状时，其中一个会被重新编号；相同的房间内容在每个窗口上总会得到相同的编号。`Sheet.n!...` 公式会跟随为其编写的形状，无论何时、在何处输入（每个公式都记住它指向哪个形状，因此即使在冲突的那一刻输入、指向较早的形状或由其他人输入，也仍指向其作者所指的形状）。能够写入的窗口随后会把编号保存到房间中，移动的形状和保留编号的形状都是如此，因此之后添加更多形状时两者的编号都不会改变，即使是同一时刻以相同编号添加的形状也一样。撤销从不会回退这一步，但撤销公式编辑仍会将其回退。只有一种情况仍可能让编号再变动一次：两个窗口在同一时刻为一个房间保存了不同的编号，而该房间还包含另一个窗口尚未看到的添加；此后房间会在每个窗口上立即稳定下来。页面和母版的重新编号方式相同（随母版添加的形状会跟随母版）。
- 一个人移除、另一个人以相同方式重新建立的连接（glue）会保留。已移除媒体的图片仍留在房间中。一个人删除某个单元格、母版或图片时，若另一个人正在编辑它，则保留该编辑；被删除的形状或页面始终优先。
- 由其他版本的协作格式架构写入的房间会被拒绝：窗口会报告错误，并且既不读取也不写入该房间。请同时升级同一房间的所有窗口。
- 撤销历史按窗口保存，从加入时开始，重新连接后不会保留。同时到达的远程更改会以一次重建应用，使图纸保持不变的更改则不应用任何内容。但当只有部分页面的内容发生了变化（且你在此期间没有编辑），则只重新读取这些页面，其余页面保留（样式、主题和母版在未变化时不会被重新读取，包中未变化的部分也不会被重新计算校验和）。页面列表、母版、样式、主题、文档设置或图片的变化、你采用的第一份图纸，以及你自己编辑之后的第一次远程更改，都会重新读取整个包，因此在这些情况下很大的图纸仍会被完整读取。

相关证据记录在[功能清单](/parity)的“Collaboration and comments”下。
