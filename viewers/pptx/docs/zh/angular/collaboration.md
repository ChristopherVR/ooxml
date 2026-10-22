---
title: 实时协作
description: 通过 Yjs CRDT 为 PowerPointViewerComponent 提供多人实时协同编辑，包括 collaboration 输入、CollaborationConfig 结构、在线状态和远程光标。
---

# 协作 {#collaboration}

对于应用已有的 Yjs 提供程序，请使用 `collaboration.externalSession`。共享约定请参见[宿主自有协作](/zh/guide/host-owned-collaboration)。

`PowerPointViewerComponent` 基于 **Yjs**（一种 CRDT）支持多人实时编辑，可使用 WebSocket 传输（`y-websocket`，需要服务器），也可使用无文档服务器的点对点传输（`y-webrtc`）。启用后提供细粒度 CRDT 文档同步（按幻灯片、元素和字段）、实时远程光标、选区高亮、用户在线状态和跟随模式。单用户模式不会加载这些内容。

::: info 可选 peer 依赖
协作需要 `yjs` peer 依赖，以及所选传输方式对应的提供程序：`y-websocket`（基于服务器）或 `y-webrtc`（点对点）。不安装它们也能完整使用查看器，只是以单用户模式运行。需要协同编辑时再安装：

```bash
npm i yjs y-websocket   # server-based
npm i yjs y-webrtc      # serverless peer-to-peer
```

:::

## 启用协作：`collaboration` 输入 {#enabling-it-the-collaboration-input}

将 `CollaborationConfig` 传入 `collaboration` 输入。提供配置后，查看器通过内部 `CollaborationService` 接入在线状态跟踪、远程光标和 CRDT 同步。

```ts
import { Component, signal } from '@angular/core';
import { PowerPointViewerComponent } from 'pptx-angular-viewer';
import type { CollaborationConfig } from 'pptx-angular-viewer';

@Component({
	standalone: true,
	imports: [PowerPointViewerComponent],
	template: `<pptx-viewer [content]="bytes" [canEdit]="true" [collaboration]="config()" />`,
})
export class Example {
	readonly config = signal<CollaborationConfig>({
		roomId: 'my-room-123',
		serverUrl: 'wss://collab.example.com',
		userName: 'Alice',
		userColor: '#6366f1',
	});
}
```

## `CollaborationConfig` {#collaborationconfig}

```ts
type CollaborationRole = 'owner' | 'collaborator' | 'viewer';
type CollaborationTransport = 'websocket' | 'webrtc';

interface CollaborationConfig {
	/** Unique room id (alphanumeric, hyphens, underscores). */
	roomId: string;
	/** WebSocket URL for the Yjs provider, e.g. "wss://collab.example.com". Ignored for webrtc. */
	serverUrl: string;
	/** Transport - 'websocket' (default) or serverless 'webrtc'. */
	transport?: CollaborationTransport;
	/** WebRTC signaling server URLs (webrtc transport only). */
	signaling?: string[];
	/** Display name for the local user. */
	userName: string;
	/** Avatar URL for the local user (optional). */
	userAvatar?: string;
	/** Hex colour for the local user's cursor / presence indicator. */
	userColor?: string;
	/** Optional auth token sent with the WebSocket handshake / used as the webrtc room password. */
	authToken?: string;
	/** Session role - defaults to 'collaborator'. */
	role?: CollaborationRole;
	/** Elected-writer persistence: the 'owner' peer receives debounced PPTX snapshots. */
	onWriteBack?: (bytes: Uint8Array) => void | Promise<void>;
	/** Debounce (ms) between the last change and onWriteBack. Default 5000. */
	writeBackDebounceMs?: number;
	/** Retries of a rejected onWriteBack, exponential backoff. Default: none. */
	writeBackRetry?: { attempts: number; baseDelayMs?: number; maxDelayMs?: number };
}
```

| 字段                  | 类型                               | 必填 | 说明                                                                                                                |
| --------------------- | ---------------------------------- | ---- | ------------------------------------------------------------------------------------------------------------------- |
| `roomId`              | `string`                           | 是   | 会经过清理，建议仅使用字母、数字、`-` 和 `_`。                                                                      |
| `serverUrl`           | `string`                           | 是   | `y-websocket` 服务器 URL；`transport: 'webrtc'` 时可以为 `''`。                                                     |
| `transport`           | `CollaborationTransport`           | 否   | 默认 `'websocket'`，也可选 `'webrtc'` 点对点模式。                                                                  |
| `signaling`           | `string[]`                         | 否   | y-webrtc 信令 URL，默认使用其公共列表。                                                                             |
| `userName`            | `string`                           | 是   | 本地用户显示名称；未设置 `authorName` 时也用作批注和标记作者。                                                      |
| `userAvatar`          | `string`                           | 否   | 经过校验的头像 URL。                                                                                                |
| `userColor`           | `string`                           | 否   | 用户光标环的十六进制颜色。                                                                                          |
| `authToken`           | `string`                           | 否   | WebSocket 握手参数或 WebRTC 房间密码。                                                                              |
| `role`                | `CollaborationRole`                | 否   | `'owner'`、默认的 `'collaborator'` 或 `'viewer'`。                                                                  |
| `onWriteBack`         | `(bytes) => void \| Promise<void>` | 否   | 仅在 `'owner'` 端调用，提供防抖后的序列化 PPTX 快照；返回的 Promise 即确认，状态见 `collaborationState.writeBack`。 |
| `writeBackDebounceMs` | `number`                           | 否   | 默认 5000 ms。                                                                                                      |

### 无需文档服务器的点对点模式 {#serverless-peer-to-peer-mode}

设置 `transport: 'webrtc'` 后，参与者通过 WebRTC 直接交换更新，不需要文档服务器。**同一浏览器**的标签页通过 BroadcastChannel 连接，即使没有网络也可使用。GitHub Pages 演示采用这种方式。跨设备会话通过 WebRTC 信令服务器发现彼此，信令只传元数据，不经过文档内容；生产环境可以通过 `signaling` 指定自己的服务。内置共享和广播对话框中，服务器地址留空会选择此传输。

::: warning 输入会经过清理
房间 ID、用户名、头像 URL、光标位置和在线状态数据都会经过协作层（`collaboration-helpers.ts`）的输入清理。建议 `roomId` 仅使用字母、数字、连字符和下划线，避免意外变化。
:::

## 使用自己的提供程序，或在同一浏览器内共享 {#your-own-provider-or-sharing-in-one-browser}

除 `CollaborationConfig` 外，`collaboration` 选项还接受 Word、Excel 和 Visio 共用的约定（来自 `ooxml-core/collab` 的 `CollaborationOptions`）。以下两种形式不需要宿主自有的 `session`；查看器自行打开会话，并在协作停止时销毁它：

```ts
// provider created by the host
collaboration = { roomId: 'deck-review', provider, user: { name: 'Ada' }, intent: 'join' as const };
// Same browser only:
collaboration = { roomId: 'deck-review', user: { name: 'Ada' } };
```

```html
<pptx-power-point-viewer [content]="bytes" [collaboration]="collaboration" />
```

- `provider`（`ProviderFactory`，例如 `transportProvider`，或通过 `adaptYjsProvider` 包装的现有提供程序）优先于 `serverUrl`。
- 未提供 `session`、`provider`、`serverUrl` 或 `transport` 时，房间通过 BroadcastChannel（`ooxml-pptx:<roomId>`）在本浏览器的标签页和窗口之间共享，无需服务器。
- 优先级为 `session`（借用，从不销毁）> `provider` > `serverUrl` > BroadcastChannel。带 `serverUrl` 的旧版 `CollaborationConfig` 仍使用 y-websocket；带 `transport: 'webrtc'` 或空 `serverUrl` 的则仍使用无服务器的 y-webrtc 传输。
- `intent: 'join'` 绝不会用本地演示文稿初始化空房间。会话在同步完成后（或单个参与者经过宽限期后）可写；`role: 'viewer'` 保持只读。
- 会话打开期间状态为 `status: 'connecting'`；房间无效或 BroadcastChannel 不可用时报告 `error`。

## 控制会话：共享与广播对话框事件 {#controlling-sessions-share-broadcast-dialog-events}

与 React 和 Vue 一样，协作状态由宿主应用**控制**。查看器的共享和广播对话框通过输出事件表达用户意图，应用据此切换 `collaboration` 输入。

| 输出或输入            | 类型                                        | 用途                                                                                                   |
| --------------------- | ------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| `startCollaboration`  | `output<CollaborationConfig>`               | 用户从共享或广播对话框启动了会话，请用此配置设置 `collaboration` 输入。                                |
| `stopCollaboration`   | `output<void>`                              | 用户停止了会话，请清空 `collaboration` 输入。                                                          |
| `collaborationChange` | `output<CollaborationState>`                | 共享状态（`active`、`roomId`、`status`、`synced`、`people`），宿主自有会话也会触发；停止时为关闭状态。 |
| `shareDefaults`       | `input<{ roomId?; userName?; serverUrl? }>` | 预填共享对话框字段；`userName` 回退为 `authorName`，其余未提供字段为空。                               |

```ts
@Component({
	template: `
		<pptx-viewer
			[content]="content"
			[canEdit]="true"
			[collaboration]="collab()"
			[shareDefaults]="{ serverUrl: 'wss://collab.example.com', userName: 'Alice' }"
			(startCollaboration)="collab.set($event)"
			(stopCollaboration)="collab.set(undefined)"
		/>
	`,
})
export class CollaborativeEditorComponent {
	readonly collab = signal<CollaborationConfig | undefined>(undefined);
}
```

## 在线状态和远程光标 {#presence-and-remote-cursors}

会话期间，组件显示：

- 幻灯片画布上的**远程光标**（`CollaborationCursorsComponent`），每个光标标有用户姓名和颜色。
- **远程选区高亮**（`RemoteSelectionOverlayComponent`），围绕其他参与者选中的元素显示。
- **跟随模式横幅**（`FollowModeBarComponent`），让本地用户跟随其他参与者当前的幻灯片。

在线状态数据通过 Yjs 的 _awareness_ 广播。每位参与者发布一条状态记录，包括客户端 ID、姓名、颜色、当前幻灯片索引、限制在有效范围内的光标 X/Y 坐标、选中元素 ID、角色和最后更新时间戳，用于驱动光标及在线状态界面。

## 自定义协作界面 {#building-custom-collaboration-ui}

自定义编辑器外壳请将 `ViewerCollaborationShellService` 与 `POWER_POINT_VIEWER_PROVIDERS` 一起使用。它组合现有的加载器、编辑器、会话、画布编辑和光标服务，不会创建第二个协作控制器或传输。在宿主组件的构造函数中调用一次 `bind()`，并传入响应式访问器：

```ts
import { Component, inject, input, viewChild } from '@angular/core';
import {
	POWER_POINT_VIEWER_PROVIDERS,
	SlideCanvasComponent,
	ViewerCollaborationShellService,
} from 'pptx-angular-viewer';
import type {
	CollaborationConfig,
	CollaborationShellState,
	ViewerCollaborationShellOptions,
} from 'pptx-angular-viewer';

@Component({
	selector: 'app-custom-slides',
	standalone: true,
	providers: [...POWER_POINT_VIEWER_PROVIDERS, ViewerCollaborationShellService],
	imports: [SlideCanvasComponent],
	template: `
		<pptx-slide-canvas
			[slide]="shell.activeSlide()"
			[canvasSize]="shell.loader.canvasSize()"
			[mediaDataUrls]="shell.loader.mediaDataUrls()"
			[templateElements]="shell.activeTemplateElements()"
			[editable]="shell.canEdit()"
			[selectedIds]="shell.editor.selectedIds()"
			[editingId]="shell.canvasEditing.editingId()"
			(elementSelect)="shell.canvasEditing.onElementSelect($event)"
			(textEditStart)="shell.canvasEditing.onTextEditStart($event.id)"
			(textInput)="shell.canvasEditing.onTextInput($event)"
			(textCommit)="shell.canvasEditing.onTextCommit($event)"
			(textCancel)="shell.canvasEditing.editingId.set(null)"
			(listSession)="shell.canvasEditing.onListSession($event)"
		/>
	`,
})
export class CustomSlidesComponent {
	readonly content = input<Uint8Array | null>(null);
	readonly collaboration = input<CollaborationConfig>();
	readonly canEdit = input(true);
	readonly shell = inject(ViewerCollaborationShellService);
	readonly canvas = viewChild(SlideCanvasComponent);

	constructor() {
		const options: ViewerCollaborationShellOptions = {
			content: this.content,
			collaboration: this.collaboration,
			canEdit: this.canEdit,
			stageElement: () => this.canvas()?.getStageElement(),
		};
		this.shell.bind(options);
	}

	getState(): CollaborationShellState {
		return this.shell.state();
	}
	getContent(): Promise<Uint8Array> {
		return this.shell.getContent();
	}
}
```

示例接入了文本编辑；请为外壳提供的操作接入画布变换、表格及其他输出。自定义修改控件请使用 `shell.canEdit()` 进行限制，而不是只看宿主请求的权限。保存时若有未完成的原生输入或 IME 组合输入，会被拒绝，而不是序列化较旧的文本快照。实际的限制还会考虑当前会话的角色和就绪状态，以及真实来源仍在加载或加载失败。没有协作的空白编辑器保留普通的宿主控制权限。即使查看者角色或暂停的会话禁用了编辑，`getContent()` 仍会保存当前可编辑模型。

`shell.state()` 提供规范化的 `remoteUsers`、连接 `status` 和 `connectedCount`。对于画布在线状态，请将指针移动转发给 `shell.cursor.onPointerMove($event)`，并将 `CollaborationCursorsComponent` 和 `RemoteSelectionOverlayComponent` 投影到 `SlideCanvasComponent` **内部**。需提供 `shell.cursor.cursors()`、`shell.collaboration.presence()`、当前幻灯片的元素以及 `shell.activeSlideIndex()`。两个叠加层都使用未缩放的幻灯片坐标，缩放由画布统一应用一次。`getStageElement()` 提供实际缩放后的幻灯片原点，无需依赖内部 CSS 选择器。

可运行的自定义外壳示例（含文本和形状编辑、保存、缩放及两个在线状态叠加层）见 `demos/pptx/demo-angular/src/host-owned-headless-editor.component.ts`。[宿主自有协作约定](/zh/guide/host-owned-collaboration)及其同一文本框并发的限制，同样适用于自定义界面和完整查看器界面。

协作服务和组件由包根入口导出，没有像 `pptx-react-viewer/viewer` 那样的独立子路径。需要自行驱动同步或渲染在线状态界面时可以使用：

```ts
import {
	CollaborationService,
	CollaborationCursorsComponent,
	RemoteSelectionOverlayComponent,
	FollowModeBarComponent,
} from 'pptx-angular-viewer';
import type { CollaborationConfig, RemoteCursor, RemotePresence } from 'pptx-angular-viewer';
```

`CollaborationService` 通过信号提供 `status`、`connected`、`active`、`activeRole`、`presence`、`cursors` 和 `connectedCount`，并提供 `connect()`、`retry()`、`disconnect()`、`broadcastSlides()`、`setCursor()`、`setSelection()`、`setActiveSlide()` 和 `followUser()` 方法。完整的底层协作构建模块（`LocalPresencePublisher`、`createWebsocketBundle` / `createWebrtcBundle`、`WriteBackScheduler`）请参见[服务 › 完整服务参考](/zh/angular/services-reference#collaboration-internals)。

## 服务器端 {#server-side}

使用默认 `websocket` 传输时，需要运行与 `y-websocket` 兼容的中继服务器，并能通过 `serverUrl` 访问。`demos/pptx/` 提供两个按生产使用方式设计的参考服务器，五种绑定共享它们，React、Vue、Angular、Svelte 和原生 JavaScript 使用相同的同步结构：

- **`demos/pptx/collab-server.example.mjs`**：无需新增依赖的 Bun 服务器，使用仓库已有的 `yjs` / `y-protocols` / `lib0`。认证有两种模式，均在 WebSocket 握手时验证（协议升级前返回 401）：

  - **JWT 模式**（生产环境）：设置 `COLLAB_AUTH_JWT_SECRET`，由应用服务器签发短期 HS256 令牌。
  - **允许列表模式**（开发环境）：通过 `COLLAB_AUTH_TOKENS=a,b,c` 指定静态令牌。

  `role: 'viewer'` 令牌获得**只读连接**：中继会丢弃其文档写入，从而在服务器端强制只读，无需信任客户端的 `canEdit`。每个房间的 Y.Doc 会快照保存到 `COLLAB_DATA_DIR`，下次加入时恢复，因此服务器重启后文档仍然保留。

  ```bash
  COLLAB_AUTH_JWT_SECRET=change-me bun demos/pptx/collab-server.example.mjs
  ```

- **`demos/pptx/collab-server-hocuspocus.example.mjs`**：在 Node/Hocuspocus 技术栈上实现相同约定（通过 `@hocuspocus/extension-sqlite` 使用 SQLite 持久化）。

两种服务器都会验证各绑定发送的令牌：`CollaborationConfig` 中的 `authToken` 会成为 WebSocket 握手的 `?token=` 查询参数。生产环境应在中继前终止 TLS（使用 `wss://`），并优先选择有效期较短的 JWT 模式。

`transport: 'webrtc'` 不需要文档服务器。

MCP 包提供自己的服务端 Yjs 编解码器（[/zh/packages/mcp](/zh/packages/mcp)）；请注意其 Y.Doc 键布局与查看器各绑定的同步 schema 不同，因此两者不能共用一个 Y.Doc。
