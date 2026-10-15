# 协作

协作正是 OpenTeams 的本质：每个频道、每条消息、每个回应和每个在线状态信号，都是多人同时编辑的共享状态。本页说明它是如何构建的，在线演示展示了什么，以及没有展示什么。

## 工作原理

- **共享文档是一个 Yjs CRDT。** 频道存放在一个 `Y.Map` 中，消息按频道各存放在一个 `Y.Map` 中，回应存放在以消息、表情和用户为键的扁平映射中，因此并发的编辑和回应会合并而不是相互覆盖。相关逻辑位于 `ooxml-core/teams`，构建在与格式无关的 `ooxml-core/collab` 区域之上。
- **在线状态是 Yjs awareness**：谁在线、谁正在输入，以及你自己设置的状态（可用、忙碌、离开）。它只是参考性的，且未经身份验证，与任何 Yjs awareness 一样。
- **通话基于 WebRTC。** 信令经由你的服务器传输（本地模式下经由 `BroadcastChannel`）；媒体以全网状方式点对点传输，因此通话适合少数几个人。
- **传输层可插拔。** 有两种模式：`server`（你的服务器使用 y-websocket 同步协议、中继通话信令并可选地存储文件）和 `local`（同一浏览器的标签页通过 `BroadcastChannel` 共享一切）。

## 在线演示：两个人，一个房间，任意框架

下面的两个窗格是同一个本地房间中的两个独立演示应用。Ada 运行原生的 `<teams-app>`，Bob 运行 React 绑定，但他们共享同一个工作空间：房间只是一个名称（`?room=`），每个绑定使用相同的核心协议。它们运行在本地模式，因此只通过你的浏览器相互通信。以 Ada 的身份输入一条消息，看着它出现在 Bob 那里。

在首页的[在线演示](/zh/#live-demo)中，每个窗格都有自己的框架选择器，因此你可以将 Vanilla、React、Vue、Angular、Svelte 和 Solid 中的任意两个配对。

<iframe
	src="/ooxml/teams/demo/?name=Ada&room=docs-collab"
	title="OpenTeams demo as Ada"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/teams/demo-react/?name=Bob&room=docs-collab"
	title="OpenTeams React demo as Bob"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

在真实的标签页之间，任意配对同样可行：打开[以 Ada 身份运行的原生演示](/demo/?name=Ada&room=docs-collab){target="_self"}，再打开例如[以 Bob 身份运行的 Vue 演示](/demo-vue/?name=Bob&room=docs-collab){target="_self"}。每个演示都接受相同的 `name` 和 `room` 参数；本地模式下只有同一浏览器的框架和标签页才能相遇，因此对于位于不同机器上的人，需要让每个绑定指向你自己的服务器（见下文）。

## 支持与不支持的内容

| 支持                                               | 不支持                                           |
| -------------------------------------------------- | ------------------------------------------------ |
| 由 Yjs CRDT 合并的并发聊天、回复和回应             | 端到端加密：你的服务器可以读取并存储一切         |
| 在线状态和输入提示                                 | 经过验证的身份：名称只是声明，参考服务器没有账户 |
| 无需服务器的同浏览器协作（本地模式，如上所示）     | 没有你自己运行的服务器时，不同机器之间的协作     |
| 通过你自己的同步服务器实现跨机器协作               | 托管服务：不提供，Pages 上的演示背后也没有服务器 |
| 音频、视频、屏幕共享和举手的通话（网状，约 12 人） | 网络研讨会规模的通话、录制、通知                 |

## 跨机器

运行参考服务器（`npx openteams-server`）或任何遵守这三个约定（同步、信令、可选的文件）的服务器，并将 `config.mode` 设置为 `'server'`。参见[自带服务器](/zh/server)和[部署](/zh/deploy)。Office 文件的协同编辑不属于 OpenTeams：文件以链接形式共享，宿主可以在 Word、Excel、PowerPoint 或 Visio 查看器中打开它们，各自有自己的协作方式。
