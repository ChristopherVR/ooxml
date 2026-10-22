---
title: 在你的应用内审阅
description: 在由宿主应用拥有账户、权限、会话、持久化和 AI 密钥的情况下，进行协作式演示文稿审阅的实践方案。
---

# 在你的应用内审阅 {#review-inside-your-application}

协作通常在用户已有的工作流中最为有用：交易台、合规队列、CRM 记录、项目看板。多人从该系统中打开同一份演示文稿，手动或借助 AI 助手做出小幅修改，而系统负责记录谁可以做什么，以及哪个版本是当前版本。

此方案沿着一条界线划分工作：

| 你的应用负责                         | 查看器负责                               |
| ------------------------------------ | ---------------------------------------- |
| 登录、账户，以及用户的显示名称和颜色 | 渲染和编辑演示文稿                       |
| 谁可以打开、编辑或仅观看每个条目     | 共享 Yjs 文档中的幻灯片 schema           |
| Yjs 文档、awareness、provider 和中继 | 在线状态（光标、选区、“跟随协作者”）     |
| 会话何时就绪，以及何时不再就绪       | 在你声明房间就绪之前暂缓编辑             |
| 已保存的版本去向何处，由谁写入       | 为所有者将共享的演示文稿转换为 PPTX 字节 |
| AI 模型、其密钥以及谁可以使用        | 助手面板、其演示文稿工具和修改建议       |

[`demos/pptx/demo-host-review`](https://github.com/ChristopherVR/ooxml/tree/main/demos/pptx/demo-host-review) 中的示例应用使用 vanilla 绑定实现了下面的每一步。所有绑定都接受相同的 `collaboration` 和 `ai` 选项，因此这些步骤在 React、Vue、Angular 和 Svelte 中的写法相同。

## 1. 将你的权限映射为角色 {#_1-map-your-permissions-to-a-role}

保留你自己的权限模型。挂载编辑器时，由它推导出查看器的 `role`：

| 你的权限（示例） | `role`           | 说明                                      |
| ---------------- | ---------------- | ----------------------------------------- |
| 条目的所有者     | `'owner'`        | 可编辑，并接收 `onWriteBack` 快照进行保存 |
| 审阅者           | `'collaborator'` | 与其他人实时编辑                          |
| 观看者           | `'viewer'`       | 只读；同时传入 `editable: false`          |
| 无访问权限       | （不要挂载）     | 你的 API 甚至不应透露该条目存在           |

`role: 'viewer'` **仅在浏览器中**强制执行。请同样在服务器上强制执行（第 3 步）。

## 2. 拥有会话 {#_2-own-the-session}

自行创建 Yjs 文档、awareness 和 provider，并将它们作为 `collaboration.externalSession` 传入（完整约定请参阅[宿主自有协作](/zh/guide/host-owned-collaboration)）。查看器绝不会连接、重连或销毁它们。

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

const grant = await api.requestGrant(reviewId); // 你的 API：ticket、role、intent、user
const doc = new Y.Doc();
const awareness = new Awareness(doc);
const provider = new WebsocketProvider(`wss://${location.host}/collab`, reviewId, doc, {
	awareness,
	params: { ticket: grant.ticket },
});

const viewer = createPptxViewer(container, {
	source: sourceBytes, // 每位参与者使用相同的源 PPTX
	editable: grant.role !== 'viewer',
	collaboration: {
		roomId: reviewId,
		serverUrl: '', // 使用外部会话时不会用到
		userName: grant.user.name, // 来自你的目录，而不是由用户输入
		userColor: grant.user.color,
		role: grant.role,
		sessionIntent: grant.intent,
		externalSession: { doc, awareness, getSnapshot, subscribe },
	},
});
```

`getSnapshot()` 返回 `{ status, synced }`。只有当 provider 已完成同步，**并且**用户仍持有挂载编辑器时所用的访问权限时，才报告 `synced: true`。将其设为 `false` 会使编辑器立即变为只读。

**谁来初始化房间。** 空房间由第一位使用 `sessionIntent: 'create'` 的参与者填充；其他所有人传入 `'join'` 并采用已有内容。请由服务器而不是客户端来决定：示例每次只向一位写入者授予 `'create'`，带有短期租约，并且仅在房间仍为空时授予，因此两个人同时打开该条目时，不会都发布各自的启动演示文稿。

## 3. 在服务器上准入并管控套接字 {#_3-admit-and-police-sockets-on-the-server}

WebSocket 无法把你站点的会话 cookie 带到另一个源上的中继，而 URL 中的长期令牌会泄漏到日志中。请改为通过普通的已认证 API 调用签发**短期的签名 ticket**，并让中继只信任 ticket 所声明的内容：

- **升级时**：验证 ticket 的签名、有效期和房间。其他一律拒绝。
- **每条消息**：只读连接可以请求文档（同步步骤 1）并共享 presence，但其同步步骤 2 和更新消息会被丢弃。这是 `role: 'viewer'` 的服务器端一半。
- **Awareness**：让每个连接只能代表它最初声明的 awareness 客户端 ID 发言，这样任何人都无法清除或冒充他人的光标。
- **Ticket 刷新**：ticket 只用于准入连接。请在其过期前刷新并更新 `provider.params.ticket`，以便自动重连时使用有效的 ticket。

你无需自己编写这个中继。`openteams-server`（OpenTeams 查看器的参考服务器，一个兼容 y-websocket 的 Yjs 服务器）提供 `authorizeSync` 钩子，按套接字返回 `'write'`、`'read'` 或拒绝，已经固定了 awareness ID，并暴露了 `closeSync`、`syncPrincipals`、`isSyncRoomEmpty` 和 `handleUpgrade`，供下列步骤使用。示例的 `server/relay.ts` 围绕它只有大约六十行。

## 4. 在人们工作期间更改访问权限 {#_4-change-access-while-people-are-working}

当所有者更改某人的访问权限时，请用应用级关闭码（示例使用 `4403`）关闭该用户已打开的套接字。客户端的反应如下：

1. 报告 `synced: false`，这会冻结编辑器；
2. 销毁查看器，然后销毁其自身的文档和 provider；
3. 请求新的授权，并以新角色挂载全新的会话；若没有授权，则显示“访问已被移除”。

降级之后请从**新的** Y.Doc 开始。用户在变更前不久所做的编辑已被中继丢弃，而重用旧文档会在本地继续显示这些编辑。

## 5. 由单一写入者保存版本 {#_5-save-versions-with-one-writer}

只挂载一位参与者（通常是条目的所有者），使用 `role: 'owner'` 和 `onWriteBack(bytes)`。查看器会对更改进行防抖（`writeBackDebounceMs`，默认 5 秒），并把共享演示文稿的 PPTX 快照交给你，作为一个版本存储。其他参与者从不回写，因此不存在最后保存者获胜的竞争。

`onWriteBack` 并不是持久存储的确认：重试、失败，以及所有者离线时如何选择写入者，仍由你来决定。如果房间必须在中继重启后保留，请同时在服务器上持久化 Yjs 更新（例如使用你自己的存储，或 y-leveldb 之类的 provider 后端）。

## 6. 添加 AI 辅助修改 {#_6-add-ai-assisted-changes}

只向可以使用 AI 的账户传入 `ai` 选项，并将其指向**你自己的**路由，这样模型密钥永远不会到达浏览器：

```ts
ai: {
	connection: {
		kind: 'endpoint',
		api: `/api/ai/chat?review=${reviewId}`,
		credentials: 'same-origin',
	},
	writePolicy: 'stage', // 每项更改都是一个建议，由审阅者接受或拒绝
},
```

该路由检查调用者的账户和权限，然后使用来自 `ooxml-ui/pptx/ai` 的仅含 schema 的演示文稿工具运行模型（请参阅 [AI 助手](/zh/guide/ai-assistant)）。工具调用在审阅者的浏览器中针对实时演示文稿运行。被接受的建议就是该审阅者的一次普通编辑：它会进入编辑器历史（一次 Ctrl+Z 即可撤销），并像键入一样到达房间。你的审计日志可以记录使用过助手；共享文档不会标记哪些编辑来自助手。

## 运行示例 {#run-the-example}

```sh
bun install                                   # 在仓库根目录
cd demos/pptx/demo-host-review
bun run start                                 # http://localhost:4190
```

在两个浏览器中（或一个普通窗口和一个隐私窗口）打开该页面，并以不同的人登录：Ana 拥有该条目，Ben 和 Chen 审阅，Dee 观看。以 Ana 的身份在侧边面板中更改 Dee 的访问权限，可看到 Dee 的编辑器在只读、编辑和无访问之间切换，而无需重新加载。要试用助手，请在启动前设置 `REVIEW_AI_MODEL`（以及 `REVIEW_AI_API_KEY`，或用于本地 OpenAI 兼容服务器的 `REVIEW_AI_BASE_URL`）。

## 限制 {#limits}

- 示例把所有内容保存在内存中：账户、版本和共享房间会在进程停止时消失。其登录只是一个账户选择器。
- Awareness 内容（显示名称和颜色）来自客户端。中继会将 awareness ID 固定到连接，但不会对照 ticket 的账户检查载荷。
- 评论、审阅批准状态和逐项更改归属不属于此示例中的共享模型；在依赖这些工作流之前，请单独验证。
- 浏览器对等端收敛并不等同于 PowerPoint 无需修复即可重新打开保存的版本；在作出此类承诺之前，请在 PowerPoint 中检查有代表性的演示文稿。
