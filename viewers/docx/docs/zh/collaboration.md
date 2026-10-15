# 协作协议

该 Web Component 支持两种互斥的协作模式：由权威端排序的 ProseMirror 步骤，以及通过应用自有的 `CollabSession` 选择启用的 Yjs 合并。两者使用同一个编辑器和六个框架适配器。宿主负责传输、身份验证、持久化和导出时机。

## 在线演示：两个窗口，任意框架

会话只是一个名称。下面两个窗格是由不同框架适配器（React 和 Vue）构建的独立演示应用，它们加入名为 `docs-collab` 的会话：第一个窗格打开示例文档并托管会话，第二个窗格按名称加入，并接收文档、每一次编辑以及对方的光标。你可以在任一窗格中输入内容。

<iframe
	src="/ooxml/docx/demo/?sample=1&room=docs-collab&name=Ada"
	title="docx-viewer 会话主机（React）"
	loading="lazy"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/docx/demo-vue/?room=docs-collab&name=Grace"
	title="docx-viewer 会话访客（Vue）"
	loading="lazy"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### 任意搭配两个框架

每个演示都接受 `?room=<session>`。同时带有 `sample=1` 的窗口会托管会话；任何使用相同 `room` 的其他窗口都会加入该会话，无论它是用哪个框架构建的。首页的[在线演示](/zh/#live-demo)为每个窗口提供了框架选择器，你也可以在两个标签页中分别打开，例如 `/demo-angular/?sample=1&room=my-room` 和 `/demo-svelte/?room=my-room`。

### 演示如何实现，以及为何仅限同一浏览器

编辑器的协议与传输方式无关，因此由演示提供传输：一个以会话命名的 `BroadcastChannel`。主机窗口运行内存中的参考权威端（`createCollaborationAuthority`），为所有人的步骤批次排序，并广播已接受的批次；访客从主机的起始文档开始并重放已接受的批次，所以各窗口使用哪个框架无关紧要。`BroadcastChannel` 只连接同一个浏览器配置文件的窗口、标签页和框架，别无其他；主机窗口关闭时会话即结束，也不会持久化任何内容。对于使用不同机器的用户，你的应用必须自行提供传输、权威端、身份和存储（见下文）。

另有一个[单页演示](/demo/collaboration.html){target="_self"}，在同一页面中放有两个编辑器，并提供**暂停投递**按钮，便于尝试并发编辑；添加 `?guest=vue` 即可让其第二个编辑器改用其他框架运行。

## 支持与不支持的内容

| 支持                                            | 不支持                                           |
| ----------------------------------------------- | ------------------------------------------------ |
| 通过权威端排序并校验的 ProseMirror 步骤批次     | 托管服务器、网络、身份验证或持久化（由宿主负责） |
| Yjs 文本/格式合并、相对光标、批注线程和本地撤销 | 正文之外的样式、编号定义和注释的并发编辑         |
| 供测试和本地演示使用的内存参考权威端            | 协作期间的表格结构性命令                         |

## Yjs 模式

在每个编辑器中加载相同的源 DOCX 包，包括其媒体、样式、批注、注释和不透明部件。宿主需要为该包提供一个稳定的 `documentId`，并指定且仅指定一个房间创建者。该标识符是宿主约定，不是身份验证令牌，也不是自动生成的包校验和。

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({
	roomId,
	provider: transportProvider({ transport }), // Application-owned transport
	user: { name: 'Ada', role: 'collaborator' },
});
if (!session.synced) {
	await new Promise<void>((resolve) => {
		const off = session.on('synced', (synced) => {
			if (synced) {
				off();
				resolve();
			}
		});
	});
}
editor.startYjsCollaboration(session, {
	documentId: sourcePackageId,
	initializeIfEmpty: designatedCreator,
});
editor.publishPresence({ name: 'Ada', color: '#2563eb' });
```

加入之前，请等待真正的初始提供方同步完成，而不是依赖宽限期写入闸门。创建者为空房间写入初始内容；其他参与者则采用已有的正文以及页面/节属性，而不是用本地快照覆盖它们。文档 ID 或房间格式不匹配会拒绝加入。权威端会话和 Yjs 会话不能同时控制同一个编辑器。

当前共享的 Word 模式会初始化 `word-yjs-v3` 房间。图片、注释引用、域标记或分隔符的每个已建模直接属性，在其导入的属性基础之上都有独立的特性。对不同属性的并发更改会合并，包括属性的移除；本地撤销会保留其他作者的更改。同一属性的冲突遵循 Yjs 的排序。格式历史仍然是每个文本片段只有一个先前快照和一位作者；完整的并发多作者审阅归属尚未完成。

这会改变房间的兼容性：更新后的查看器会拒绝 v1 和 v2 房间，旧客户端则会拒绝 v3 房间。使用旧版模式的核心使用方仍可使用匹配的 v1/v2 房间，但只能撰写文本批注。要升级已持久化的房间，请用匹配的旧客户端导出其当前 DOCX，在更新后的客户端中加载该包，并由指定的创建者初始化一个新的空 v3 房间。请让所有参与者保持使用相匹配的模式。绑定层不会改写正在使用的房间，也不会自动迁移提供方的持久化数据。

Yjs 模式使用稳定的 Yjs 13 绑定 `y-prosemirror` 1.3.7。由于该绑定不会同步文档根属性，因此根属性使用单独的映射。文本、格式和修订标记通过正文片段传输。新插入的 PNG、JPEG、GIF、BMP 和 SVG 图片部件使用共享的核心资源路由，并携带原始二进制负载；名称在每个客户端内是唯一的，部件不可变。停止协作之后，图片字节仍可用于撤销/重做和导出。已有的包资源仍保留在每个编辑器各自匹配的已加载源中。宿主仍然控制导出同步和房间持久化；这并不协调同时发生的保存，也不能证明导出是无损的。

修订记录设置同样通过文档事务传输。切换该设置会更新对等端，参与本地撤销/重做，并导出为 `w:trackRevisions`。加入缺少该设置的兼容房间时，会保留已加载包中的值，直到某位参与者更改它。修订标记会保留编辑作者的姓名以及记录该事务时的 UTC 时间戳。对等端会保留这些值，而不会为远程编辑重新归属；客户端提供的姓名和时钟并不能证明作者身份或时间。

Ctrl/Cmd+Z 和功能区撤销使用本地 Yjs 历史。普通编辑也会把线程记录和锚点放在可撤销的事务中；本地线程快照不包含在 Yjs 文档特性中。停止 Yjs 之后，其当前线程会成为普通的历史快照。历史在编辑器分离和重新挂载后依然保留；断开连接的视图会在重新挂载时采纳更改。在线状态（Awareness）使用相对位置来映射光标/选区，并沿用现有的本地化光标界面。`publishPresence` 会更新 awareness 配置，并在 Yjs 模式下返回 `null`；`leavePresence` 会清除光标。提供方直接传输 awareness，因此 `presence-send` 和 `receivePresence` 仍然是权威端模式的 API。

`reconnectCollaboration()` 会重启 Yjs 提供方，同时保留共享文档。`resyncCollaboration()` 会请求状态交换，并在提供方无法执行时返回 `false`。宿主还可以使用 `session.connect()`、`disconnect()`、`peers()`、状态/同步/错误订阅以及 `destroy()`。`stopCollaboration()` 会分离 Word 绑定，并保留其当前内容和新增的图片资源（包括在卸载期间）；它不会销毁宿主拥有的会话。宿主离开房间时，请销毁该会话。

只读编辑器和查看者角色会接收远程更新，同时阻止本地文档写入。客户端角色仅作参考：请同样在服务器或提供方中强制执行授权。正文之外的编辑和表格结构性命令仍被禁用。页面和节属性会被保留并同步，但多故事（multi-story）协作和完整的包引导尚未完成。

新的 Yjs 房间还会通过现有的审阅窗格共享批注锚点、回复、解决状态和删除操作。创建者负责写入已加载的批注。每条批注都有独立的文本特性、相对于元素的范围以及一条记录；并发的回复会保留各自的 ID。解决状态遵循 Yjs map 的冲突排序。删除根批注会隐藏其回复，包括离线期间并发写入的回复。撤销会同时恢复根批注及其锚点，并显示仍然存在的回复；每位作者的撤销都会保留其他作者的操作。即使编辑器处于分离状态，导出时也会对最新的共享线程做快照。

V3 房间在其 Word 标识映射中带有 `inline-relative-v1` 批注能力。图片、分隔符、注释引用和域标记在并发新增、删除、相邻文本编辑和本地撤销过程中，仍保留各自独立的锚点。创建新房间时，导入的锚点会为这些范围提供初始值。对平衡的复杂域的结果或标记添加批注时，会覆盖整个域，这与记录到的 Word 原生行为一致。匹配的旧版房间保留其 `independent-v1` 纯文本能力。没有该能力的兼容房间会保留已加载的批注，但禁用批注编辑。请通过规范的已保存快照创建新房间来升级；不要在正在使用的房间上更改该能力。权威端步骤模式不会同步批注元数据，因此其批注编辑保持禁用。这些客户端行为不提供服务器授权、持久存储，也不提供现代 Word 批注中的提及、通知和任务分配。

试试 [Yjs 双端演示](/demo/collaboration.html?mode=yjs){target="_self"}。其中的“暂停投递”“重新连接提供方”和“重新同步提供方”操作用于演练内存中的恢复流程。它不提供生产环境的网络，也不提供持久存储。

## 协议

该 Web Component 暴露了 `startCollaboration`、`getPendingCollaboration`、`receiveCollaboration` 和 `stopCollaboration`。宿主可以使用 `createCollaborationAuthority(model, options)` 创建内存中的参考权威端，它与编辑器使用相同的模式和模型转换。

## 客户端与宿主流程

编辑同一份文档的所有参与者使用相同的 `sessionId`；每个编辑器实例使用唯一的 `clientId`。重复使用客户端 ID 可能导致对等端把彼此的步骤误认为确认。所有参与者必须使用相同的模式和起始文档。要加入现有的权威端，请使用其当前版本和规范文档来初始化客户端：

```ts
const client = new CollaborationClient({ sessionId, clientId, version: currentVersion });
const state = EditorState.create({ doc: canonicalDoc, plugins: [history(), client.plugin] });
```

对于自定义元素，请在加载匹配的模型快照之后调用 `editor.startCollaboration({ sessionId, clientId, version })`。参与期间请保持每个接收端处于挂载状态：`receiveCollaboration` 通过存活的编辑器视图来应用。监听 `collaboration-send`；把每个 `StepBatch` 提交给权威端，并用 `receiveCollaboration(batch)` 把其已接受的批次投递给每个编辑器。诸如 `stale`、`out-of-order` 或 `wrong-session` 之类的拒绝，应触发宿主层面的恢复，而不是被悄悄丢弃。

本地事务之后，调用 `client.createPendingBatch(state)`。没有未确认步骤时它返回 `null`，否则返回一个不可变的信封。请持续重试这个完全相同的信封，直到权威端接受；在请求仍在进行时，不要发起另一个请求。在收到与之匹配的已接受批次之前，该方法始终返回同一个信封。

宿主把信封提交给其权威端，并把已接受的 `batch` 转发给每位参与者，包括发送者。已接受批次上的 `version` 是该批次步骤应用之前的权威端版本。参与者把信封传给 `client.receive(state, batch)`。当结果为 `applied` 时，通过编辑器常规的状态更新路径分派返回的事务。这样，普通的模型更新、状态事件和保存行为对远程更改同样会运行。应用远程批次时不会参考视图的只读设置；该设置只阻止本地编辑。

协议每个批次最多接受 500 个步骤和 1,000,000 个序列化 JSON 字符。参考实现 `CollaborationAuthority` 是内存中的排序和变基实现。它接受当前版本的请求，并把过期请求映射到之后已接受的步骤上。其默认历史保留 2,000 个步骤，并缓存 5,000 条请求确认以便重试。早于保留历史的过期请求会被拒绝；未来的版本会因乱序而被拒绝。客户端会记住最近收到的 1,000 个批次 ID，以忽略重复项。出现间隙、版本过期或超大批次被拒绝之后，宿主应提供追赶机制，或从其规范快照重新加载。

## 在线状态

`PresenceClient` 提供一个独立的临时通道，用于传递对等端的姓名、颜色、光标和选区。它不使用也不改变文档步骤流。请把它的 `.plugin` 添加在协作插件旁边。`publish(state, { name, color })` 会返回一条供宿主传输使用的消息；当本地文档步骤仍在等待确认时，返回 `null`。`leave(state)` 返回一条移除消息。收到消息后，当 `receive(state, message)` 的状态为 `applied` 时，请分派其返回的事务；忽略 `duplicate`，并按宿主策略从 `stale`、`out-of-order` 或 `wrong-session` 中恢复。

在线状态消息携带共享的会话 ID、发送者客户端 ID、已提交的协作版本，以及每个发送者单调递增的序号。请按顺序转发消息，并在编辑确认之后重新发布；位置是相对于消息所带的已提交版本定义的。该插件会把现有光标和选区映射到后续的文档事务之上，包括编辑器处于只读状态时收到的编辑。它接受最长 80 个可打印字符的姓名、来自 `PRESENCE_PALETTE` 的颜色，以及最多 100 个可见对等端。宿主应在其传输层对断开连接的对等端进行限流和过期清理。在线状态不会被持久化、认证，也不会包含在 DOCX 导出中。

## 稳定标识

对于共享的自定义元素，在线状态由 `startCollaboration` 启用：

```ts
import { PRESENCE_PALETTE } from 'docx-vanilla-viewer';

editor.addEventListener('presence-send', (event) => {
	transport.sendPresence(event.detail); // Application-owned transport
});
editor.publishPresence({ name: 'Ada', color: PRESENCE_PALETTE[0] });
// Incoming messages: editor.receivePresence(message)
// Explicit departure: editor.leavePresence()
```

元素会在选区变化和编辑确认之后重新发布其存储的配置。当本地编辑仍待处理时，对于收到的选区位置它会返回 `out-of-order`；宿主可以在同步之后请求最新的在线状态。离开消息不包含位置，可以在不同的文档版本移除某个对等端，同时仍会强制执行会话和发送者序号检查。

段落和表格 ID 属于文档内容，并随 ProseMirror 步骤传输。对于本地事务，请使用 `repairCollaborativeDocumentIds(state, clientId, idGenerator)`，而不是非协作的 ID 修复辅助函数。请用 `createCollaborationIdGenerator(clientId)` 只创建一次 `idGenerator`，并在该编辑器的整个生命周期内复用；这样可以防止删除之后 ID 被重复使用，并为本地插入的表格、单元格和段落 ID 提供命名空间。宿主必须为每个存活的编辑器实例分配不同的客户端 ID；会话 ID 和客户端 ID 是协议标识符，而不是凭据。

`stopCollaboration()` 要求待处理的编辑已被确认。如果宿主有意放弃它们，请调用 `stopCollaboration(true)`，显式丢弃本地传输队列。断开连接或卸载不应被视为确认；请保留规范的快照/版本，并通过宿主的传输策略恢复或继续。

## 权威端的职责与限制

应用宿主负责传输、身份验证与授权、持久化存储、快照分发、客户端身份、重新连接策略和导出时机。内存中的权威端适用于测试和本地演示，而不是多进程或生产环境的服务器。如果应用需要共享保存的语义，DOCX 导出应使用规范且已充分同步的模型；本库不协调文件锁或同时进行的导出。远程对等端只会收到模式所支持的 ProseMirror 步骤，权威端会在接受之前校验这些步骤。协作期间表格结构性命令被禁用，因为当前的表格编辑器会整体替换表格结构；在安全地启用这些命令之前，需要细粒度的单元格级事务。

受支持的内联图片、分隔符和注释链接由文本片段属性承载，因此导入或编辑它们并不依赖会被 Yjs 丢弃的元素标记。链接对话框通过与普通文本链接相同的核心命令来读取和编辑这些目标。重新指定目标和移除操作会被同步，并支持本地作者撤销；共享的图片媒体在 DOCX 导出中仍然可用。公式链接和超链接修订归属不在此支持范围内。
