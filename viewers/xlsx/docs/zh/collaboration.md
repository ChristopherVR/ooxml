# 协作

多人可以同时编辑同一个工作簿。编辑器通过 `ooxml-core/xlsx/collab` 将工作簿映射为 Yjs 文档，并通过共享的 `ooxml-core/collab` 会话加入房间（Word 和 Visio 编辑器使用的是同一会话和同一批提供方）。每次编辑都会即时发送，对不同单元格的并发编辑会合并，每个人都能看到其他人所在的位置。

## 开始共享

在编辑器中：**文件 > 共享**（或功能区选项卡行右端的“共享”按钮），输入会话名称并选择**开始共享**。同一浏览器中另一个已打开工作簿、并使用相同会话名称开始共享的窗口或标签页会加入该会话。如果页面没有设置服务器，任何内容都不会离开浏览器：会话通过 `BroadcastChannel` 运行。

在代码中，设置 `collaboration` 属性（在每种绑定中也可作为 prop）：

```ts
editor.newWorkbook(); // a window joining an existing room still needs a workbook open
editor.collaboration = {
	roomId: 'budget-2026',
	serverUrl: 'wss://collab.example.com/rooms', // y-websocket compatible; omit for same-browser only
	user: { name: 'Ada', color: '#2563eb' },
};
editor.addEventListener('collaboration-change', (event) => console.log(event.detail));
```

| 选项        | 说明                                                                                                     |
| ----------- | -------------------------------------------------------------------------------------------------------- |
| `roomId`    | 1 到 128 个字母、数字、连字符或下划线。                                                                  |
| `serverUrl` | 与 y-websocket 兼容的服务器的基础 URL；房间 ID 会追加在其后。                                            |
| `provider`  | 你自己的提供方工厂（WebRTC 信令，或通过 `externalProvider` 使用现有的 Yjs 提供方）。优先于 `serverUrl`。 |
| `session`   | 使用 `sanitizePayload: sanitizeXlsxPresence` 创建的现有 `CollabSession`。编辑器绝不会销毁传入的会话。    |
| `user`      | `name`（默认为 `authorName`）、`color`（默认为调色板中的颜色）、`role`（`viewer` 会使编辑器变为只读）。  |

设置了 `collaboration` 期间，之后打开的工作簿会加入同一房间；如果房间中已有工作簿，则房间的内容会替换刚打开的工作簿。将其设置为 `null` 即离开房间。对应的命令式方法有 `startCollaboration(options)`（加入后才 resolve）、`stopCollaboration()`、`reconnectCollaboration()`、`collaborationState` 和 `share()`（打开“文件 > 共享”）。每种绑定都会转发 `collaboration`，报告 `collaboration-change`（`onCollaborationChange`，Vue 中为 `collaboration-change`，Angular 中为 `collaborationChange`，Svelte 中为 `oncollaborationchange`），并在其句柄上公开 `share()`；卸载编辑器会离开房间。

## 共享的内容

- 单元格值、公式（连同其缓存值）、数字格式和单元格格式。
- 工作表：添加、重命名、移动、删除、隐藏和标签颜色。
- 行高、列宽、隐藏的行和列、合并单元格以及已定义名称。
- 每个人的选区，在其他人的网格中以带有对应颜色和姓名的轮廓绘制；会话中的人员以头像显示在标题栏中。

在工作簿处于共享状态时，撤销（Ctrl+Z、功能区和快速访问工具栏）只撤销你自己的编辑；协作者的工作绝不会被还原。连接中断时，编辑器会给出提示并保留你的编辑；连接恢复后这些编辑会被发送，并由第二条通知确认已重新连接。

## 不共享的内容

图表、图片和其他绘图、批注、超链接、条件格式、数据验证、表格和筛选、数据透视表、冻结窗格及其他视图设置、页面设置、保护、文档属性和计算设置都保留在各自的窗口中，并从那里保存。行和列的插入与删除以它们所移动的单元格的形式传递，因此对已移动单元格的并发编辑会落在其旧地址。协作者输入的公式如果使用了同时被重命名的工作表名称，则保留旧名称。冲突规则列在 `ooxml` 仓库中（`docs/collab-area.md`）。

## 宿主负责的内容

编辑器不附带服务器。你的应用负责传输层（服务器 URL 或提供方）、身份验证和授权（客户端中的角色仅供参考）、房间持久化，以及何时、从哪个窗口保存工作簿：编辑器不会协调同时发生的保存。这不是 Excel 的共同创作，也不声称与其对等。
