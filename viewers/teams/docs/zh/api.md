# 元素 API 与事件

每个绑定都封装同一个 `<teams-app>` 自定义元素。绑定会转发这些属性，并以其框架的惯用方式重新发出这些事件（参见[框架指南](/frameworks/react)）。

## 特性与属性

| 属性             | 特性            | 说明                                                                                                                 |
| ---------------- | --------------- | -------------------------------------------------------------------------------------------------------------------- |
| `workspaceId`    | `workspace-id`  | 共享房间：1 到 128 个字符，由字母、数字、`-` 和 `_` 组成。                                                           |
| `userName`       | `user-name`     | 显示名称。                                                                                                           |
| `userId`         | `user-id`       | 稳定的 id；默认使用此浏览器中记住的 id。                                                                             |
| `config`         | `server-config` | `{ mode: 'local' \| 'server', syncUrl?, signalingUrl?, iceServers, iceTransportPolicy?, token? }`。该特性接受 JSON。 |
| `uploadFile`     |                 | 附件的去向；默认为服务器的 `/files` 端点。                                                                           |
| `openers`        |                 | 按 Office 类型区分的处理程序（`docx`、`xlsx`、`pptx`、`vsdx`）。                                                     |
| `embeds`         |                 | 按文件类型区分的回调，返回与沙盒兼容的查看器页面 URL；对 PowerPoint 很有用。                                         |
| `client`（只读） |                 | 该元素背后的核心客户端（状态输入，动作输出）。                                                                       |

省略 `config` 时，该元素会提供一个设置对话框（保存在此浏览器中），或以本地模式运行。

## 内容预览

打开附件时默认显示工作区预览。Word 和 Excel 使用原生编辑器元素，起初为只读；Visio 使用其查看器。Markdown 使用安全子集渲染器；文本文件显示字面文本；HTML 和网站使用沙盒框架。`openers` 和可取消的 `teams-open-file` 事件仍然优先。

PowerPoint 需要由宿主提供嵌入页面：

```ts
const embeds = {
	pptx: ({ url }: { url: string | undefined }) =>
		`https://viewer.example.com/embed?source=${encodeURIComponent(url ?? '')}`,
};
```

将 `embeds` 传给任何组件绑定。该页面必须能理解文件 URL，并支持允许脚本和表单但不含同源权限的沙盒。文件获取不携带凭据；跨源的文件服务器必须允许 CORS。没有配置 PowerPoint 查看器时，窗格会报告这一局限并提供外部打开。

元素的 `previewContent({ attachment, url })` 方法会在本地打开由宿主提供的内容。对于网站，请使用 `attachment: { name: 'Project site', kind: 'other', mime: 'text/html' }`。这不会创建共享的频道标签页。文件视图还提供**预览网站**。拒绝被嵌入框架的网站可以在外部打开。

## 共享的频道标签页与 Excel 副本

使用已上传文件旁边的**固定为标签页**，或使用**添加标签页**来添加一个命名的网站。定义是共享的；每个用户选择自己的标签页。创建者可以通过客户端所有权检查来重命名或移除它。在其他地方移除标签页时，已经打开的本地窗格在关闭之前仍然可用。

配置了文件存储时，Excel 提供本地编辑和**将副本保存到频道**。每次保存都会上传一个名称唯一的副本，并将其发布到原频道，原文件保持不变。保存失败时会保留编辑内容；离开页面时会确认是否放弃未保存的工作，并等待正在进行的保存完成。这不是协同编辑或带版本的回写。强制替换工作区或元素的宿主必须自行处理未保存工作的确认。

原始客户端动作：

```ts
const tab = teams.addTab('Budget', { type: 'file', attachment });
teams.addTab('Project site', { type: 'website', url: 'https://example.com' });
teams.renameTab(tab!.id, 'Budget review');
teams.removeTab(tab!.id);
await teams.saveFileCopy(channelId, file);
```

`TeamsState.tabs` 包含所选频道的定义；`canUploadFiles` 表示已配置存储。`OpenFileDetail.channelId` 记录打开时捕获的频道。当宿主希望启用保存副本到频道时，`previewContent` 也接受此字段。所有权检查需要服务器端授权，才能被视为访问控制。

## 文件创建、上传与搜索

两个文件视图都提供上传、新建空白 Excel 工作簿和搜索。上传和新建工作簿以所选频道为目标，该频道的名称显示在控件旁边。它们需要配置存储；失败后可在原频道重试。同名文件使用各自独立的存储 URL。上传最多接受 20 个文件，每个 32 MiB（参考服务器自身的上限为每个文件 25 MB，除非您调高 `maxFileBytes`）；界面会报告活动状态，但不提供字节级进度或取消。

```ts
await teams.uploadFiles(channelId, files);
const workbook = await teams.createWorkbook(channelId, 'Budget'); // Budget.xlsx
```

创建操作会发布一个由真实 XLSX 文件支持的附件。搜索按文件名、作者、频道和识别出的类型过滤已加载的元数据，不会搜索文件内容。这些动作在异步创建或上传之前就捕获了明确的频道，因此选择其他频道不会改变结果的去向。

## 事件

| 事件                  | 含义                                                               |
| --------------------- | ------------------------------------------------------------------ |
| `teams-ready`         | 元素已就绪。                                                       |
| `teams-open-file`     | 用户打开了附件文件。可取消：调用 `preventDefault()` 即可自行打开。 |
| `teams-config-change` | 用户在设置对话框中保存了新的服务器设置。                           |

各绑定的写法：`onReady`、`onOpenFile`、`onConfigChange`（React、Solid、Svelte、原生）；`@ready`、`@open-file`、`@config-change`（Vue）；`(ready)`、`(openFile)`、`(configChange)`（Angular）。

## 原始客户端

`createTeams(options)`（每个绑定中的原始 hook）返回核心客户端：

```ts
const stop = teams.subscribe(() => render(teams.getState())); // coalesced per tick
teams.send({ text: 'hello' });
teams.on('notice', (notice) => show(notice)); // errors a user should see
```

`getState()` 返回一个不可变的 `TeamsState`：连接 `status`、带未读计数的 `channels`、所选频道的 `messages` 和 `files`、`people`（在线状态）、`typing`、编辑器状态、搜索结果和 `call`。动作包括 `select`、`createChannel`、`send`、`startReply`、`startEdit`、`deleteMessage`、`toggleReaction`、`notifyTyping`、`setAvailability`、`search`、`openCall`、`joinCall`、`leaveCall`、`toggleMic`、`toggleCamera`、`toggleScreenShare`、`toggleHand` 和 `destroy`。各层如何配合，请参阅[架构](/architecture)。
