# 元素 API 与事件

所有框架适配器封装的都是同一个 `<docx-editor>` 自定义元素。适配器会转发这些属性，并以各自框架的习惯方式重新触发这些事件（参见[框架绑定](/bindings)）。

## 特性与属性

| 属性                         | 特性              | 说明                                                                                                                                                   |
| ---------------------------- | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `documentModel`              |                   | 当前显示的 `DocumentModel`。对其赋值属于外部替换，不会触发编辑事件。                                                                                   |
| `readOnly`                   | `read-only`       | 阻止本地编辑。远程协作批次仍会应用。                                                                                                                   |
| `locale`                     | `locale`          | 界面语言：`en`、`fr`、`de`、`es` 或 `zh-CN`。文档内容绝不会被翻译。                                                                                    |
| `fileName`                   | `file-name`       | 当前打开文档的名称（默认为 `Document1.docx`）。                                                                                                        |
| `reviewAuthor`               | `review-author`   | 记录在修订和批注上的作者（默认为 `Author`）。                                                                                                          |
| `theme`                      | `theme`           | `light`、`dark` 或 `auto`。参见[主题](/theming)。                                                                                                      |
| `themeColors`                |                   | 以 `--dve-*` 自定义属性形式应用的部分令牌覆盖。                                                                                                        |
| `showThumbnails`             | `show-thumbnails` | 左侧的页面缩略图栏（仅限打印布局）。                                                                                                                   |
| `readAloud`                  |                   | 朗读选项。`stopOnEdit`（默认 `false`）会在文档被编辑时同时停止朗读；替换、加载或卸载总是会停止朗读。                                                   |
| `showToolbar`                | `show-toolbar`    | 设为 `false` 会隐藏功能区；标题栏和状态栏仍保留。                                                                                                      |
| `hiddenActions`              |                   | 要按稳定 ID 隐藏的功能区控件；参见[绑定](/bindings)中的表格。                                                                                          |
| `ribbonAddIns`               |                   | 宿主在内置选项卡之后添加的选项卡；参见[绑定](/bindings#ribbon-add-in-tabs)。                                                                           |
| `collaboration`              |                   | 声明式协作（`DocxCollaborationOptions` 或 `null`）；参见[协作](/zh/collaboration#collaboration-属性)。                                                 |
| `customization`              |                   | 能力（capabilities）、宿主的“文件”选项卡页面和宿主菜单命令（参见[自定义](#customization)）。                                                           |
| `surfaceOnly`                |                   | 只挂载文档及其编辑界面。须在元素连接之前选择。                                                                                                         |
| `surfaceActions`             |                   | 宿主为表面模式提供的对话框：`onOpenHyperlinkDialog`、`onOpenFindReplace` 和 `onOpenImagePicker`。                                                      |
| `googleFonts`                |                   | 为既未安装也未嵌入的文档字体提供网页字体。未设置时遵循“文件 > 选项 > 信任中心”（默认关闭，直到开启）；`false` 将其关闭；`{ baseUrl }` 设置 CSS2 端点。 |
| `storage`                    |                   | 宿主管理的存储，用于保存、另存为和打开；见下方“由宿主管理的存储”。                                                                                     |
| `fileProviders`              |                   | “文件 > 打开”的云位置；见下方“‘文件 > 打开’中的云位置”。                                                                                               |
| `loading`（只读）            |                   | 文档加载期间为 true。                                                                                                                                  |
| `collaborationState`（只读） |                   | 协作房间：是否处于活动状态、房间 ID、状态、同步状态和参与者。                                                                                          |
| `dirty`（只读）              |                   | 编辑后为 true，保存、加载或 `markClean()` 之后为 false。                                                                                               |

## 自定义 {#customization}

`customization` 接受一个 `WordCustomization` 对象，即 `ooxml-core/embedding` 的共享约定，其他编辑器也使用它。每个绑定都将其作为 `customization` prop 或输入转发。对象标识变化时才会应用，因此请保持其稳定；`null` 恢复默认值。

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [
			{ id: 'send', label: 'Send to chat', onSelect: ({ selectedText }) => send(selectedText) },
		],
	},
};
```

- `capabilities`：`false` 会移除通往该能力（`edit`, `print`, `export`, `save`, `open`, `share`, `copy`, `comment`, `review`, `find`, `settings`）的所有入口：功能区、“文件”选项卡页面、菜单和快捷键。`hiddenActions` 不变，两者取并集。
- `backstage.extraPages`：宿主在“文件”选项卡中添加的页面 `{ id, label, group?, order?, render }`，位于其分组的内置页面之后。`render` 是挂载函数、自定义元素标签或 `{ element, properties }`，并会收到一个 `WordRegionContext`。
- `contextMenu.extraCommands`：宿主命令 `{ id, label, icon?, group?, when?, disabled?, onSelect }`，位于内置项之上（`top`，默认）或之下（`bottom`）。`onSelect` 在菜单关闭后运行，并收到一个 `WordMenuContext`。标签按原样显示（请自行本地化）。

`edit: false` 是独立的只读锁，与宿主的只读开关和协作角色 `viewer` 并存：彼此互不解除。能力是界面策略，而不是文档保护。Word 暂不提供区域替换（选项对话框、窗格）。各产品汇总：[嵌入指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md)。

## 方法

| 方法                                                            | 说明                                                                          |
| --------------------------------------------------------------- | ----------------------------------------------------------------------------- |
| `load(bytes)`                                                   | 加载 DOCX 或旧版 DOC 字节。失败时拒绝并触发 `document-error`。                |
| `save()`                                                        | 解析为已序列化文档的 `Blob`。                                                 |
| `saveBytes()`                                                   | 解析为原始的 `Uint8Array`。                                                   |
| `download(fileName?)`                                           | 保存并启动浏览器下载。                                                        |
| `markClean()`                                                   | 在宿主持久化文档之后清除脏标记。                                              |
| `setLoadedDocument(session)`                                    | 挂载由 `loadDocument` 已解析的 `LoadedDocument`。                             |
| `openCloudFile(file)`                                           | 打开来自 `CloudFileProvider` 的文件，与“文件 > 打开 > OneDrive”相同。         |
| `getToolbarProps()`、`getNavigationProps()`、`getStatusProps()` | 宿主自建工具栏、页面导航和状态控件所需的状态与回调。                          |
| `subscribeEditorState(listener)`                                | 选区和格式变化时调用 `listener`，包括不会使文档变脏的变化。返回取消订阅函数。 |
| `startCollaboration(config)`                                    | 加入协作会话；参见[协作](/zh/collaboration)。                                 |
| `startYjsCollaboration(session, options)`                       | 加入同步的 Yjs 房间及其匹配的源包。                                           |
| `reconnectCollaboration()`, `resyncCollaboration()`             | 恢复 Yjs 连接，或请求提供方交换状态。                                         |
| `getPendingCollaboration()`                                     | 待发送的步骤批次，或 `null`。                                                 |
| `receiveCollaboration(batch)`                                   | 应用权威端已接受的批次。                                                      |
| `stopCollaboration(discardPending?)`                            | 离开会话；待处理的编辑必须被确认，或被显式丢弃。                              |
| `publishPresence`, `receivePresence`, `leavePresence`           | 临时在线状态通道；参见[协作](/zh/collaboration#在线状态)。                    |

## 事件

所有事件都会冒泡且为 composed，因此可以穿过 Shadow 边界。

| 事件                   | Detail                               | 含义                                                                                               |
| ---------------------- | ------------------------------------ | -------------------------------------------------------------------------------------------------- |
| `document-change`      | `DocumentModel`                      | 文档因编辑或审阅操作而发生变化。                                                                   |
| `document-error`       | `Error`                              | 加载、保存、打印或某次编辑失败。                                                                   |
| `document-warning`     | `string`                             | 非致命的提示。                                                                                     |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }` | 用户在界面中切换了编辑/查看模式。                                                                  |
| `file-command`         | `{ command, fileName? }`             | 可取消。`command` 为 `new`、`open`、`save`、`saveAs`、`export`、`exportText`、`print` 或 `share`。 |
| `dirty-change`         | `boolean`                            | 未保存更改的状态发生了翻转。                                                                       |
| `page-change`          | `{ page, pageCount }`                | 打印布局的页码或页数发生变化（分页为近似结果）。                                                   |
| `ribbon-action`        | `RibbonAction`                       | 某个功能区控件被激活。                                                                             |
| `ribbon-customize`     | `readonly string[]`                  | 用户更改了显示的功能区命令；如有需要，请持久化这些 ID。                                            |
| `office-ribbon-add-in` | `{ tab, command }`                   | 选择了宿主选项卡（`ribbonAddIns`）中的某个命令。                                                   |
| `presence-send`        | `PresenceMessage`                    | 本地在线状态发生变化；请将其传输给其他客户端。                                                     |
| `collaboration-send`   | `StepBatch`                          | 本地步骤待发送；请将其传输给权威端。                                                               |
| `collaboration-change` | `DocxCollaborationState`             | 协作状态已变化：`{ active, roomId, status, synced, people }`。                                     |

该包导出了类型化的事件映射（`DocxEditorEventMap`）和运行时列表 `DOCX_EDITOR_EVENTS`。

各框架绑定以其框架的命名方式（例如 `onFileCommand`、`fileCommand` 和 `onfilecommand`）转发 props `fileName`、`reviewAuthor` 和 `themeColors`，以及事件 `file-command`、`document-warning`、`readonly-change`、`ribbon-action` 和 `ribbon-customize`。对照表见[框架绑定](/bindings#binding-props-and-events)。

## Shadow 部件

标题栏（文件名、保存状态、快速访问和“告诉我”）是 shadow 部件 `title-bar`。
自行绘制文档栏的宿主可用 `docx-editor::part(title-bar) { display: none; }` 将其隐藏；功能区和状态栏保持显示。

## “文件 > 打开”中的云位置

`fileProviders`（元素的属性，也是每个绑定的 prop）接受来自 `ooxml-core/embedding` 的
`CloudFileProvider`，例如 `createOneDriveFileProvider`。此时“文件 > 打开”会列出“此设备”以及每个提供方对应的一个位置。从某个位置打开的文件通过该位置自己的存储保存；当提供方表示文件不可编辑时，文件为只读。参见[OneDrive 集成](/onedrive#open-from-onedrive-in-file-open)。

## 由宿主管理的存储

`storage`（元素的属性，也是每个绑定的 prop）接受来自 `ooxml-core/embedding` 的 `DocumentStorage`。设置后：

- “文件 > 保存”、“另存为”和 Ctrl+S 会把保存的字节交给 `storage.save(bytes, meta)`，而不是下载。`meta` 为 `{ fileName, mimeType, reason }`，其中 `reason` 为 `'save'` 或 `'saveAs'`。Promise 完成后文档被标记为已保存；如果它以 `{ fileName }` 完成，标题栏会采用该名称。拒绝会以 `document-error` 报告，文档保持未保存状态。
- 当存储提供 `open` 时，“文件 > 打开”会调用 `storage.open()`，并打开其返回的 `{ bytes, fileName }`（`undefined` 表示用户已取消）。没有 `open` 时使用文件选择器。
- 同一时间只运行一次保存：保存进行中时再次保存会被忽略，而不是排队。宿主保存期间所做的编辑会让文档保持未保存状态，因为已保存的字节不包含它。
- `file-command` 仍然先触发。调用 `preventDefault()` 的宿主自行处理该命令，此时不会调用 `storage`。“另存为”会先发出 `saveAs`，再发出 `save`。
- 移除了保存功能（`customization.capabilities.save: false`）的宿主既不会收到 `save` 事件，也不会调用 `storage.save`；此时“另存为”（`export` 能力）只会下载。

导出、导出为文本和打印仍使用浏览器。Word 没有自动恢复功能，因此不会使用 `DocumentStorage` 的恢复方法。

```ts
editor.storage = {
	async save(bytes, { fileName, mimeType }) {
		await fetch(`/files/${encodeURIComponent(fileName)}`, {
			method: 'PUT',
			headers: { 'content-type': mimeType },
			body: bytes,
		});
	},
};
```

## 朗读 {#read-aloud}

Building Blocks 的钩子（`useViewerBuildingBlocks`）在 `toolbarProps` 上提供五个朗读字段。
它们不是元素的属性或事件。它们使用编辑器自身的实现（浏览器的语音合成；不会发送任何数据）。
[Building Blocks 指南](/zh/building-blocks#read-aloud)说明了它们的行为。

| 字段              | 类型                     | 行为                                                                                                     |
| ----------------- | ------------------------ | -------------------------------------------------------------------------------------------------------- |
| `canReadAloud`    | `boolean`                | 屏幕上有文档且浏览器可以朗读时为 true。只读文档同样计入。                                                |
| `readingAloud`    | `boolean`                | 仅在持有语音的编辑器中为 true。若另一个编辑器正在朗读，则为 false。                                      |
| `readAloudIssue`  | `ReadAloudIssue \| null` | 最近的问题：浏览器没有语音合成时为 `'unsupported'`，否则为 `'empty'` 或 `'error'`。没有问题时为 `null`。 |
| `onReadAloud`     | `() => ReadAloudStart`   | 朗读所选内容；未选择时从插入点朗读到末尾。若已在朗读则重新开始。返回 `'started'`，或未能开始的原因。     |
| `onStopReadAloud` | `() => boolean`          | 停止此编辑器的朗读。返回此前是否在朗读。                                                                 |

这些代码是简单的字符串联合类型。库只提供代码；提示文字由宿主编写。

```ts
type ReadAloudIssue = 'unsupported' | 'empty' | 'error';
type ReadAloudStart = 'started' | ReadAloudIssue;
```

`'error'` 表示语音失败。被取消或中断的朗读不会报告为错误。

一个切换朗读并提示空内容情况的宿主 React 按钮：

```tsx
import { DocumentCanvas, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel }) {
	const { canvasProps, toolbarProps } = useViewerBuildingBlocks({ documentModel });
	return (
		<>
			{toolbarProps && toolbarProps.readAloudIssue !== 'unsupported' && (
				<>
					<button
						disabled={!toolbarProps.canReadAloud}
						aria-pressed={toolbarProps.readingAloud}
						onClick={() =>
							toolbarProps.readingAloud
								? toolbarProps.onStopReadAloud()
								: toolbarProps.onReadAloud()
						}
					>
						{toolbarProps.readingAloud ? 'Stop reading' : 'Read aloud'}
					</button>
					{toolbarProps.readAloudIssue === 'empty' && <span role="status">Nothing to read.</span>}
				</>
			)}
			<DocumentCanvas {...canvasProps} />
		</>
	);
}
```

文档被替换或重新加载时，以及编辑器卸载时，朗读会停止；普通编辑只有在设置 `readAloud: { stopOnEdit: true }` 时才会停止朗读。可选的 `Toolbar` 会使用同样的字段绘制它自己的“Read aloud”开关。
