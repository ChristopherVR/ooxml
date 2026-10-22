# 元素 API 与事件

所有框架适配器封装的都是同一个 `<docx-editor>` 自定义元素。适配器会转发这些属性，并以各自框架的习惯方式重新触发这些事件（参见[框架绑定](/bindings)）。

## 特性与属性

| 属性             | 特性              | 说明                                                                                         |
| ---------------- | ----------------- | -------------------------------------------------------------------------------------------- |
| `documentModel`  |                   | 当前显示的 `DocumentModel`。对其赋值属于外部替换，不会触发编辑事件。                         |
| `readOnly`       | `read-only`       | 阻止本地编辑。远程协作批次仍会应用。                                                         |
| `locale`         | `locale`          | 界面语言：`en`、`fr`、`de`、`es` 或 `zh-CN`。文档内容绝不会被翻译。                          |
| `fileName`       | `file-name`       | 当前打开文档的名称（默认为 `Document1.docx`）。                                              |
| `reviewAuthor`   | `review-author`   | 记录在修订和批注上的作者（默认为 `Author`）。                                                |
| `theme`          | `theme`           | `light`、`dark` 或 `auto`。参见[主题](/theming)。                                            |
| `themeColors`    |                   | 以 `--dve-*` 自定义属性形式应用的部分令牌覆盖。                                              |
| `showThumbnails` | `show-thumbnails` | 左侧的页面缩略图栏（仅限打印布局）。                                                         |
| `showToolbar`    | `show-toolbar`    | 设为 `false` 会隐藏功能区；标题栏和状态栏仍保留。                                            |
| `hiddenActions`  |                   | 要按稳定 ID 隐藏的功能区控件；参见[绑定](/bindings)中的表格。                                |
| `ribbonAddIns`   |                   | 宿主在内置选项卡之后添加的选项卡；参见[绑定](/bindings#ribbon-add-in-tabs)。                 |
| `customization`  |                   | 能力（capabilities）、宿主的“文件”选项卡页面和宿主菜单命令（参见[自定义](#customization)）。 |
| `dirty`（只读）  |                   | 编辑后为 true，保存、加载或 `markClean()` 之后为 false。                                     |

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

| 方法                                                  | 说明                                                           |
| ----------------------------------------------------- | -------------------------------------------------------------- |
| `load(bytes)`                                         | 加载 DOCX 或旧版 DOC 字节。失败时拒绝并触发 `document-error`。 |
| `save()`                                              | 解析为已序列化文档的 `Blob`。                                  |
| `saveBytes()`                                         | 解析为原始的 `Uint8Array`。                                    |
| `download(fileName?)`                                 | 保存并启动浏览器下载。                                         |
| `markClean()`                                         | 在宿主持久化文档之后清除脏标记。                               |
| `startCollaboration(config)`                          | 加入协作会话；参见[协作](/zh/collaboration)。                  |
| `startYjsCollaboration(session, options)`             | 加入同步的 Yjs 房间及其匹配的源包。                            |
| `reconnectCollaboration()`, `resyncCollaboration()`   | 恢复 Yjs 连接，或请求提供方交换状态。                          |
| `getPendingCollaboration()`                           | 待发送的步骤批次，或 `null`。                                  |
| `receiveCollaboration(batch)`                         | 应用权威端已接受的批次。                                       |
| `stopCollaboration(discardPending?)`                  | 离开会话；待处理的编辑必须被确认，或被显式丢弃。               |
| `publishPresence`, `receivePresence`, `leavePresence` | 临时在线状态通道；参见[协作](/zh/collaboration#在线状态)。     |

## 事件

所有事件都会冒泡且为 composed，因此可以穿过 Shadow 边界。

| 事件                   | Detail                | 含义                                                                       |
| ---------------------- | --------------------- | -------------------------------------------------------------------------- |
| `document-change`      | `DocumentModel`       | 文档因编辑或审阅操作而发生变化。                                           |
| `document-error`       | `Error`               | 加载、保存、打印或某次编辑失败。                                           |
| `document-warning`     | `string`              | 非致命的提示。                                                             |
| `readonly-change`      | `boolean`             | 用户在界面中切换了编辑/查看模式。                                          |
| `file-command`         | `{ command }`         | 可取消。`command` 为 `new`、`open`、`save`、`export`、`print` 或 `share`。 |
| `dirty-change`         | `boolean`             | 未保存更改的状态发生了翻转。                                               |
| `page-change`          | `{ page, pageCount }` | 打印布局的页码或页数发生变化（分页为近似结果）。                           |
| `ribbon-action`        | `RibbonAction`        | 某个功能区控件被激活。                                                     |
| `ribbon-customize`     | `readonly string[]`   | 用户更改了显示的功能区命令；如有需要，请持久化这些 ID。                    |
| `office-ribbon-add-in` | `{ tab, command }`    | 选择了宿主选项卡（`ribbonAddIns`）中的某个命令。                           |
| `presence-send`        | `PresenceMessage`     | 本地在线状态发生变化；请将其传输给其他客户端。                             |
| `collaboration-send`   | `StepBatch`           | 本地步骤待发送；请将其传输给权威端。                                       |

该包导出了类型化的事件映射（`DocxEditorEventMap`）和运行时列表 `DOCX_EDITOR_EVENTS`。
