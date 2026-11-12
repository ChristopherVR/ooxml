# 框架绑定

所有绑定都通过同一个 `mountEditor` 函数挂载 `<docx-editor>`。对属性的赋值属于外部替换文档，不会触发编辑事件。`document-change` 直接携带一个 `DocumentModel`，`document-error` 携带一个 `Error`。两个事件都会穿越 Shadow DOM 边界冒泡。

React：

```tsx
import { WordEditor } from 'docx-react-viewer';
<WordEditor documentModel={model} readOnly={false} onDocumentChange={setModel} />;
```

`ref` 暴露 `element`、`load(bytes)`、`save()`、`download()`、`markClean()`、`openCloudFile()` 和 `dirty`。

Vue：

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

Angular：

```ts
import { WordEditorComponent } from 'docx-angular-viewer';
// Add WordEditorComponent to your standalone component imports.
```

```html
<word-editor [documentModel]="model" [readOnly]="false" (documentChange)="model = $event" />
```

Svelte 5：

```svelte
<script>
  import WordEditor from 'docx-svelte-viewer';
  let model = $state(initialDocument);
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

SolidJS：

```tsx
import { createSignal } from 'solid-js';
import { WordEditor } from 'docx-solid-viewer';

function DocumentEditor() {
	const [model, setModel] = createSignal(initialDocument);
	return <WordEditor documentModel={model()} onDocumentChange={setModel} />;
}
```

Solid 使用与共享适配器相同的 `documentModel`、`readOnly`、`locale`、`onDocumentChange` 和 `onDocumentError` 选项。使用 `editorRef` 可以获取已挂载编辑器的 `element`、`load` 和 `save` 方法。组件卸载时，Solid 的 owner 会清理编辑器及其监听器。

自定义元素：

```ts
import { registerDocxEditor } from 'docx-vanilla-viewer';
registerDocxEditor();
const element = document.createElement('docx-editor');
element.documentModel = model;
document.body.append(element);
```

原生 JS 的挂载适配器从 `docx-vanilla-viewer` 导出，返回 update、destroy、load、save 和 element 操作。当父级返回组件发出的模型时，它可以避免反馈重置。外部修改模型时请提供一个新对象；原地修改模型不是受支持的响应式机制。

框架包使用各自原生的生命周期和 ref 接口。任何框架模板中都不包含工具栏、页面渲染器或文档命令的实现。SSR 导入不会注册元素；挂载是客户端操作。当前的浏览器契约只验证已安装的版本，而不是所有历史版本的 peer 依赖。

可以在任意框架的编辑器上设置可选的 `locale` 属性，使功能区和搜索控件本地化（`en`、`fr`、`de`、`es` 或 `zh-CN`，例如 `locale="de"`）。它只影响界面文字，不会翻译文档内容。

## 窗口外观与文件命令 {#window-chrome-and-file-commands}

编辑器在其 Shadow Root 内绘制 Word 风格的外观：标题栏（快速访问的保存、撤销和重做，文件名与保存状态，以及“告诉我你想要做什么”(Tell me what you want to do) 命令搜索）、编辑/查看模式切换、功能区选项卡行右端的 **评论**（Comments）和 **共享**（Share）（与 Word 365 一致），以及 **文件** 选项卡及其后台视图（信息，其中包含兼容性说明；新建、打开、保存、另存为 DOCX、打印），还有状态栏（页数、字数、兼容性说明、Web/打印版式及缩放）。

设置 `element.fileName` 可以为打开的文档命名。文件命令会先以可取消的 `file-command` 事件发出，其 `detail.command` 为 `new`、`open`、`save`、`saveAs`、`export`、`exportText`、`print` 或 `share` 之一。调用 `preventDefault()` 即可由你的应用自行处理（例如保存到你自己的存储）；否则编辑器使用其仅限浏览器的默认行为：文件选择器、下载或打印对话框。`share`（共享按钮）会打开所有 Office 编辑器都使用的共享对话框：可以开始一个会话（会话名称、你的名称以及可选的服务器），或通过邀请链接或会话名称加入会话；然后复制邀请链接、查看会话中的成员并停止共享，这些操作会设置 `collaboration` 属性；取消该事件即可改用你自己的共享方式。选择 **查看**（Viewing）会触发 `readonly-change`。

```ts
editor.fileName = 'Quarterly report.docx';
editor.addEventListener('file-command', (event) => {
	const { command } = (event as CustomEvent<{ command: string }>).detail;
	if (command === 'save') {
		event.preventDefault();
		void editor.save().then(uploadToMyStorage);
	}
});
```

## 保存、未保存状态与界面选项 {#saving-dirty-state-and-ui-options}

`save()` 解析为序列化文档的 `Blob`（`saveBytes()` 返回原始的 `Uint8Array`），`download(fileName?)` 会保存并启动浏览器下载，`dirty` 和 `dirty-change` 用于报告未保存的编辑。在 `document-change` 时 `dirty` 变为 true；在执行文件 > 保存、`download()`、加载文档或调用 `markClean()` 之后变为 false；仅调用 `save()` 不会清除它，因为只有宿主知道该 Blob 是否已被持久化。每个适配器都会在其句柄上暴露 `save`、`download`、`markClean` 和 `dirty`，并提供 `onDirtyChange` / `onPageChange`（或框架对应的事件形式）。

| 属性（特性）                         | 默认值  | 作用                                                                                                                                                                                                                                                    |
| ------------------------------------ | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `showThumbnails` (`show-thumbnails`) | `false` | 左侧的页面缩略图栏。缩略图需要打印版式；在草稿视图中，该栏会说明这一点，而不是列出页面。                                                                                                                                                                |
| `showToolbar` (`show-toolbar`)       | `true`  | `show-toolbar="false"` 会隐藏功能区；标题栏和状态栏保持显示。                                                                                                                                                                                           |
| `hiddenActions`                      | `[]`    | 按稳定的 kebab-case id 隐藏功能区控件（`RIBBON_ACTION_IDS`，例如 `'bold'`、`'insert-table'`、`'track-changes'`、`'paste'`、`'word-count'`）；空的分组和选项卡也会隐藏。颜色按钮与其下拉箭头会一起隐藏。未知 id 会被忽略。这些 id 不会随 `locale` 改变。 |
| `ribbonAddIns`                       | 无      | 宿主在内置选项卡之后添加的选项卡；参见[功能区加载项选项卡](#ribbon-add-in-tabs)。                                                                                                                                                                       |

**已弃用：** 在 id 出现之前，`hiddenActions` 接受的是英文控件标签（`'Bold'`、`'Insert table'`）。这些 `LegacyRibbonLabel` 字符串在一个版本内仍被接受：它们会被映射为 id，并且每次赋值时元素会派发一个 `document-warning`。读取 `hiddenActions` 时始终返回 id。`ribbon-action` 事件从未携带标签（其 detail 以命令类型为键），因此不受影响。

`page-change`（`{ page, pageCount }`）仅在打印版式中触发。分页来自本编辑器自身的排版引擎，是一种近似值，并不等同于 Word 的分页；状态栏和缩略图栏的提示文字会对此加以说明。

## 功能区加载项选项卡 {#ribbon-add-in-tabs}

宿主可以像 Office 加载项一样，在 Word 自带的选项卡之后添加自己的功能区选项卡。`ribbonAddIns` 接受纯数据，是每个绑定都提供的属性；被选中的命令会运行其 `run` 回调，并通过 `onRibbonAddIn`（Vue 中为 `office-ribbon-add-in`，Angular 中为 `ribbonAddIn`，Svelte 中为 `onribbonaddin`）报告，其参数为 `{ tab, command }`。

```tsx
const tabs = [
	{
		id: 'reports',
		label: 'Reports',
		groups: [
			{
				label: 'Export',
				commands: [
					{ id: 'export', label: 'Export', icon: 'save', run: () => exportReport() },
					{ id: 'options', label: 'Options', size: 'small' },
					{ id: 'send', label: 'Send', items: [{ id: 'send-mail', label: 'By mail' }] },
				],
			},
		],
	},
];

<WordEditor ribbonAddIns={tabs} onRibbonAddIn={({ tab, command }) => track(tab, command)} />;
```

`size: 'small'` 的命令会以每列三个的方式排列，`items` 则会生成一个下拉菜单。如果框架在每次渲染时都传入一个相等的描述符（这是常见做法），面板会被保留，只替换回调；省略该属性则会移除这些选项卡。选项卡不能使用内置选项卡的 id，其命令也不会进入编辑器自身的命令处理。同一个描述符（`RibbonAddInTab`，来自 `ooxml-core/ribbon`，由每个框架包重新导出）也适用于 Word、Excel、PowerPoint 和 Visio 编辑器。编辑器本身不附带任何加载项选项卡。

## 绑定属性与事件 {#binding-props-and-events}

每个框架适配器都会转发相同的元素属性并重新发出相同的事件，因此下表中的名称在所有绑定中含义一致。事件名称与元素一致（`file-command`），各框架则使用其自身的写法。

| 元素属性       | React / Solid / 原生 JS | Vue 属性       | Angular 输入   | Svelte 属性    |
| -------------- | ----------------------- | -------------- | -------------- | -------------- |
| `fileName`     | `fileName`              | `fileName`     | `fileName`     | `fileName`     |
| `reviewAuthor` | `reviewAuthor`          | `reviewAuthor` | `reviewAuthor` | `reviewAuthor` |
| `themeColors`  | `themeColors`           | `themeColors`  | `themeColors`  | `themeColors`  |

设置或更改 `fileName`、`reviewAuthor` 和 `themeColors` 时才会转发它们，因此通过“文件 > 打开”所做的重命名，不会被传入旧名称的重新渲染所还原。

| 元素事件           | React / Solid / 原生 JS 选项 | Vue 事件           | Angular 输出      | Svelte 回调         |
| ------------------ | ---------------------------- | ------------------ | ----------------- | ------------------- |
| `file-command`     | `onFileCommand(event)`       | `file-command`     | `fileCommand`     | `onfilecommand`     |
| `document-warning` | `onDocumentWarning(message)` | `document-warning` | `documentWarning` | `ondocumentwarning` |
| `readonly-change`  | `onReadOnlyChange(readOnly)` | `readonly-change`  | `readOnlyChange`  | `onreadonlychange`  |
| `ribbon-action`    | `onRibbonAction(detail)`     | `ribbon-action`    | `ribbonAction`    | `onribbonaction`    |
| `ribbon-customize` | `onRibbonCustomize(ids)`     | `ribbon-customize` | `ribbonCustomize` | `onribboncustomize` |

`file-command` 是唯一可取消的事件，因此其处理函数接收的是事件对象本身，而不是 detail。调用 `event.preventDefault()` 即可接管文件命令（保存、打开、新建等）；`event.detail.command` 表示具体是哪一个命令。Vue、Angular 和 Svelte 的处理函数接收的是同一个 `CustomEvent` 对象。其他事件则直接传递其 detail。
