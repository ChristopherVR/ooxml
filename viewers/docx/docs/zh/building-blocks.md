# 构建模块

`docx-react-viewer` 导出 `useViewerBuildingBlocks`、`DocumentCanvas`，以及可选的 `Toolbar`、`PageNavigation` 和 `StatusBar` 组件。它们使用 WordEditor 现有的编辑器、排版、历史记录和序列化引擎，并沿用 PowerPoint 中 `useViewerBuildingBlocks` 的结构：该 Hook 返回属性，由宿主决定渲染什么。

```tsx
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-react-viewer';

function EmbeddedWord({ documentModel, canEdit, handle }) {
	const { canvasProps, toolbarProps, navigationProps, statusProps, loading, error } =
		useViewerBuildingBlocks({
			documentModel,
			canEdit,
			handle,
			onDirtyChange: (dirty) => console.log(dirty),
			// Host-owned dialogs: omit one and its action reports as not enabled.
			onOpenHyperlinkDialog: (request) => openMyLinkDialog(request),
			onOpenFindReplace: (request) => openMyFindBar(request),
			onOpenImagePicker: (request) => openMyFilePicker(request),
		});
	return (
		<>
			{toolbarProps && (
				<button disabled={!toolbarProps.canEdit} onClick={toolbarProps.onBold}>
					My bold control
				</button>
			)}
			<DocumentCanvas {...canvasProps} />
			{statusProps && <span>{statusProps.wordText}</span>}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

即使在加载期间也要渲染画布。在挂载之前，这些属性均为 null。请为画布容器设置高度。

## Hook 的返回值 {#what-the-hook-returns}

| 字段                       | 内容                                                          |
| -------------------------- | ------------------------------------------------------------- |
| `canvasProps`              | 展开到 `DocumentCanvas` 上。负责挂载并释放编辑器。            |
| `toolbarProps`             | 选区与格式状态，以及编辑回调（见下文）。                      |
| `navigationProps`          | 页码、总页数、视图模式以及页面导航回调。                      |
| `statusProps`              | 字数和缩放，以及对应的回调。                                  |
| `loading`、`error`、`mode` | 加载状态、最近一次加载或文档错误，以及 `'view'` 或 `'edit'`。 |

### toolbarProps {#toolbarprops}

状态：`selection`（`{ from, to }`）、`bold`、`italic`、`underline`（当选区混合了多个值时为 `null`）、`fontFamily`、`fontSize`（混合时为 `null`）、`alignment`（`'left' | 'center' | 'right' | 'justify'` 或 `null`）、`list`（`'bullet' | 'number' | 'other' | null`）、`canEdit`、`canUndo`、`canRedo`，以及已翻译的 `labels`。

回调：`onBold`、`onItalic`、`onUnderline`、`onFontFamilyChange(family)`、`onFontSizeChange(points)`、`onAlignmentChange(alignment)`、`onListChange('bullet' | 'number' | 'remove')`、`onUndo`、`onRedo`。每个回调在操作不被允许时（只读或仍在加载）都会返回 `false`。`list` 状态和 `onListChange` 使用相同的 `'bullet' | 'number'` 取值。

### navigationProps {#navigationprops}

`page`、`pageCount`、`pageText`（“Page 2 of 5”）、`pagesAvailable`、`canGoPrevious`、`canGoNext`、`viewMode`（`'draft' | 'print'`），以及回调 `onGoToPage(n)`、`onPreviousPage`、`onNextPage` 和 `onViewModeChange(mode)`。页面仅存在于打印版式中：在草稿视图下，`pagesAvailable` 为 false，页面相关回调返回 `false`。可通过 `onViewModeChange('print')` 切换。

### statusProps {#statusprops}

`words`、`wordText`（“42 words”）、`zoom`（百分比）、`minZoom`、`maxZoom`，以及回调 `onZoomChange(percent)`、`onZoomIn` 和 `onZoomOut`。滚动时同样会更新 `navigationProps`。

## 宿主对话框 {#host-dialogs}

表面本身不挂载任何对话框。这些工作流属于宿主，命名方式与 PowerPoint 的 `onOpen...` 回调一致。将它们作为选项传入即可；此时 `toolbarProps` 会携带对应的 `open...` 打开器（`openHyperlinkDialog`、`openFindReplace`、`openImagePicker`）、一个 `can...` 标志，以及 `hostWorkflows`（列出已提供了哪些回调）。

| 选项                    | 请求对象                                                                         | 启用条件                     |
| ----------------------- | -------------------------------------------------------------------------------- | ---------------------------- |
| `onOpenHyperlinkDialog` | `{ href, text, apply(href, text?), remove() }`，针对选区中的链接                 | 可编辑                       |
| `onOpenFindReplace`     | `{ search, findNext, findPrevious, replaceOne, replaceAll, status, canReplace }` | 已加载（只读时保留查找功能） |
| `onOpenImagePicker`     | `{ insert(file) }`：由宿主选择文件并将其交回                                     | 可编辑                       |

省略某个回调时，`canInsertLink`、`canFindReplace` 或 `canInsertPicture` 为 false，对应的打开器返回 `false`。请求在使用时会重新检查权限，因此在切换为只读后仍保持打开的对话框无法进行编辑。

## 朗读 {#read-aloud}

`toolbarProps` 带有一个朗读控件，使用编辑器自身的实现（浏览器的语音合成；不会发送任何数据）。
它朗读所选内容；未选择时，从插入点朗读到文档末尾。它在只读文档中同样可用：不会编辑文档，也不会将其标记为已修改。
调用它不会移动选区或焦点，因此您界面中的按钮可以先获得焦点。

| 属性或回调          | 含义                                                                                  |
| ------------------- | ------------------------------------------------------------------------------------- |
| `canReadAloud`      | 屏幕上有文档，且浏览器可以朗读                                                        |
| `readingAloud`      | 此编辑器正在朗读                                                                      |
| `readAloudIssue`    | `'unsupported'`（无语音合成）、`'empty'`（无可朗读内容）、`'error'` 或 `null`         |
| `onReadAloud()`     | 开始朗读（必要时重新开始）；返回 `'started'`、`'unsupported'`、`'empty'` 或 `'error'` |
| `onStopReadAloud()` | 停止此编辑器的朗读；返回此前是否在朗读                                                |

浏览器只有一个语音。在一个编辑器中开始朗读会从另一个编辑器接管，后者的 `readingAloud` 变为 `false`。
编辑器在其文档被编辑或替换、加载文件或卸载时停止朗读。提示文字由您自行编写；库只提供代码。
可选的 `Toolbar` 会绘制一个“Read aloud”开关，并在不支持语音的环境中省略它。字段说明（含类型和示例）见[元素 API](/zh/api#read-aloud)。

## 库组件 {#library-components}

`Toolbar` 绘制撤销和重做、字体与字号、粗体、斜体、下划线、四种对齐方式、项目符号和编号列表，以及你提供的每个宿主工作流对应的按钮。`PageNavigation` 绘制上一页、“Page n of m”、下一页，以及草稿或打印版式的切换。`StatusBar` 绘制字数和缩放下拉框。这些组件都是可选的，由属性驱动；每个控件都有稳定的节点，因此格式更新时键盘焦点得以保留。它们使用共享的 Office 样式。自定义控件无需使用它们。

```tsx
{
	toolbarProps && <Toolbar {...toolbarProps} />;
}
{
	navigationProps && <PageNavigation {...navigationProps} />;
}
{
	statusProps && <StatusBar {...statusProps} />;
}
```

## 只读表面 {#read-only-surfaces}

除非宿主传入 `canEdit: true`，否则该表面**保持只读**，这与 PowerPoint 和 Visio 的构建模块保持一致（完整的 `WordEditor` 默认仍可编辑）。该权限会在画布挂载前生效，因此 `mode` 从 `'view'` 开始，不会出现可编辑状态的闪烁。`readOnly` 的优先级仍高于 `canEdit: true`。在只读状态下，每个编辑回调都返回 `false`，文档无法被修改。读者仍可以让表面获得焦点，使用鼠标或键盘选择文本（方向键、Shift+方向键、Home、End、Ctrl+A），进行复制，并从 `toolbarProps.selection` 读取选区。属于应用本身的快捷键（查找、保存、打印）不会安装在表面上；请将它们接到你自己的控件。

**迁移宿主。** 构建模块过去默认可编辑。依赖这一行为的宿主现在必须传入 `canEdit: true`（在选项对象中，或对于原生适配器，在响应式选项中）；否则表面会以只读方式打开，每个编辑回调都返回 `false`。

## 生命周期 {#lifecycle}

使用 `bytes` 加载文件，或使用句柄的 `load`、`save`、`download`、`markClean` 和 `dirty`。在宿主确认已持久化之前，保存操作不会将文档标记为干净。画布不会挂载标题栏、功能区、后台视图、状态栏或应用程序对话框。

每个框架包还导出 `mountEditorSurface(host, options)`，供原生生命周期适配器使用。其绑定暴露 `getToolbarProps()`、`getNavigationProps()`、`getStatusProps()` 和 `subscribe(listener)`，以及编辑器句柄、`update` 和 `destroy`。表面选项不包括仅适用于功能区的选项（`showToolbar`、`showThumbnails`、`hiddenActions`、`ribbonAddIns`、`onRibbonAddIn`），因此 TypeScript 宿主无法传入一个会被忽略的选项。

组合 API 需要对应的 `ooxml-ui` 运行时，该运行时随这些新增功能一同发布。现有的完整查看器导入仍兼容之前发布的运行时。较旧的运行时会明确拒绝组合请求，而不是意外地挂载应用界面。

## 原生框架组合 {#native-framework-composition}

框架适配器位于各自的框架包中。它们共享引擎生命周期辅助函数，以及可选组件所绘制的控件描述符（`bindings/src/controls.ts`）。每个适配器都暴露 `toolbarProps`、`navigationProps`、`statusProps`、`loading`、`error`、`mode` 和 `canvasProps`，以及编辑器 `handle`（React 将 `handle` ref 作为选项传入，而不是返回它），还有组件 `Toolbar`、`PageNavigation` 和 `StatusBar`（Angular 为 `ToolbarComponent`、`PageNavigationComponent`、`StatusBarComponent`；Svelte 为 `docx-svelte-viewer/Toolbar`、`/PageNavigation`、`/StatusBar`）。它们的状态使用各框架常规的响应式原语，其画布组件则挂载现有的 Word 文档表面。

### Vue {#vue}

```vue
<script setup lang="ts">
import { computed } from 'vue';
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from 'docx-vue-viewer';
const props = defineProps<{ documentModel: import('docx-core').DocumentModel; canEdit: boolean }>();
const { canvasProps, toolbarProps, error, handle } = useViewerBuildingBlocks(
	computed(() => ({ documentModel: props.documentModel, canEdit: props.canEdit })),
);
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<DocumentCanvas v-bind="canvasProps" class="document-surface" />
	<p v-if="error" role="alert">{{ error.message }}</p>
</template>
```

该组合式函数会监听一个 ref、computed ref 或 getter，并在其作用域销毁时释放引擎。`toolbarProps`、`loading`、`error` 和 `mode` 都是只读的 computed ref。`handle` 暴露现有的命令式编辑器 API。

### Angular {#angular}

为每个宿主组件提供一个 `ViewerBuildingBlocks` 服务。在模板中读取的 signal 会参与 Angular 的变更检测，包括无 Zone 的应用。当其 provider 被销毁时，`DestroyRef` 会释放引擎。

```ts
import { Component, signal } from '@angular/core';
import { createDocument } from 'docx-core';
import {
	createViewerBuildingBlocks,
	ViewerBuildingBlocks,
	DocumentCanvasComponent,
	ToolbarComponent,
} from 'docx-angular-viewer';

@Component({
	selector: 'embedded-word',
	standalone: true,
	imports: [DocumentCanvasComponent, ToolbarComponent],
	providers: [ViewerBuildingBlocks],
	template: `
		@if (blocks.toolbarProps(); as toolbar) {
			<word-toolbar [toolbarProps]="toolbar" />
		}
		<word-document-canvas [canvasProps]="blocks.canvasProps" />
	`,
})
export class EmbeddedWord {
	readonly canEdit = signal(true);
	readonly model = createDocument();
	readonly blocks = createViewerBuildingBlocks(() => ({
		documentModel: this.model,
		canEdit: this.canEdit(),
	}));
}
```

该服务暴露只读的 `toolbarProps`、`loading`、`error` 和 `mode` signal，一个稳定的 `handle`，以及带类型的 `documentChange`、`documentError`、`dirtyChange` 和 `pageChange` 事件发射器。可以订阅这些发射器，也可以提供现有的选项回调。对于命令式宿主，请注入该服务并调用 `update(options)`，而不是使用辅助函数。

### Svelte 5 {#svelte-5}

```svelte
<script lang="ts">
  import { writable } from 'svelte/store';
  import { createDocument } from 'docx-core';
  import { createViewerBuildingBlocks } from 'docx-svelte-viewer/runtime';
  import DocumentCanvas from 'docx-svelte-viewer/DocumentCanvas';
  import Toolbar from 'docx-svelte-viewer/Toolbar';
  const options = writable({ documentModel: createDocument(), canEdit: true });
  const blocks = createViewerBuildingBlocks(options);
  const { toolbarProps, error } = blocks;
</script>
{#if $toolbarProps}<Toolbar toolbarProps={$toolbarProps} />{/if}
<DocumentCanvas {...blocks.canvasProps} />
{#if $error}<p role="alert">{$error.message}</p>{/if}
```

在组件初始化期间调用该工厂函数。它接受一个可读的选项 store，或一个初始选项对象；`blocks.update(next)` 也支持由 rune 驱动的 `$effect` 更新。快照是可读的 store。`documentCanvas` 以 action 的形式导出，供自定义宿主标记使用。组件销毁时会取消订阅选项并销毁引擎。

### Solid {#solid}

```tsx
import { createSignal, Show } from 'solid-js';
import { createDocument } from 'docx-core';
import { createViewerBuildingBlocks, DocumentCanvas, Toolbar } from 'docx-solid-viewer';
function EmbeddedWord() {
	const [canEdit, setCanEdit] = createSignal(true);
	const model = createDocument();
	const blocks = createViewerBuildingBlocks(() => ({ documentModel: model, canEdit: canEdit() }));
	return (
		<>
			<Show when={blocks.toolbarProps()}>{(toolbar) => <Toolbar toolbarProps={toolbar()} />}</Show>
			<DocumentCanvas {...blocks.canvasProps} />
		</>
	);
}
```

该工厂函数在 effect 中跟踪其选项访问器，并为 `toolbarProps`、`loading`、`error` 和 `mode` 返回 signal 访问器。其 owner 通过 `onCleanup` 控制清理。可选的工具栏控件在格式状态变化时保留其 DOM 节点，因此键盘焦点得以保留。所有适配器都以响应式方式接受 `bytes`，在新模型或新加载到达时清除失败的加载，并在画布替换时保持其句柄稳定。

### 原生 JavaScript {#vanilla}

`docx-vanilla-viewer` 将共享会话导出为 `createViewerBuildingBlocks(options)`（`canvasProps`、`handle`、`state`、`subscribe`、`update`、`destroy`），以及返回销毁函数的挂载辅助函数：`mountDocumentCanvas(host, blocks)`、`mountToolbar(host, blocks)`、`mountPageNavigation(host, blocks)` 和 `mountStatusBar(host, blocks)`。

```ts
import {
	createViewerBuildingBlocks,
	mountDocumentCanvas,
	mountToolbar,
	mountStatusBar,
} from 'docx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ documentModel, canEdit: true });
const stops = [
	mountDocumentCanvas(document.querySelector('#page')!, blocks),
	mountToolbar(document.querySelector('#tools')!, blocks),
	mountStatusBar(document.querySelector('#status')!, blocks),
];
blocks.subscribe((state) => console.log(state.toolbarProps?.bold)); // your own controls
// later
stops.forEach((stop) => stop());
blocks.destroy();
```

## 共享样式 {#shared-styling}

可选的原生工具栏组件使用共享的 `ooxml-ui` 样式，包括间距、背景、边框、按下与禁用状态以及键盘焦点环。请将工具栏和画布置于同一个 `data-office-theme` 祖先之下，并在该祖先上自定义 `--office-*` 变量。更改会立即应用于所有框架。现有的祖先级 `--dve-*` 覆盖仍然受支持。宿主创建的控件保留其自身样式，除非宿主选用 `ooxml-ui/theme` 中共享的 `prepareOfficeUiBlock(element, 'toolbar', 'docx')` 辅助函数。
