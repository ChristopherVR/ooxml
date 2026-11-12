# 构建模块

`visio-react-viewer` 导出 `useViewerBuildingBlocks`、`DiagramCanvas`，以及可选的 `Toolbar`、`PageNavigation`、`Inspector` 和 `ShapesPane`。它们复用 ViewerController、现有的 SVG 渲染器、选择检查器和形状窗口。其他每个框架（包括原生 JavaScript）都有相同的组成部分（见下文）。

```tsx
import { DiagramCanvas, PageNavigation, useViewerBuildingBlocks } from 'visio-react-viewer';

function EmbeddedDiagram({ content, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error, markClean } = useViewerBuildingBlocks({
		content,
		canEdit,
		handle,
		onDirtyChange: (dirty) => console.log(dirty),
		onOpenSettings: () => openMySettingsDialog(),
	});
	return (
		<>
			<PageNavigation {...toolbarProps} />
			<DiagramCanvas {...canvasProps} />
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

`content` 接受与查看器相同的字节/Blob 来源。`handle` 会接收现有的 ViewerController；宿主持久化时，可使用 `load`、`applyEdits`、`exportVsdx` 及其响应式的 `state.edit.dirty`。`onDocumentChange`、`onDirtyChange` 和 `onDocumentError` 均为可选。请为画布容器设置高度，并在加载期间保持其挂载。

扁平的 `toolbarProps` 是自定义界面所需的全部约定；每个字段都取自现有的控制器，因此没有任何内容被重复存储：

| 字段                                                                                                                                | 含义                                                                                                                                                                                                             |
| ----------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `mode`, `canEdit`                                                                                                                   | 只有当宿主授予 `canEdit`**且**打开的图纸是基于源的 `.vsdx` 时，`mode` 才为 `'edit'`（旧版 `.vsd` 或仅有模型的图纸报告为 `'view'`）。只有当宿主允许编辑、图纸基于源且没有加载或编辑正在进行时，`canEdit` 才为真。 |
| `pageIndex`, `pages`, `zoom`, `selection`                                                                                           | 当前页面、页面列表、缩放系数和所选形状。                                                                                                                                                                         |
| `canUndo`, `canRedo`, `canDuplicate`, `canGroup`, `canUngroup`, `canDelete`, `canCut`, `canCopy`, `canPaste`                        | 启用标志。只读时编辑操作均为 false；`canCopy` 需要基于源的图纸以及当前页面上的选择，但不需要编辑权限。                                                                                                           |
| `canZoomIn`, `canZoomOut`                                                                                                           | 缩放限制（10% 至 800%）。                                                                                                                                                                                        |
| `search`                                                                                                                            | 来自控制器搜索的 `{ query, resultCount, activeIndex, truncated }`。在宿主进行导航之前，`activeIndex` 为 -1。                                                                                                     |
| `tool`, `tools`, `canDraw`                                                                                                          | 当前画布工具（`'pointer'` 或 `rectangle`、`ellipse`、`line`、`freeform`、`arc`、`pencil`、`text`、`connector` 之一）、带本地化标签和激活标志的列表，以及是否允许绘制。                                           |
| `labels`                                                                                                                            | 类型化的本地化标签（`DiagramToolbarLabels`），每个控件一项，语言包括 `en`、`de`、`es`、`fr` 和 `zh-CN`。                                                                                                         |
| `onPageChange`, `onZoomChange`, `onZoomIn`, `onZoomOut`, `onFit`                                                                    | 查看操作，始终可用。                                                                                                                                                                                             |
| `onUndo`, `onRedo`, `onDuplicate`, `onGroup`, `onUngroup`, `onDelete`, `onCut`, `onCopy`, `onPaste`, `onToolChange`, `onTextChange` | 编辑操作。它们保留控制器的权限、加载和源格式检查。剪贴板操作使用已挂载的画布和浏览器剪贴板，因此需要用户手势。                                                                                                   |
| `onSearchChange`, `onNextSearchResult`, `onPreviousSearchResult`                                                                    | 搜索，始终可用。                                                                                                                                                                                                 |

`Inspector` 使用 `inspectorProps`，`ShapesPane` 使用 `shapesProps`。画布本身不挂载标题栏、功能区、页面控件、形状/模具窗格、检查器、文件界面、恢复、设置或共享对话框。现有的 VSDX 编辑和保护限制仍然适用。

## 权限与模式 {#permission-and-mode}

`canEdit` 遵循 PowerPoint 构建模块的做法：除非宿主传入 `canEdit: true`，否则图纸为**只读**，并且该权限在控制器创建时、画布挂载之前就已应用，因此不会出现短暂可编辑的情况（完整查看器仍保持其可编辑的默认值）。撤销该权限会取消待处理的编辑、移除绘制工具，并移除调整大小、旋转和端点手柄。页面导航、选择、搜索、滚动/平移、缩放、适应页面、复制和序列化仍然可用。源格式的编辑支持与宿主权限是分开的，这也是 `mode` 同时检查它的原因。

## 画布能做什么 {#what-the-canvas-can-do}

独立画布使用完整查看器的手势（同一套连接方式，而非复制）：移动、行内文本（双击、F2 或在选中的形状上键入）、调整大小和旋转手柄、线条端点、绘制工具，以及通过 `onToolChange` 选择的文本工具和连接符工具。文本工具拖出一个框并打开纯文本草稿（Ctrl+Enter 添加文本框，Escape 取消）；连接符工具从形状拖到形状，或从连接点拖到连接点，并像完整查看器那样以直角路由连接两端。在获得焦点的画布上，Delete、Ctrl+Z、Ctrl+Y 和 Escape（退出绘制工具）均可使用，浏览器的复制、剪切和粘贴事件也会得到处理。它们的每一项都受 `canEdit` 控制，与完整查看器完全一致。编辑被拒绝时，会在画布内的礼貌实时区域中通知。

## 可选的库界面 {#optional-library-ui}

`Toolbar` 会渲染 `toolbarProps` 支持的每个控件：历史记录、缩放和适应页面、复制、组合、取消组合和删除、剪切/复制/粘贴、绘制工具、搜索，以及下文的宿主对话框按钮，并带有启用状态。它是可选的：传入 `include`（`'file' | 'history' | 'zoom' | 'selection' | 'clipboard' | 'tools' | 'search'`）即可只显示部分分组，或者跳过它，并将 `toolbarProps` 连接到你自己的控件上。`PageNavigation` 是页面选择器。

`ShapesPane` 是形状（模具）窗口：包括搜索形状、图纸的文档模具以及内置模具。点击某个母版，或将其拖放到已挂载的画布上，都会经过与完整查看器相同的创建流程，因此服从相同的权限。请将画布挂载到同一个控制器上；若没有画布，母版会添加到页面中央。

## 宿主自有的对话框 {#host-owned-dialogs}

与 PowerPoint 的 `onOpen...` 回调一样，对话框归宿主所有，库在表面模式下从不挂载对话框。将 `onOpenFile`、`onOpenPageSetup`、`onOpenSettings` 和 `onOpenShareDialog` 中的任意一个传给 Hook 或工厂即可。它们会通过 `toolbarProps` 返回，并附带 `canOpenFile`、`canOpenPageSetup`、`canOpenSettings` 和 `canOpenShareDialog`；你省略的回调不会出现，其标志为 `false`（未启用），而库中的 `Toolbar` 只会为你提供的回调显示按钮。

## 持久化与未保存标记 {#persistence-and-dirty-tracking}

`controller.exportVsdx()` 返回字节，并且从不更改 `dirty`。宿主完成持久化后，调用 `markClean()`（可在 Hook 结果、控制器或句柄上调用）。这会清除 `state.edit.dirty`，触发 `onDirtyChange(false)` 以及 `kind` 为 `save` 的 `document-change` 事件。此后的编辑和撤销步骤都以该保存点为基准计算。

## 编辑形状文字时保存 {#saving-while-editing-text}

当用户在形状内输入时按下宿主自有的“保存”，草稿尚未进入文档。请先调用 `await commitTextEdit()`（可在 Hook 结果、会话、vanilla 构建块、Angular 服务或控制器/句柄上调用），再调用 `exportVsdx()`。它只在草稿的编辑完成后才会兑现，即使失焦或 Esc 已经启动了该编辑，并返回 `TextEditOutcome`：

- `committed`：草稿已作为一次可撤销的编辑应用；`text` 为已应用的文字。
- `unchanged`（`reason`：`no-draft` 或 `same-text`）：无需应用。
- `refused`（`reason`：`read-only` 或 `rejected`）：编辑器拒绝；`message` 说明原因，`text` 保留草稿以便恢复。
- `cancelled`（`reason`：`stale-document`、`superseded` 或 `destroyed`）：草稿所属文档已不存在，或该编辑已被取代；`text` 为草稿。

每个结果都带有 `sourceGeneration` 和 `documentGeneration`，有草稿时还有 `pageId` 和 `shapeId`。草稿属于输入它的那个文档：加载其他文件或卸载画布会丢弃草稿，因此它绝不会应用到另一个文档。验证和撤销/重做走正常的编辑路径。

## 软件包 {#packages}

每个框架软件包都会重新导出 `ooxml-ui/visio`，因此任何一个都可以使用 `mountDiagramCanvas(host, controller)`、`mountDiagramInspector(host, controller)`、`mountDiagramShapes(host, controller, { locale? })`、`diagramToolbarProps(controller, fit, locale?, hostCallbacks?)` 和 `fitDiagramCanvas`。请通过现有的控制器订阅，并在卸载时调用返回的清理函数。这里不会引入第二个状态存储。

查看器软件包与 `ooxml-ui` 一同发布，每个查看器软件包都依赖其发布时所对应的 `ooxml-ui` 版本，因此组合 API 始终存在。若宿主手动混用不同版本，较旧的 `ooxml-ui` 会给出明确的错误，而不会意外挂载应用界面。

## 协作 {#collaboration}

构建模块不接受 `collaboration` 选项：在房间中共享图纸是完整 `<visio-viewer>` 的属性（参见[协作](/zh/collaboration)）。自定义界面可以通过 `onOpenShareDialog` 提供自己的共享入口。

## 原生框架组合 {#native-framework-composition}

框架特定的组合位于各个框架软件包中。所有适配器都拥有一个 `ViewerController`，通过框架的响应式机制公开其状态，取消被替换的加载，并随其所有者一同释放订阅和画布处理程序。独立移除的画布可以重新挂载到同一个控制器上，而无需重新加载源。

### Vue {#vue}

```vue
<script setup lang="ts">
import {
	DiagramCanvas,
	Toolbar,
	PageNavigation,
	Inspector,
	useViewerBuildingBlocks,
} from 'visio-vue-viewer';
const props = defineProps<{ content?: Uint8Array; canEdit?: boolean }>();
const blocks = useViewerBuildingBlocks(() => ({ content: props.content, canEdit: props.canEdit }));
const { toolbarProps, loading, error, mode } = blocks;
defineExpose({ controller: blocks.controller });
</script>
<DiagramCanvas v-bind="blocks.canvasProps" />
<Toolbar :toolbar-props="toolbarProps" />
<PageNavigation :toolbar-props="toolbarProps" />
<Inspector v-bind="blocks.inspectorProps" />
```

`toolbarProps`、`state`、`loading`、`error` 和 `mode` 都是 ref；Vue 模板会自动解包它们。传入选项的 getter 或响应式 ref，即可跟踪源、权限和语言的变化。

### Solid {#solid}

```tsx
import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-solid-viewer';
function EmbeddedDiagram(props: { content?: Uint8Array; canEdit?: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		content: props.content,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<DiagramCanvas {...blocks.canvasProps} />
			<Toolbar {...blocks.toolbarProps()} />
		</>
	);
}
```

`toolbarProps`、`state`、`loading`、`error` 和 `mode` 都是访问器。控制器可通过 `blocks.controller` 获取；释放操作遵循 Solid 所有者。`PageNavigation` 接受相同的扁平工具栏描述符，`Inspector` 接受 `inspectorProps`，`ShapesPane` 接受 `shapesProps`。

### Svelte 5 {#svelte-5}

```svelte
<script lang="ts">
 import { writable } from 'svelte/store';
 import { createViewerBuildingBlocks, DiagramCanvas, Toolbar } from 'visio-svelte-viewer';
 const options = writable({ canEdit: true });
 const blocks = createViewerBuildingBlocks(options);
 const toolbarProps = blocks.toolbarProps;
 // options.set({ canEdit: false }) updates permissions through the store subscription.
</script>
<DiagramCanvas {...blocks.canvasProps} />
<Toolbar toolbarProps={$toolbarProps} />
```

选项参数接受可读存储或普通的初始对象；使用普通对象的宿主可以在 rune effect 中调用 `blocks.update(options)`。`toolbarProps`、`state`、`loading`、`error` 和 `mode` 都是可读存储。`PageNavigation` 接受 `toolbarProps`，`Inspector` 接受 `inspectorProps`，`ShapesPane` 接受 `shapesProps`。控制器可通过 `blocks.controller` 获取。

### 原生 JavaScript {#vanilla-javascript}

```ts
import { createViewerBuildingBlocks } from 'visio-vanilla-viewer';
const blocks = createViewerBuildingBlocks({ canEdit: false, onOpenSettings: openSettings });
blocks.mountCanvas(document.querySelector('#canvas')!);
blocks.mountToolbar(document.querySelector('#toolbar')!, ['history', 'zoom']);
blocks.mountPageNavigation(document.querySelector('select')!);
blocks.mountInspector(document.querySelector('#inspector')!);
blocks.mountShapes(document.querySelector('#shapes')!);
blocks.update({ content: bytes, canEdit: true });
blocks.subscribe(() => render(blocks.toolbarProps, blocks.mode));
// later: blocks.markClean(); blocks.destroy();
```

原生工厂运行与其他适配器相同的会话。`toolbarProps`、`state`、`loading`、`error` 和 `mode` 都是 getter，`subscribe(listener)` 会报告每次状态或选项的变化，每个 `mount...` 辅助函数都会返回其清理函数；`destroy()` 会释放仍处于挂载状态的所有内容。已挂载的工具栏会就地更新其元素，因此刷新时焦点得以保留。

### Angular {#angular}

```ts
import { Component, Input, inject, type OnChanges } from '@angular/core';
import {
	ViewerBuildingBlocks,
	DiagramCanvasComponent,
	ToolbarComponent,
} from 'visio-angular-viewer';
@Component({
	selector: 'embedded-diagram',
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [DiagramCanvasComponent, ToolbarComponent],
	template: '<visio-diagram-canvas /><visio-diagram-toolbar />',
})
export class EmbeddedDiagram implements OnChanges {
	@Input() content?: Uint8Array;
	@Input() canEdit = false;
	readonly blocks = inject(ViewerBuildingBlocks);
	ngOnChanges() {
		this.blocks.update({ content: this.content, canEdit: this.canEdit });
	}
}
```

该作用域服务公开只读的 `state`、`toolbarProps`、`loading`、`error` 和 `mode` 信号、现有的控制器，以及 `documentChange`、`dirtyChange` 和 `documentError` 事件发射器。`DiagramCanvasComponent`、`ToolbarComponent`、`PageNavigationComponent`、`InspectorComponent` 和 `ShapesPaneComponent` 都注入该所有者。应将提供者放在每个编辑器宿主上，而不是应用范围内，以隔离多个文档。也可以用 `ViewerBuildingBlocksComponent` 将控件包裹在 `<visio-building-blocks [content]="bytes" [canEdit]="allowed" (dirtyChange)="saveNeeded = $event">` 中，它会提供输入、输出和作用域提供者。

所有可选控件都渲染原生框架模板。宿主可以省略它们，并从自己的界面调用扁平的工具栏操作。工具栏回调保留控制器的权限、加载和源格式检查。浏览器回归测试涵盖了全部六种原生组合，包括编辑、撤销/重做、权限撤销、序列化/重新加载、错误恢复和释放。

对于已经使用 Angular 信号的宿主，`createViewerBuildingBlocks(() => ({
content: bytes(), canEdit: allowed() }))` 会注入作用域的 `ViewerBuildingBlocks` 服务，并通过 Angular effect 连接选项 getter。请在组件的注入上下文中调用它。即使在第一次显式的 `update` 之前，该服务的事件发射器也处于活动状态。

更改独立画布的 `mount` prop 时，会先释放其之前的所有者，再将下一个所有者附加到同一宿主上。更改检查器的控制器也会类似地替换其订阅。这支持在彼此独立的编辑器会话之间切换，而不会留下孤立的输入处理程序或状态订阅。

## 共享的 Office 主题 {#shared-office-theme}

每个原生的可选工具栏、页面选择器和检查器，都遵循与图纸画布相同的 `--office-*` CSS 变量。构建模块会自动安装其共享样式表；不存在特定于框架的主题对象，也不存在第二个文档状态存储。请在你自己的宿主元素上设置主题范围：

```html
<section
	data-office-theme="dark"
	style="--office-accent: #a78bfa; --office-font: Inter, sans-serif"
>
	<!-- DiagramCanvas, Toolbar, PageNavigation and Inspector -->
</section>
```

可在文档根元素或包含它的区域上使用 `data-office-theme="light"`、`"dark"` 或 `"auto"`。`auto` 会跟随系统偏好，包括运行时的变化。带范围的浅色和深色宿主可以同时存在于一个页面上。若没有范围属性，构建模块会继承根主题，并跨越画布的影子边界继承宿主变量。更改主题不会重新加载画布、重置选择，也不会修改文档颜色或序列化的字节。

若需要为宿主定制调色板，请在包含区域上覆盖 `--office-background`、`--office-surface`、`--office-foreground`、`--office-muted-foreground`、`--office-border`、`--office-accent`、`--office-ring` 和 `--office-font`。旧版 Visio 覆盖项 `--vv-surface`、`--vv-ink`、`--vv-border`、`--vv-accent`、`--vv-focus` 和 `--vv-font` 仍受支持，且优先于其规范对应项。画布还会保留 `--vv-background`、`--vv-secondary` 以及现有的状态栏覆盖项。图纸页面保留其自身的文档外观，而周围的控件则跟随界面主题。
