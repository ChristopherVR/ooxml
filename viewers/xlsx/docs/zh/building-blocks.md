# 构建模块

每个框架包（`@christophervr/xlsx-react-viewer`、`xlsx-vue-viewer`、`xlsx-angular-viewer`、
`xlsx-svelte-viewer`、`xlsx-solid-viewer` 和 `xlsx-vanilla-viewer`）都导出一个构建模块的 hook 或工厂
（`useViewerBuildingBlocks`、`createViewerBuildingBlocks`），以及 `SpreadsheetCanvas`，并可选择性地导出
`Toolbar`、`FormulaBar` 和 `SheetTabs` 控件。它们全部使用现有的工作簿会话、网格、公式、验证、保护和历史记录；
不存在第二个状态存储。其结构与 PowerPoint 查看器的 `useViewerBuildingBlocks` 保持一致。

```tsx
import {
	SpreadsheetCanvas,
	FormulaBar,
	SheetTabs,
	useViewerBuildingBlocks,
} from '@christophervr/xlsx-react-viewer';

function EmbeddedSheet({ workbook, canEdit, handle }) {
	const { canvasProps, toolbarProps, loading, error } = useViewerBuildingBlocks({
		workbook,
		canEdit,
		handle,
	});
	return (
		<>
			{toolbarProps && <FormulaBar {...toolbarProps} />}
			<SpreadsheetCanvas {...canvasProps} />
			{toolbarProps && <SheetTabs {...toolbarProps} />}
			{error && <p role="alert">{error.message}</p>}
		</>
	);
}
```

该 hook 还会返回 `mode`（`'view'` 或 `'edit'`）以及响应式的 `dirty`、`selection`（A1 文本）、
`collaboration`（房间中的成员和连接状态）和 `editError`（参见[验证失败](#validation-failures)）。加载期间渲染画布；
`toolbarProps` 在挂载之前为 null。请为其容器设置高度。画布只挂载网格。除非宿主提供单独的控件，否则不包含“文件”、
设置、共享、标题栏、功能区、状态栏、编辑栏和工作表标签。

`toolbarProps` 包含选区/活动单元格、工作表状态、公式文本、格式设置以及 `enabled` 状态。可将 `onBold`、`onItalic`、
`onUnderline`、`onFontFamilyChange`、`onFontSizeChange`、`onNumberFormatChange`、`onAlignmentChange`、`onUndo` 和
`onRedo` 连接到自定义控件。对于混合值，格式字段为 null。格式检查最多限定为 4096 个单元格：超过此数量的选区返回 null，
并返回 `formattingComplete=false`。

`enabled` 为每个回调提供一个标志：`bold`、`italic`、`underline`、`fontFamily`、`fontSize`、`alignment`、`numberFormat`、
`undo`、`redo`，以及宿主工作流 `openShareDialog`、`settings`、`formatCells`、`findReplace` 和 `insertFunction`。
`canUndo` 和 `canRedo` 重复了撤销和重做标志。`options` 列出自定义字体、字号、对齐方式或数字格式控件可提供的选项
（`fontFamilies`、`fontSizes`、`alignments`、`numberFormats`），`labels` 则保存本地化的控件名称。

当单元格验证阻止操作时，`onSelect`、`onSheetChange`、`onFormulaChange` 和 `commitEdit` 返回 false。格式化回调返回
`Promise<boolean>`。它们会先提交正在进行的单元格编辑再执行操作，并在自定义控件获得焦点时保留其选区。公式文本由可选的
FormulaBar 在按下 Enter 时提交；被拒绝的输入会设置 aria-invalid。现有句柄保留保存/序列化和 dirty 跟踪。如果输入无法提交，
保存会被拒绝。权限和工作表保护仍具有最终决定权。

对于每个框架，`mountEditorSurface(host, options)` 都提供相同的界面：现有句柄、`getToolbarProps()`、`subscribe(listener)`、
`update` 和 `destroy`。宿主负责文件及可选对话框工作流；无需隐藏的库工具栏。

`formula`、`formulaCaret`、`editing` 和 `editMode` 读取网格的实时编辑桥。`onFormulaInput(text, caret?)` 更新同一份草稿；
`onFormulaChange(text)` 提交它。被拒绝的草稿可以通过 `onFormulaInput` 修正，而无需先提交无效值。`onCancelEdit()` 取消编辑。
不存在单独的公式状态存储。

组合 API 需要同一工作区发布版本中的 `ooxml-ui` 运行时；框架包依赖它，并与它一同发布。
现有的完整查看器导入仍兼容之前发布的运行时。旧版运行时会明确拒绝组合请求，而不会意外挂载应用 UI。

## 对话框的宿主回调 {#host-callbacks-for-dialogs}

应用会为共享、设置、设置单元格格式、查找和替换以及插入函数打开对话框。界面本身从不挂载这些对话框。
与 PowerPoint 的构建模块一样，请传入回调（`onOpenSettings`、`onOpenShareDialog`），并打开你自己的 UI：

```ts
useViewerBuildingBlocks({
	workbook,
	onOpenShareDialog: ({ activeCell }) => openMyShareDialog(activeCell),
	onOpenFormatCells: ({ selection }) => openMyFormatPanel(selection),
});
```

可用的回调包括 `onOpenShareDialog`、`onOpenSettings`、`onOpenFormatCells`、`onOpenFindReplace` 和 `onOpenInsertFunction`。
每个回调都会接收 `{ sheet, sheetName, selection, activeCell }`。`toolbarProps` 提供对应的 `onOpenShareDialog`、
`onOpenSettings`、`onOpenFormatCells`、`onOpenFindReplace` 和 `onOpenInsertFunction`（它们会先提交进行中的编辑，并返回
`Promise<boolean>`）。只有在提供了回调时，`enabled.openShareDialog` 等标志才为 true（格式和插入相关的回调还需要编辑权限）。
`handle.share()` 会调用 `onOpenShareDialog`；如果没有该回调，它什么也不做，且 `enabled.openShareDialog` 为 false。
没有回调的工作流是无法触发的，绝不会被静默挂载。希望使用内置对话框的宿主应改为渲染完整的 `SpreadsheetEditor`。

## 上下文菜单和工作流操作 {#context-menus-and-workflow-operations}

类型、签名和拒绝执行的条件见 [元素 API 参考](/zh/api#host-menus-and-workflows)。

`onOpenFormatCells`、`onOpenFindReplace` 和 `onOpenInsertFunction` 除了上述上下文外，还会收到用于完成工作流的操作。
相同的操作也在 `toolbarProps.workflows` 上，因此始终打开的面板无需回调。每个操作在执行时都会重新读取只读状态、
工作表保护和进行中的单元格编辑，并通过编辑会话执行（撤销、重做、脏状态和 `onWorkbookChange` 与内置界面一致），
拒绝时通过返回值报告。

- `workflows.formatCells`：`getState()` 返回 `{ font, alignment, border, protection }`（各为 `{ value, mixed }`：
  共有属性和不一致的属性名）、`numFmt` 和 `fill`（混合时为 null）以及 `complete`（选区过大时为 false）。
  `canApply()` 和 `apply(patch)` 将样式补丁作为一个撤销步骤应用到选区。
- `workflows.find`：`findAll(query)`、`findNext(query, { backwards })`、`select(match)`、`replace(query, text)`、
  `replaceAll(query, text)` 和 `canReplace()`。查找在只读工作簿中可用；替换在只读时被拒绝，被工作表保护锁定的单元格会被跳过。与 Excel 自带的替换一样，`replace` 和 `replaceAll` 不应用数据验证：即使单元格的规则在键入时会拒绝该值，替换结果也会写入。
- `workflows.insertFunction`：`categories()`、`search({ text, category })`（包含 `name`、`category`、`syntax` 和本地化的
  `description`）、`canInsert()` 和 `insert(name, args?)`。它在当前公式草稿的插入点插入 `NAME(`（没有草稿时新建），
  并返回 `{ text, caret }` 以便您的输入框同步。纯文本草稿会被拒绝。

要绘制自己的上下文菜单，请传入 `onContextMenu(request)`。在单元格、行或列标题以及（挂载了标签栏时）工作表标签上右键时触发。
返回 `true` 可抑制原生菜单，并在 `request.x`、`request.y` 处渲染 `request.items`
（`id`、`label`、`command`、`enabled`、`checked`、`shortcut`、`separatorBefore`）；用 `request.execute(id)` 执行所选项。
返回其他值则显示原生菜单。没有右键事件时，`toolbarProps.contextMenu('cell' | 'row' | 'column' | 'tab', sheet?)`
返回同样的模型，拥有自己工作表标签的宿主可据此获得标签菜单。表面模式只注册少量命令，因此多数项以 `enabled: false` 列出；
当安装了 `onOpenFormatCells` 且工作簿可编辑时，Format Cells 为可用。`enabled.contextMenu`、`enabled.find`、
`enabled.replace`、`enabled.applyFormatCells` 和 `enabled.insertFunctionCall` 反映当前可用的功能。
若要保留原生菜单并修改它，请改为传入 `onContextMenuItems(request)`（仅当 `onContextMenu` 未返回 `true` 时运行），并返回 `{ hide, add }`。`hide` 列出 `request.items` 中要隐藏的原生项 id；`add` 列出宿主项 `{ id, label, enabled, requiresEdit, checked, separatorBefore, position, onSelect(request) }`，放在原生项上方（`position: 'top'`，默认）或下方（`'bottom'`）。`enabled` 可以是函数。原生项保持其禁用状态并在执行时重新检查；宿主项同样会重新检查（其 `enabled`，以及 `requiresEdit` 相对于只读工作簿），因此遗留的打开菜单无法执行过时的选择。它补充了 `customization` 中的 `hiddenActions` 和 `contextMenu.extraCommands`（适用于随附的应用程序）。

## 验证失败 {#validation-failures}

在格式设置、导航或保存之前，必须先提交正在进行的单元格编辑。当数据验证拒绝该编辑时，回调返回 `false`，编辑保持打开，
失败通过三种方式呈现：

- 选项中的 `onValidationError({ message, title?, ref, style? })`（元素会以相同的详情分发 `validation-error` 事件）；
- 响应式的 `editError` 字符串（也就是 `toolbarProps.editError`），草稿被修改、修正或取消后会再次变为 null；
- 可选的 `FormulaBar` 会设置 `aria-invalid`，并在 `role="alert"` 的 span 中显示消息。

### 数据验证提示 {#the-data-validation-alert}

构建模块从不绘制网格的“重试 / 取消”提示框。Excel 的三种提示样式由宿主通过 `onOpenValidationAlert(request)` 显示
（该回调也可用于 `surfaceActions` 以及每个框架适配器）。请求包含 `{ ref, sheet, sheetName, value, style, title, message }`，
其中 `value` 是用户输入的文本，并提供与 Excel 按钮对应的方法：

| `style`       | Excel 按钮     | 方法                                          |
| ------------- | -------------- | --------------------------------------------- |
| `stop`        | 重试 / 取消    | `retry()`、`cancel()`                         |
| `warning`     | 是 / 否 / 取消 | `accept()`（是）、`retry()`（否）、`cancel()` |
| `information` | 确定 / 取消    | `accept()`（确定）、`cancel()`                |

`retry()` 会保留草稿并让编辑器保持打开，`cancel()` 恢复旧值，`accept()` 则无论如何都提交该值，与内置按钮的行为完全一致。
每个请求只能解析一次；之后的调用会被忽略。在解析之前，单元格编辑器会忽略后续的提交，因此务必解析该请求
（抛出异常的回调视同于 `retry()`）。`onValidationError` 和 `editError` 仍会报告失败。

如果没有该回调，则不会挂载任何内容：每种样式都等同于“重试”，编辑保持打开，失败通过 `onValidationError` / `editError` 报告。
只有宿主提供了回调时，`toolbarProps.enabled.validationAlert` 才为 true。完整的 `SpreadsheetEditor` 和 `<xlsx-editor>` 仍保留其内置提示框。

自定义控件还应显示 `editError`，并让用户通过 `onFormulaInput` 修正草稿，或通过 `onCancelEdit` 丢弃草稿。

## 可选的工具栏 {#optional-toolbar}

库中的 `Toolbar` 会渲染属性所支持的每个控件，因此可以替代手工构建的工具栏：撤销、重做、字体、字号、加粗、倾斜、
下划线、左/中/右对齐以及数字格式。按钮会暴露 `aria-pressed`（混合选区时为 `mixed`），下拉框在格式混合时显示空白值，
所有控件都遵循 `enabled`。按钮不会获得焦点，因此在设置格式时网格会保留其选区。工具栏不渲染上述对话框工作流；
请将 `onOpenShareDialog` 及其他 `onOpen*` 回调连接到你自己的控件。`toolbarModel(toolbarProps)`（由 `xlsx-vanilla-viewer`
和 Svelte 运行时导出）会返回相同的控件列表，供自定义渲染器使用。

## 协作 {#collaboration}

实时共同编辑在界面上即可工作：在选项中传入 `collaboration`（房间 ID、服务器 URL、提供方或会话），hook 会返回响应式的
`collaboration` 状态（`active`、`roomId`、`status`、`synced`、`people`）。共享期间，`onUndo` 和 `onRedo` 仅撤销本窗口中的编辑。
离开或卸载会退出房间。标题栏中的在线成员列表属于应用本身，因此宿主应自行渲染 `collaboration.people`。

## 原生 JS {#vanilla}

`xlsx-vanilla-viewer` 导出 `createViewerBuildingBlocks(options)`，它与其他框架基于同一会话：

```ts
import { createViewerBuildingBlocks } from 'xlsx-vanilla-viewer';

const blocks = createViewerBuildingBlocks({ workbook, canEdit: true });
const stops = [
	blocks.mountToolbar(toolbarHost),
	blocks.mountFormulaBar(formulaHost),
	blocks.mountSheetTabs(tabsHost),
	blocks.mountCanvas(canvasHost),
];
blocks.subscribe((state) => render(state.dirty, state.editError, state.toolbarProps));
blocks.update({ workbook, canEdit: false });
blocks.getState().toolbarProps?.onBold();
blocks.destroy();
```

在宿主传入 `canEdit: true` **之前，该界面为只读**，与 PowerPoint 和 Visio 的构建模块一致（完整的 `SpreadsheetEditor` 默认仍可编辑）。
权限在画布挂载之前即已生效，因此 `mode` 从 `'view'` 开始，不会出现可编辑状态的闪烁。`readOnly` 仍优先于 `canEdit: true`。

每个 `mount*` 都会返回其释放函数（disposer）。`getState()` 返回与 hook 相同的快照（`toolbarProps`、`loading`、`error`、`mode`、
`dirty`、`selection`、`collaboration`、`editError`），`handle` 则是常用的命令式句柄。控件会就地更新其元素，因此更新时焦点和草稿都会保留。

## 原生框架组合 {#native-framework-composition}

框架代码随各自的框架包一同发布。所有适配器都拥有一个现有的网格引擎，并暴露相同的 `EditorHandle`；其原生响应式快照提供
`toolbarProps`、`loading`、`error`、`mode`、`dirty`、`selection`、`collaboration` 和 `editError`。可选的原生控件接收一个
`toolbarProps` 属性或输入。自定义控件调用同一快照中的操作。

### Vue {#vue}

在 `setup` 期间调用该组合式函数。传入一个 getter，以跟踪权限、文档、字节和回调的变化。计算引用会在模板中自动解包；
卸载时会释放引擎。

```vue
<script setup lang="ts">
import {
	useViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-vue-viewer';
const props = defineProps<{ workbook: import('xlsx-vue-viewer').Workbook; canEdit: boolean }>();
const { canvasProps, toolbarProps, loading, error, handle } = useViewerBuildingBlocks(() => ({
	workbook: props.workbook,
	canEdit: props.canEdit,
}));
</script>
<template>
	<Toolbar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<FormulaBar v-if="toolbarProps" :toolbar-props="toolbarProps" />
	<SpreadsheetCanvas v-bind="canvasProps" style="height: 600px" />
	<SheetTabs v-if="toolbarProps" :toolbar-props="toolbarProps" />
</template>
```

### Angular {#angular}

为每个宿主提供 `ViewerBuildingBlocks`，然后将该实例注入宿主或其自定义控件。`options` 是一个可写信号；快照是计算信号。
注入器的 `DestroyRef` 负责清理。该服务还公开了 `workbookChange`、`workbookError`、`selectionChange` 和 `dirtyChange`
事件发射器，供宿主订阅，与 options 中的回调并存。在宿主中订阅时请使用 `takeUntilDestroyed()`。

```ts
@Component({
	standalone: true,
	providers: [ViewerBuildingBlocks],
	imports: [SpreadsheetCanvasComponent, ToolbarComponent, FormulaBarComponent, SheetTabsComponent],
	template: `
		@if (blocks.toolbarProps(); as p) {
			<spreadsheet-toolbar [toolbarProps]="p" />
			<spreadsheet-formula-bar [toolbarProps]="p" />
			<button (click)="p.onBold()" [disabled]="!p.canEdit">Custom bold</button>
			<spreadsheet-sheet-tabs [toolbarProps]="p" />
		}
		<spreadsheet-canvas [mount]="blocks.canvasProps.mount" style="display:block;height:600px" />
	`,
})
export class WorkbookHost {
	readonly workbook = input.required<Workbook>();
	readonly canEdit = input(true);
	readonly blocks = inject(ViewerBuildingBlocks);
	constructor() {
		effect(() => this.blocks.options.set({ workbook: this.workbook(), canEdit: this.canEdit() }));
	}
}
```

请从 `xlsx-angular-viewer` 导入组件、服务和 `Workbook`，并从 `@angular/core` 导入 Angular 的 `Component`、`input`、`effect` 和 `inject`。
`SpreadsheetBuildingBlocks` 是同一服务的产品专用名称。其 `handle` 保留加载/保存/选区功能。

### Solid {#solid}

在 Solid 的所有者（owner）内部创建。输入 getter 和输出访问器参与 Solid 的响应式图；清理随该所有者进行。组件接收响应式 getter 作为属性。

```tsx
import { Show } from 'solid-js';
import {
	createViewerBuildingBlocks,
	SpreadsheetCanvas,
	Toolbar,
	FormulaBar,
	SheetTabs,
} from 'xlsx-solid-viewer';
function WorkbookHost(props: { workbook: Workbook; canEdit: boolean }) {
	const blocks = createViewerBuildingBlocks(() => ({
		workbook: props.workbook,
		canEdit: props.canEdit,
	}));
	return (
		<>
			<Show when={blocks.toolbarProps()}>
				{(p) => (
					<>
						<Toolbar toolbarProps={p()} />
						<FormulaBar toolbarProps={p()} />
						<button onClick={() => p().onBold()}>Custom bold</button>
						<SheetTabs toolbarProps={p()} />
					</>
				)}
			</Show>
			<SpreadsheetCanvas {...blocks.canvasProps} />
		</>
	);
}
```

### Svelte 5 {#svelte-5}

在组件初始化期间创建 store。`$blocks` 原生订阅；使用 `$effect` 转发响应式选项。引擎清理随组件销毁进行。
组件源文件及其声明以显式子路径发布。

```svelte
<script lang="ts">
  import { createViewerBuildingBlocks, type Workbook } from 'xlsx-svelte-viewer/runtime';
  import SpreadsheetCanvas from 'xlsx-svelte-viewer/SpreadsheetCanvas.svelte';
  import Toolbar from 'xlsx-svelte-viewer/Toolbar.svelte';
  import FormulaBar from 'xlsx-svelte-viewer/FormulaBar.svelte';
  import SheetTabs from 'xlsx-svelte-viewer/SheetTabs.svelte';
  let { workbook, canEdit = false }: {workbook: Workbook; canEdit?: boolean} = $props();
  const blocks = createViewerBuildingBlocks();
  $effect(() => blocks.update({workbook, canEdit}));
</script>
{#if $blocks.toolbarProps}
  <Toolbar toolbarProps={$blocks.toolbarProps} />
  <FormulaBar toolbarProps={$blocks.toolbarProps} />
  <button onclick={() => $blocks.toolbarProps?.onBold()}>Custom bold</button>
  <SheetTabs toolbarProps={$blocks.toolbarProps} />
{/if}
<SpreadsheetCanvas {...blocks.canvasProps} />
```

`FormulaBar` 与网格共享实时草稿。按下 Enter 提交；被拒绝的输入会设置 `aria-invalid`。“重试”允许修正；Escape 取消。
当验证拒绝草稿时，与之冲突的格式设置、选区、工作表导航和保存仍会被阻止。只读模式会阻止修改，同时保留选区、
工作表导航和序列化。

## 共享外观 {#shared-appearance}

将画布和可选的原生控件放在同一个 `data-office-theme="light|dark|auto"` 祖先元素之下。五种原生框架集成都会安装相同的共享控件样式，
并读取继承的 `--office-*` 令牌。可在该作用域上自定义 `--office-surface`、`--office-background`、`--office-foreground`、
`--office-border`、`--office-accent`、`--office-ring`、`--office-font` 和 `--office-radius`。运行时的更改会保留草稿和焦点。

画布现有的 `theme` 和 `themeColors` 选项仍然只在本地生效；它们不会改写宿主的祖先元素或同级控件。公共祖先上的旧版 `--xve-*`
覆盖仍会同时作用于画布和可选控件。有关作用域和兼容性示例，请参见[主题](/zh/theming)。
