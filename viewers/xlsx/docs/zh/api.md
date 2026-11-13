# 元素 API

`<xlsx-editor>`（类 `XlsxEditorElement`）是每个软件包挂载的唯一编辑器。使用 `defineXlsxEditor()` 注册它（每个软件包都会导出；注册是幂等的，没有 `customElements` 时什么也不做），或者让 `mountEditor` 和框架组件替你完成。

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const editor = document.createElement('xlsx-editor');
editor.fileName = 'Budget.xlsx';
document.body.append(editor);
await editor.load(bytes, 'Budget.xlsx');
```

## 特性与属性

| 特性               | 属性             | 类型                                                      | 默认值       |
| ------------------ | ---------------- | --------------------------------------------------------- | ------------ |
| `locale`           | `locale`         | `en \| fr \| de \| es \| zh-CN`（任何标签都会映射到其一） | `en`         |
| `read-only`        | `readOnly`       | 布尔值                                                    | `false`      |
| `file-name`        | `fileName`       | 字符串                                                    | `Book1.xlsx` |
| `author-name`      | `authorName`     | 字符串                                                    | `Author`     |
| `theme`            | `theme`          | `light \| dark \| auto`                                   | `auto`       |
| `show-toolbar`     | `showToolbar`    | 布尔值（`"false"` 表示隐藏）                              | `true`       |
| `show-formula-bar` | `showFormulaBar` | 布尔值（`"false"` 表示隐藏）                              | `true`       |

仅属性：

| 属性                 | 类型                               | 说明                                                                                                               |
| -------------------- | ---------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| `themeColors`        | `Partial<Record<...>>`             | 以 `--xve-*` 自定义属性应用的令牌覆盖（[主题](/theming)）。                                                        |
| `validateReplace`    | `boolean`                          | “替换”和“全部替换”会检查数据验证；被拒绝的单元格保持不变（默认 `false`，与 Excel 一致）。                          |
| `hiddenActions`      | `string[]`                         | 按稳定 ID 隐藏的功能区控件、标题栏项目和右键菜单项。                                                               |
| `ribbonAddIns`       | `RibbonAddInTab[]`                 | 宿主在内置选项卡之后添加的选项卡（[绑定](/bindings#ribbon-add-in-tabs)）。                                         |
| `customization`      | `ExcelCustomization \| null`       | 能力（capabilities）、宿主的“文件”选项卡页面和宿主菜单命令（参见[自定义](#customization)）。                       |
| `workbook`           | `Workbook \| null`                 | 获取已打开的工作簿，或设置一个工作簿以替换它。                                                                     |
| `dirty`              | `boolean`（只读）                  | 存在未保存的编辑。                                                                                                 |
| `activeSheet`        | `number`                           | 当前可见工作表的从零开始的索引。                                                                                   |
| `collaboration`      | `XlsxCollaborationOptions \| null` | 设置期间在房间中共享已打开的工作簿（[协作](/zh/collaboration)）。                                                  |
| `collaborationState` | `XlsxCollaborationState`（只读）   | 是否正在共享、连接状态以及房间中的人员。                                                                           |
| `googleFonts`        | `false \| { baseUrl }`             | 对读取器既未安装也未嵌入的字体，回退到网页字体。未设置时，采用读取器“信任中心”的选择（默认关闭）。                 |
| `surfaceOnly`        | 布尔值                             | 挂载前设置，以省略应用界面；与 `surfaceActions` 配合使用。                                                         |
| `surfaceActions`     | `SurfaceHostActions`               | 当设置了 `surfaceOnly` 时，代替对话框的宿主回调。没有回调的工作流会在 `getToolbarProps().enabled` 中报告为未启用。 |
| `loading`            | `boolean`（只读）                  | 工作簿是否正在打开。                                                                                               |

## 自定义 {#customization}

`customization` 接受一个 `ExcelCustomization` 对象，即 `ooxml-core/embedding` 的共享约定，其他编辑器也使用它。每个绑定都将其作为 `customization` prop 或输入转发。对象标识变化时才会应用，因此请保持其稳定；`null` 恢复默认值。

```ts
editor.customization = {
	capabilities: { print: false, export: false, copy: false },
	backstage: { extraPages: [{ id: 'policies', label: 'Policies', render: 'acme-policies' }] },
	contextMenu: {
		extraCommands: [{ id: 'send', label: 'Send range', onSelect: ({ ref }) => send(ref) }],
	},
};
```

- `capabilities`：`false` 会移除通往该能力（`edit`, `open`, `save`, `export`, `print`, `share`, `settings`, `copy`, `comment`, `find`）的所有入口：功能区、“文件”选项卡页面、菜单和快捷键。`hiddenActions` 不变，两者取并集。
- `backstage.extraPages`：宿主在“文件”选项卡中添加的页面 `{ id, label, group?, order?, render }`，位于其分组的内置页面之后。`render` 是挂载函数、自定义元素标签或 `{ element, properties }`，并会收到一个 `ExcelRegionContext`。
- `contextMenu.extraCommands`：宿主命令 `{ id, label, icon?, group?, when?, disabled?, onSelect }`，位于内置项之上（`top`，默认）或之下（`bottom`）。`onSelect` 在菜单关闭后运行，并收到一个 `ExcelMenuContext`。标签按原样显示（请自行本地化）。

`edit: false` 是独立的只读锁，与宿主的只读开关和协作角色 `viewer` 并存：彼此互不解除。能力是界面策略，而不是文档保护。Excel 暂不提供区域替换（选项对话框、窗格）。各产品汇总：[嵌入指南](https://github.com/ChristopherVR/ooxml/blob/main/docs/embedding.md)。

## 方法

| 方法                             | 返回值                    | 说明                                                                                                                                                                             |
| -------------------------------- | ------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `load(bytes, fileName?)`         | `Promise<void>`           | `.xlsx`、`.xlsm`、`.xltx`、`.xls`、`.csv`；格式从字节中嗅探。失败时触发 `workbook-error` 并 reject。                                                                             |
| `newWorkbook()`                  | `void`                    | 一个名为 `Book1.xlsx` 的空白工作簿。                                                                                                                                             |
| `save()`                         | `Promise<Blob>`           | `.xlsx`。不会清除 `dirty`。                                                                                                                                                      |
| `saveBytes(format?)`             | `Promise<Uint8Array>`     | `'xlsx'`（默认）或 `'csv'`（当前工作表，已格式化的值）。                                                                                                                         |
| `download(name?)`                | `Promise<void>`           | 保存、下载，然后 `markClean()`。                                                                                                                                                 |
| `markClean()`                    | `void`                    |                                                                                                                                                                                  |
| `select(ref)`                    | `void`                    | 当前工作表上的 A1 单元格或区域（`'C3'`、`'A1:D20'`）。                                                                                                                           |
| `getSelection()`                 | `string`                  | 以 A1 引用表示的当前选区。                                                                                                                                                       |
| `setActiveSheet(index)`          | `void`                    |                                                                                                                                                                                  |
| `undo()`, `redo()`               | `void`                    | 共享期间，只撤销本窗口的编辑。                                                                                                                                                   |
| `startCollaboration(options)`    | `Promise<void>`           | 加入房间；如果已在共享或没有打开的工作簿，则 reject。                                                                                                                            |
| `stopCollaboration()`            | `void`                    | 离开房间并清除 `collaboration`；工作簿保持原样。                                                                                                                                 |
| `reconnectCollaboration()`       | `void`                    | 重新启动连接，并保留离线时所做的更改。                                                                                                                                           |
| `share()`                        | `void`                    | 打开“共享”对话框。                                                                                                                                                               |
| `focusGrid()`                    | `void`                    | 将键盘焦点移到网格，使输入编辑当前单元格。                                                                                                                                       |
| `openCloudFile(file)`            | `Promise<void>`           | 打开宿主从 `CloudFileProvider` 获得的文件，类似“文件 > 打开 > OneDrive”：当 `file.canEdit` 为 false 时以只读方式打开，“保存”和“另存为”路由到 `file.storage`。加载失败时 reject。 |
| `commitEdit()`                   | `boolean`                 | 在宿主导航或持久化之前提交正在进行的单元格编辑。`false` 表示校验阻止了提交。                                                                                                     |
| `getToolbarProps()`              | `SpreadsheetToolbarProps` | 无界面的工具栏状态以及供自定义控件使用的 `enabled` 标志（[构建模块](/building-blocks)）。                                                                                        |
| `subscribeEditorState(listener)` | `() => void`              | 每次编辑器状态变化时调用 `listener`；返回取消订阅函数。                                                                                                                          |

## 事件

所有事件都是会冒泡并穿越 shadow 边界（`composed`）的 `CustomEvent`。

| 事件                   | `detail`                                                                                              | 说明                                                           |
| ---------------------- | ----------------------------------------------------------------------------------------------------- | -------------------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                                        | 编辑、加载或新建工作簿之后。                                   |
| `workbook-error`       | `{ error, message }`                                                                                  | 加载或保存失败。                                               |
| `workbook-warning`     | `{ warnings }`                                                                                        | 加载器读取到但无法建模的内容。                                 |
| `readonly-change`      | `{ readOnly, locked, hostReadOnly }`                                                                  |                                                                |
| `ribbon-action`        | `{ id }`                                                                                              | 执行了一个功能区命令。                                         |
| `ribbon-customize`     | `{ hiddenActions }`                                                                                   | 用户更改了显示哪些功能区命令（“文件 > 选项 > 自定义功能区”）。 |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' \| 'share' }` | **可取消**：调用 `preventDefault()` 以自行处理。               |
| `selection-change`     | `{ sheet, ref, active }`                                                                              | `ref` 为所选区域，`active` 为活动单元格。                      |
| `sheet-change`         | `{ index, name }`                                                                                     |                                                                |
| `dirty-change`         | `{ dirty }`                                                                                           |                                                                |
| `office-ribbon-add-in` | `{ tab, command }`                                                                                    | 选择了宿主选项卡（`ribbonAddIns`）的一个命令。                 |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                                          | 共享开始或停止、连接变化，或有人加入或离开。                   |
| `validation-error`     | `{ message, title?, ref, style? }`                                                                    | 单元格编辑未通过校验，保持打开状态，直到更正或取消。           |

```ts
editor.addEventListener('file-command', (event) => {
	if (event.detail.command === 'save') {
		event.preventDefault();
		void editor
			.save()
			.then(uploadToMyStorage)
			.then(() => editor.markClean());
	}
});
```

## 宿主自有存储 {#host-owned-storage}

`storage`（元素的属性，也是每个绑定的 prop）接受来自 `ooxml-core/embedding` 的 `DocumentStorage`。设置后：

- “文件 > 保存”“另存为”和 Ctrl+S 会把保存的工作簿交给 `storage.save(bytes, meta)`，而不是下载它。`meta` 为 `{ fileName, mimeType, reason }`，其中 `reason` 为 `'save'` 或 `'saveAs'`；与下载一样，`.xls` 和 `.csv` 源文件会保存为 `.xlsx`。Promise 解决后，工作簿被标记为未修改；如果以 `{ fileName }` 解决，标题栏会采用该名称。拒绝会以 `workbook-error` 报告，工作簿保持已修改状态。
- 当存储提供 `open` 时，“文件 > 打开”会调用 `storage.open()`，并打开其解决得到的 `{ bytes, fileName }`（`undefined` 表示用户已取消）。没有 `open` 时使用文件选择器。
- 同一时间只运行一次保存：保存进行中时再次保存会被忽略，而不是排队。在宿主保存运行期间所做的编辑会使工作簿保持已修改状态，因为已保存的字节不包含该编辑。
- `file-command` 仍会先触发；调用 `preventDefault()` 的宿主自行处理该命令，此时不会调用 `storage`。

“导出”“另存为 CSV”和“打印”仍在浏览器中下载或打印。Excel 没有自动恢复，因此不会使用 `DocumentStorage` 的恢复方法。

```ts
editor.storage = {
	save: (bytes, { fileName }) => myDrive.put(fileName, bytes),
	open: () => myDrive.pick(), // resolves { bytes, fileName } or undefined
};
```

## “文件 > 打开”中的云位置 {#cloud-places-in-file-open}

`fileProviders`（元素的属性，也是每个绑定的 prop）接受一个 `CloudFileProvider`（`ooxml-core/embedding`）数组，例如来自 `ooxml-core/embedding/onedrive` 的 `createOneDriveFileProvider`。为空或省略时，“文件 > 打开”保持不变。有提供程序时，“文件 > 打开”会列出“此设备”（浏览）以及每个提供程序对应的一个位置，该位置显示共享的 `<office-ui-cloud-files>` 浏览器，仅限 `.xlsx`。其文本跟随编辑器的 `locale`。

- 打开文件时会像“打开”一样询问是否放弃未保存的更改，然后通过常规加载路径加载字节，采用该文件的名称并关闭“文件”视图。
- 之后工作簿通过该文件自己的 `storage` 保存，而不是宿主的 `storage`：保存和 Ctrl+S 写入同一项，“另存为”写入副本，只有在提供程序解决后工作簿才被标记为未修改。被拒绝的保存会以 `workbook-error` 报告，工作簿保持已修改状态。打开、新建或加载另一个工作簿会让保存重新使用宿主的 `storage`。
- 提供程序报告为 `canEdit: false` 的文件以只读方式打开（`readOnly` 为 true，并触发 `readonly-change` 事件）。宿主的 `readOnly` 与文件两者中更严格的一方生效，宿主自身的值不会被更改。

```ts
import { createOneDriveFileProvider } from 'ooxml-core/embedding/onedrive';
import { mergeWorkbookSnapshots } from 'ooxml-core/xlsx/collab';

editor.fileProviders = [
	createOneDriveFileProvider({
		getAccessToken,
		signIn,
		mergers: {
			xlsx: ({ base, local, remote }) => mergeWorkbookSnapshots(base.bytes, local, remote.bytes),
		},
	}),
];
```

## 宿主自有的菜单与工作流 {#host-menus-and-workflows}

由宿主掌控的表面（[构建模块](/zh/building-blocks)，或设置了 `surfaceOnly` 和 `surfaceActions` 的 `<xlsx-editor>`）没有自己的右键菜单，也没有设置单元格格式、查找和替换或插入函数对话框。宿主通过下面的回调，或通过 `toolbarProps` 上的操作来绘制它们。行为说明见 [上下文菜单和工作流操作](/zh/building-blocks#context-menus-and-workflow-operations)；本节列出类型和拒绝执行的条件。

回调可以在 `surfaceActions` 上设置，也可以作为选项传给 `useViewerBuildingBlocks`、`createViewerBuildingBlocks` 和 `mountEditorSurface`。

### 上下文菜单回调 {#context-menu-callbacks}

```ts
onContextMenu?: (request: ContextMenuRequest) => boolean | void;
onContextMenuItems?: (request: ContextMenuRequest) => ContextMenuCustomization | void;
```

在单元格、行标题、列标题或工作表标签上右键单击，或按 Shift+F10 时，会运行 `onContextMenu`。返回 `true` 可抑制原生菜单并绘制自己的菜单；返回其他值则显示原生菜单。只有当 `onContextMenu` 没有返回 `true` 时，才会运行 `onContextMenuItems`。它用于过滤和扩展原生菜单，原生菜单仍由表面绘制。

```ts
type ContextMenuKind = 'cell' | 'row' | 'column' | 'tab';

interface SurfaceActionContext {
	sheet: number;
	sheetName: string;
	selection: string; // A1 文本
	activeCell: string; // A1 文本
}

interface ContextMenuItem {
	id: string;
	label: string; // 本地化
	command?: string;
	enabled: boolean;
	checked?: boolean;
	shortcut?: string;
	separatorBefore?: boolean;
}

interface ContextMenuModel extends SurfaceActionContext {
	menu: ContextMenuKind;
	menuSheet: number; // 'tab' 时为被点击的标签，否则为活动工作表
	readOnly: boolean;
	items: ContextMenuItem[];
	execute(id: string): Promise<boolean>;
}

interface ContextMenuRequest extends ContextMenuModel {
	x: number; // 右键单击的客户端坐标
	y: number;
}

interface HostContextMenuItem {
	id: string; // 在同一菜单的宿主项中须唯一；空或重复的 id 会被忽略
	label: string;
	enabled?: boolean | (() => boolean); // 函数会在项目运行时重新求值
	requiresEdit?: boolean; // 只读工作簿中禁用，运行时也会被拒绝
	checked?: boolean;
	separatorBefore?: boolean;
	position?: 'top' | 'bottom'; // 默认 'top'
	onSelect(request: ContextMenuRequest): void | Promise<void>;
}

interface ContextMenuCustomization {
	hide?: readonly string[]; // request.items 中的原生项 id
	add?: readonly HostContextMenuItem[];
}
```

`request.execute(id)` 在项目已运行时解析为 `true`。以下情况解析为 `false`：正在进行的单元格编辑未通过验证（编辑保持打开）；`id` 不在 `items` 中；或该项此时无法运行（只读、工作表保护，或表面中不可用的命令）。工作表标签菜单会先激活其工作表。来自 `onContextMenuItems` 的宿主项在菜单关闭后运行 `onSelect(request)`，除非到那时已被禁用。

### toolbarProps.contextMenu {#toolbarprops-contextmenu}

```ts
contextMenu(menu: ContextMenuKind, sheet?: number): ContextMenuModel;
```

在不发生右键单击的情况下返回同一模型，适用于绘制自己工作表标签或工具栏的宿主。`sheet` 默认为活动工作表，也是 `'tab'` 菜单所对应的标签。表面中不可用的命令会以 `enabled: false` 列出。当安装了 `onOpenFormatCells` 且工作簿可编辑时，“设置单元格格式”为可用。运行内置工作表标签操作（例如重命名）的条目会以禁用状态列出。

### 工作流回调 {#workflow-callbacks}

```ts
onOpenFormatCells?: (request: FormatCellsRequest) => void;
onOpenFindReplace?: (request: FindReplaceRequest) => void;
onOpenInsertFunction?: (request: InsertFunctionRequest) => void;

interface FormatCellsRequest extends SurfaceActionContext, FormatCellsOperations {}
interface FindReplaceRequest extends SurfaceActionContext, FindOperations {}
interface InsertFunctionRequest extends SurfaceActionContext, InsertFunctionOperations {}
```

每个请求包含上述上下文，以及该工作流的操作（见下文）。只有工作流可用时才会运行回调：设置单元格格式和插入函数需要编辑权限，查找和替换则不需要。运行前会先提交正在进行的单元格编辑。`toolbarProps` 上的 `onOpen*` 函数返回 `Promise<boolean>`，当回调缺失、缺少编辑权限或正在进行的编辑未通过验证时为 `false`。

### toolbarProps.workflows {#toolbarprops-workflows}

`toolbarProps.workflows` 提供相同的操作但不需要回调，因此一直打开的面板无需 `onOpen*` 回调。每个操作在运行时都会重新读取只读状态、工作表保护和正在进行的单元格编辑。

#### 设置单元格格式 {#format-cells-operations}

```ts
interface FormatCellsOperations {
	getState(): FormatCellsState;
	canApply(): boolean;
	apply(patch: StylePatch): boolean;
}

interface FormatAspect<T> {
	value: T; // 所有被检查单元格一致的属性
	mixed: string[]; // 存在差异的属性名
}

interface FormatCellsState {
	font: FormatAspect<Font>;
	alignment: FormatAspect<Alignment>;
	border: FormatAspect<Border>;
	protection: FormatAspect<Protection>;
	numFmt: string | null; // 选区混合了多种格式时为 null
	fill: Fill | null; // 选区混合了多种填充时为 null
	complete: boolean; // 选区超过 4096 个单元格时为 false
}
```

- `getState()` 返回选区的格式。当 `complete` 为 `false` 时，混合属性列表是不完整的。
- `canApply()` 在工作簿可编辑且工作表保护允许设置格式时为 true。
- `apply(patch)` 将一个 `StylePatch` 作为单个撤销步骤应用到选区。当 `canApply()` 为 `false` 或正在进行的单元格编辑未通过验证时返回 `false`，且不做任何更改。

#### 查找和替换 {#find-and-replace-operations}

```ts
interface FindQuery {
	text: string;
	sheet?: number;
	range?: CellRange; // 仅与 sheet 一起使用
	matchCase?: boolean;
	wholeCell?: boolean;
	lookIn?: 'formulas' | 'values' | 'comments';
	wildcards?: boolean; // 默认 true
	order?: 'rows' | 'columns';
}

interface FindMatch {
	sheet: number;
	row: number;
	col: number;
	text: string;
}

interface FindOperations {
	canReplace(): boolean;
	findAll(query: FindQuery): FindMatch[];
	findNext(query: FindQuery, options?: { backwards?: boolean }): FindMatch | undefined;
	select(match: FindMatch): boolean;
	replace(query: FindQuery, replacement: string): boolean;
	replaceAll(query: FindQuery, replacement: string): number;
}
```

- `canReplace()` 在工作簿可编辑时为 true。查找从不需要编辑权限。
- `findAll(query)` 返回所有匹配项，在只读工作簿中也可用。当 `text` 为空，或编辑器未连接或正在加载时，返回 `[]`。
- `findNext(query, { backwards })` 选中活动单元格之后的下一个（或上一个）匹配项，并在必要时回绕。没有匹配项或正在进行的单元格编辑未通过验证时返回 `undefined`。
- `select(match)` 选中 `findAll` 返回的匹配项。编辑器未就绪或工作表不存在时返回 `false`。
- `replace(query, replacement)` 替换活动单元格中的匹配项（如果有），并选中下一个匹配项。只有在确实替换了一个单元格时才返回 `true`。工作簿为只读、正在进行的编辑未通过验证，或活动单元格是受保护工作表中的锁定单元格时，返回 `false`。与 Excel 自身的替换一样，它不应用数据验证。
- `replaceAll(query, replacement)` 返回被更改的单元格数量。工作簿为只读或正在进行的编辑未通过验证时返回 `0`。被工作表保护锁定的单元格会被跳过。不应用数据验证。

#### 插入函数 {#insert-function-operations}

```ts
interface FunctionDoc {
	name: string;
	category: string;
	syntax: string; // 例如 SUM(number1, [number2], ...)
	description: string; // 本地化
}

interface InsertFunctionOperations {
	categories(): string[];
	search(options?: { text?: string; category?: 'all' | 'recent' | string }): FunctionDoc[];
	canInsert(): boolean;
	insert(name: string, args?: string): { text: string; caret: number } | undefined;
}
```

- `categories()` 返回去重并排序后的类别名称。
- `search(options)` 先将 `text` 与名称比对，再与说明中的词语比对。`category` 可以是 `'all'`（默认）、`'recent'`（最近使用的函数）或某个类别名称。
- `canInsert()` 在工作簿可编辑时为 true。
- `insert(name, args?)` 在已打开的公式草稿的插入点插入 `NAME(`；若给出 `args`，则插入 `NAME(args)`。若没有打开的草稿，则新建 `=NAME(`。插入位置永远不会在开头的 `=` 之前。它返回插入后草稿的 `text` 和 `caret`，便于输入从那里继续。工作簿为只读、`name` 不是已记录的函数，或打开的草稿是纯文本（不以 `=` 开头）时，返回 `undefined` 且不做任何更改。

### toolbarProps.enabled 标志 {#toolbarprops-enabled}

| 标志                 | 为 true 的条件                                                                         |
| -------------------- | -------------------------------------------------------------------------------------- |
| `contextMenu`        | 已安装 `onContextMenu` 且编辑器已连接。                                                |
| `find`               | 编辑器已连接、已加载并有会话。不需要编辑权限。                                         |
| `replace`            | 工作簿可编辑（`workflows.find.canReplace()`）。                                        |
| `applyFormatCells`   | 现在可以设置格式：工作簿可编辑且工作表保护允许（`workflows.formatCells.canApply()`）。 |
| `insertFunctionCall` | 工作簿可编辑（`workflows.insertFunction.canInsert()`）。                               |
| `formatCells`        | 已安装 `onOpenFormatCells` 且工作簿可编辑。                                            |
| `findReplace`        | 已安装 `onOpenFindReplace`。                                                           |
| `insertFunction`     | 已安装 `onOpenInsertFunction` 且工作簿可编辑。                                         |

每个 `toolbarProps` 快照都会计算这些标志。每个操作在运行时都会重新检查其状态。

### 示例 {#host-menus-and-workflows-example}

```tsx
import { useViewerBuildingBlocks } from '@christophervr/xlsx-react-viewer';

const { toolbarProps } = useViewerBuildingBlocks({
	workbook,
	canEdit: true,
	onContextMenu(request) {
		if (request.menu !== 'cell') return; // 标题和标签保留原生菜单
		showMyMenu(request.x, request.y, request.items, (id) => request.execute(id));
		return true; // 抑制原生菜单
	},
	onOpenFindReplace: ({ findAll, select }) => showFindPanel({ findAll, select }),
});

// 之后，从你自己的面板中调用（在画布挂载之前 toolbarProps 为 null）：
toolbarProps?.workflows.formatCells.apply({ font: { bold: true } });
```

## Shadow part

该元素将其各个区域公开为 CSS shadow part，可用 `xlsx-editor::part(grid)` 设置样式，也可用于测试：`ribbon`、`name-box`、`formula-bar`、`grid`、`sheet-tabs`、`status-bar`、`title-bar`、`backstage` 和 `page-layout`（页面布局视图）。
