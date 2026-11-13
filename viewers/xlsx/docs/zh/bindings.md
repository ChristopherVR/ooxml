# 框架绑定

所有绑定都通过 `packages/bindings/src/index.ts` 中的同一个 `mountEditor` 函数挂载 `<xlsx-editor>`，因此每个框架的语义完全一致。

## 属性 {#props}

| 属性              | 类型                          | 默认值       | 说明                                                                                                                  |
| ----------------- | ----------------------------- | ------------ | --------------------------------------------------------------------------------------------------------------------- |
| `workbook`        | `Workbook`                    | 无           | 传入**新对象**时显示。编辑器发出的工作簿不会回写到该属性（不会形成循环）。                                            |
| `bytes`           | `Uint8Array \| ArrayBuffer`   | 无           | 文件字节；每次传入新数组时都会重新加载。                                                                              |
| `src`             | `string`                      | 无           | 字符串变化时获取并加载的 URL；文件名默认为 URL 的最后一段。                                                           |
| `fileName`        | `string`                      | `Book1.xlsx` | 仅在父组件修改时转发，因此编辑器中的“文件 > 打开”可以重命名工作簿。                                                   |
| `readOnly`        | `boolean`                     | `false`      |                                                                                                                       |
| `locale`          | `string`                      | `en`         | `en`、`fr`、`de`、`es`、`zh-CN`，或映射到其中之一的标签（如 `de-DE`）。                                               |
| `theme`           | `'light' \| 'dark' \| 'auto'` | `auto`       | `auto` 跟随操作系统。                                                                                                 |
| `authorName`      | `string`                      | `Author`     | 记录在新批注上。                                                                                                      |
| `validateReplace` | `boolean`                     | `false`      | 在“替换”和“全部替换”中按单元格的数据验证检查每个替换，并让被拒绝的单元格保持不变。Excel 不会这样做。                  |
| `showToolbar`     | `boolean`                     | `true`       | 功能区。                                                                                                              |
| `showFormulaBar`  | `boolean`                     | `true`       | 名称框和编辑栏。                                                                                                      |
| `hiddenActions`   | `string[]`                    | `[]`         | 按稳定 id 隐藏的功能区控件、标题栏项和右键菜单项。                                                                    |
| `ribbonAddIns`    | `RibbonAddInTab[]`            | 无           | 宿主在内置选项卡之后添加的选项卡；参见[功能区加载项选项卡](#ribbon-add-in-tabs)。                                     |
| `themeColors`     | `Partial<Record<...>>`        | `{}`         | 令牌覆盖，参见[主题](/zh/theming)。                                                                                   |
| `collaboration`   | `XlsxCollaborationOptions`    | 无           | 设置后会在房间中共享打开的工作簿；传入新对象即重新加入，传入 `null` 即离开。参见[协作](/zh/collaboration)。           |
| `storage`         | `DocumentStorage`             | 无           | 宿主自有的存储：“保存”和“另存为”调用 `storage.save()`，若存在则“文件 > 打开”调用 `storage.open()`。                   |
| `fileProviders`   | `CloudFileProvider[]`         | 无           | “文件 > 打开”中“此设备”旁边显示的云端位置；从其中打开的工作簿会保存回该文件。请保持该数组稳定。                       |
| `customization`   | `ExcelCustomization`          | 无           | `capabilities`、`backstage.extraPages` 和 `contextMenu.extraCommands`；请保持该对象稳定（传入 `null` 可恢复默认值）。 |

## 回调与事件 {#callbacks-and-events}

| 回调（React、Solid、原生 JS） | Vue 事件               | Angular 输出          | Svelte 属性             | 载荷                                     |
| ----------------------------- | ---------------------- | --------------------- | ----------------------- | ---------------------------------------- |
| `onWorkbookChange`            | `workbook-change`      | `workbookChange`      | `onworkbookchange`      | `Workbook`                               |
| `onWorkbookError`             | `workbook-error`       | `workbookError`       | `onworkbookerror`       | `Error`                                  |
| `onSelectionChange`           | `selection-change`     | `selectionChange`     | `onselectionchange`     | `{ sheet, ref, active }`                 |
| `onDirtyChange`               | `dirty-change`         | `dirtyChange`         | `ondirtychange`         | `boolean`                                |
| `onReadOnlyChange`            | `readonly-change`      | `readOnlyChange`      | `onreadonlychange`      | `boolean`                                |
| `onFileCommand`               | `file-command`         | `fileCommand`         | `onfilecommand`         | `CustomEvent<FileCommandDetail>`，可取消 |
| `onWorkbookWarning`           | `workbook-warning`     | `workbookWarning`     | `onworkbookwarning`     | `string[]`                               |
| `onRibbonAction`              | `ribbon-action`        | `ribbonAction`        | `onribbonaction`        | `{ id }`                                 |
| `onRibbonCustomize`           | `ribbon-customize`     | `ribbonCustomize`     | `onribboncustomize`     | `string[]`（隐藏的操作）                 |
| `onSheetChange`               | `sheet-change`         | `sheetChange`         | `onsheetchange`         | `{ index, name }`                        |
| `onRibbonAddIn`               | `office-ribbon-add-in` | `ribbonAddIn`         | `onribbonaddin`         | `{ tab, command }`                       |
| `onCollaborationChange`       | `collaboration-change` | `collaborationChange` | `oncollaborationchange` | `XlsxCollaborationState`                 |
| `onReady`                     | `ready`                | `ready`               | `onready`               | `<xlsx-editor>` 元素                     |

加载失败（文件损坏、加密或不受支持）会触发 `onWorkbookError`；`src` 获取失败也会触发。当新的 `bytes` 或新的 `workbook` 到达，或 `src` 被清空或修改时，仍在进行中的 `src` 获取会被丢弃。

`onFileCommand` 接收的是可取消的 `file-command` 事件本身，而不是其 detail：调用 `event.preventDefault()` 可接管某个文件命令（新建、打开、保存、另存为、导出、导出 CSV、打印、共享），并读取 `event.detail.command` 查看是哪个命令。Vue、Angular 和 Svelte 会以相同的事件对象发出或回调。其他回调直接接收载荷。参见[元素 API](/zh/api)。

每个属性只有在值变化时才会转发（`hiddenActions` 和 `themeColors` 按值比较），因此父组件重新渲染不会撤销用户在编辑器内所做的更改：编辑/查看模式、“文件 > 选项”中的主题或语言，或自定义功能区。受控的父组件可通过 `onReadOnlyChange` 和 `onRibbonCustomize` 保持状态同步（在 Angular 中为 `[(readOnly)]`）。

## 功能区加载项选项卡 {#ribbon-add-in-tabs}

宿主可以像 Office 加载项那样，在 Excel 自带选项卡之后添加自己的功能区选项卡。
`ribbonAddIns` 接收纯数据，是每个绑定都具有的属性；被选中的命令会运行其 `run` 回调，并通过 `onRibbonAddIn`（Vue 为 `office-ribbon-add-in`，Angular 为 `ribbonAddIn`，Svelte 为 `onribbonaddin`）以 `{ tab, command }` 的形式报告。

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

<SpreadsheetEditor ribbonAddIns={tabs} onRibbonAddIn={({ tab, command }) => track(tab, command)} />;
```

`size: 'small'` 的命令按三列排列，`items` 会生成下拉菜单。
框架在每次渲染时都会再次传入一个相等的描述对象，此时面板会保留，只替换回调；如果省略该属性，选项卡会被移除。选项卡不能使用内置选项卡的 id，其命令也不会进入编辑器自身的命令处理。同一描述对象（`RibbonAddInTab`，来自 `ooxml-core/ribbon`，每个框架包都会重新导出）同样适用于 Word、Excel、PowerPoint 和 Visio 编辑器。编辑器本身不附带任何加载项选项卡。

## 句柄 {#handle}

每个适配器都提供相同的句柄：React 通过 `ref`，Vue 通过模板引用，Angular 通过组件实例，Solid 通过 `editorRef`，Svelte 通过组件的导出，原生 JS 则是 `mountEditor` 的返回值。

| 成员                            | 作用                                                                                                            |
| ------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| `element`                       | 已挂载的 `<xlsx-editor>`。                                                                                      |
| `load(bytes, fileName?)`        | 打开 `.xlsx`、`.xlsm`、`.xls` 或 `.csv` 字节；失败时拒绝（reject）。                                            |
| `newWorkbook()`                 | 用空白工作簿替换当前工作簿。                                                                                    |
| `save()`                        | 解析为 `.xlsx` Blob。不会清除 `dirty`。                                                                         |
| `saveBytes(format?)`            | `'xlsx'`（默认）或 `'csv'`（当前工作表）字节。                                                                  |
| `download(fileName?)`           | 保存、在浏览器中下载，并将工作簿标记为已保存。                                                                  |
| `markClean()`                   | 在你自行保存 Blob 后清除 `dirty`。                                                                              |
| `select(ref)`, `getSelection()` | 以 A1 引用（`'B2:D8'`）表示的选区。                                                                             |
| `setActiveSheet(index)`         | 切换工作表。                                                                                                    |
| `share()`                       | 打开“共享”对话框。卸载编辑器会离开任何已共享的房间。                                                            |
| `openCloudFile(file)`           | 从 `CloudFileProvider` 打开文件，与“文件 > 打开 > OneDrive”的行为相同：使用其名称、只读状态以及用于保存的存储。 |
| `dirty`                         | 未保存的编辑。Svelte 的 `isDirty()` 已弃用；请读取 `dirty`。                                                    |

在编辑器挂载之前调用句柄方法会抛出 `Editor is not mounted`。

## 服务端渲染 {#server-rendering}

所有软件包都可以在没有 DOM 的环境中导入。注册（`defineXlsxEditor`）和挂载都是客户端操作。
