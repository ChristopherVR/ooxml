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

| 属性                 | 类型                               | 说明                                                                       |
| -------------------- | ---------------------------------- | -------------------------------------------------------------------------- |
| `themeColors`        | `Partial<Record<...>>`             | 以 `--xve-*` 自定义属性应用的令牌覆盖（[主题](/theming)）。                |
| `hiddenActions`      | `string[]`                         | 按稳定 ID 隐藏的功能区控件。                                               |
| `ribbonAddIns`       | `RibbonAddInTab[]`                 | 宿主在内置选项卡之后添加的选项卡（[绑定](/bindings#ribbon-add-in-tabs)）。 |
| `workbook`           | `Workbook \| null`                 | 获取已打开的工作簿，或设置一个工作簿以替换它。                             |
| `dirty`              | `boolean`（只读）                  | 存在未保存的编辑。                                                         |
| `activeSheet`        | `number`                           | 当前可见工作表的从零开始的索引。                                           |
| `collaboration`      | `XlsxCollaborationOptions \| null` | 设置期间在房间中共享已打开的工作簿（[协作](/zh/collaboration)）。          |
| `collaborationState` | `XlsxCollaborationState`（只读）   | 是否正在共享、连接状态以及房间中的人员。                                   |

## 方法

| 方法                          | 返回值                | 说明                                                                                                 |
| ----------------------------- | --------------------- | ---------------------------------------------------------------------------------------------------- |
| `load(bytes, fileName?)`      | `Promise<void>`       | `.xlsx`、`.xlsm`、`.xltx`、`.xls`、`.csv`；格式从字节中嗅探。失败时触发 `workbook-error` 并 reject。 |
| `newWorkbook()`               | `void`                | 一个名为 `Book1.xlsx` 的空白工作簿。                                                                 |
| `save()`                      | `Promise<Blob>`       | `.xlsx`。不会清除 `dirty`。                                                                          |
| `saveBytes(format?)`          | `Promise<Uint8Array>` | `'xlsx'`（默认）或 `'csv'`（当前工作表，已格式化的值）。                                             |
| `download(name?)`             | `Promise<void>`       | 保存、下载，然后 `markClean()`。                                                                     |
| `markClean()`                 | `void`                |                                                                                                      |
| `select(ref)`                 | `void`                | 当前工作表上的 A1 单元格或区域（`'C3'`、`'A1:D20'`）。                                               |
| `getSelection()`              | `string`              | 以 A1 引用表示的当前选区。                                                                           |
| `setActiveSheet(index)`       | `void`                |                                                                                                      |
| `undo()`, `redo()`            | `void`                | 共享期间，只撤销本窗口的编辑。                                                                       |
| `startCollaboration(options)` | `Promise<void>`       | 加入房间；如果已在共享或没有打开的工作簿，则 reject。                                                |
| `stopCollaboration()`         | `void`                | 离开房间并清除 `collaboration`；工作簿保持原样。                                                     |
| `reconnectCollaboration()`    | `void`                | 重新启动连接，并保留离线时所做的更改。                                                               |
| `share()`                     | `void`                | 打开“文件 > 共享”。                                                                                  |
| `focusGrid()`                 | `void`                | 将键盘焦点移到网格，使输入编辑当前单元格。                                                           |

## 事件

所有事件都是会冒泡并穿越 shadow 边界（`composed`）的 `CustomEvent`。

| 事件                   | `detail`                                                                                   | 说明                                                 |
| ---------------------- | ------------------------------------------------------------------------------------------ | ---------------------------------------------------- |
| `workbook-change`      | `{ workbook }`                                                                             | 编辑、加载或新建工作簿之后。                         |
| `workbook-error`       | `{ error, message }`                                                                       | 加载或保存失败。                                     |
| `workbook-warning`     | `{ warnings }`                                                                             | 加载器读取到但无法建模的内容。                       |
| `readonly-change`      | `{ readOnly }`                                                                             |                                                      |
| `ribbon-action`        | `{ id }`                                                                                   | 执行了一个功能区命令。                               |
| `file-command`         | `{ command: 'new' \| 'open' \| 'save' \| 'saveAs' \| 'export' \| 'exportCsv' \| 'print' }` | **可取消**：调用 `preventDefault()` 以自行处理。     |
| `selection-change`     | `{ sheet, ref, active }`                                                                   | `ref` 为所选区域，`active` 为活动单元格。            |
| `sheet-change`         | `{ index, name }`                                                                          |                                                      |
| `dirty-change`         | `{ dirty }`                                                                                |                                                      |
| `office-ribbon-add-in` | `{ tab, command }`                                                                         | 选择了宿主选项卡（`ribbonAddIns`）的一个命令。       |
| `collaboration-change` | `{ active, roomId, status, synced, people }`                                               | 共享开始或停止、连接变化，或有人加入或离开。         |
| `validation-error`     | `{ message, title?, ref, style? }`                                                         | 单元格编辑未通过校验，保持打开状态，直到更正或取消。 |

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

## Shadow part

该元素将其各个区域公开为 CSS shadow part，可用 `xlsx-editor::part(grid)` 设置样式，也可用于测试：`ribbon`、`name-box`、`formula-bar`、`grid`、`sheet-tabs` 和 `status-bar`。
