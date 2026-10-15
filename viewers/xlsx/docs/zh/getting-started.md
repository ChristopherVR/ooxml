# 快速开始

## 从源码运行演示

需要 [Bun](https://bun.sh/) 1.3 和 Node.js 24。

```bash
git clone https://github.com/ChristopherVR/ooxml.git
cd ooxml/viewers/xlsx
bun install
bun run demo   # vanilla demo; add ?framework=react|vue|angular|svelte|solid
```

落地页可以打开 `.xlsx`、`.xlsm`、旧版 `.xls` 或 `.csv` 文件（拖放或浏览选择），也可以新建工作簿或打开示例工作簿。文件只在浏览器中处理。

## 选择软件包

为你的框架安装一个自包含的编辑器软件包。每个软件包都打包了 `<xlsx-editor>` Web 组件及其适配器，依赖工作簿模型（`@christophervr/xlsx-core`）、`ooxml-core` 和 `ooxml-ui`，并重新导出模型 API（`createWorkbook`、`loadWorkbook`、`saveWorkbook` 等），因此应用只需一次安装和一个导入路径。

| 框架     | 软件包                             | 组件                           |
| -------- | ---------------------------------- | ------------------------------ |
| React    | `@christophervr/xlsx-react-viewer` | `SpreadsheetEditor`            |
| Vue 3    | `xlsx-vue-viewer`                  | `SpreadsheetEditor`            |
| Angular  | `xlsx-angular-viewer`              | `SpreadsheetEditorComponent`   |
| Svelte 5 | `xlsx-svelte-viewer`               | 默认导出，`/runtime`           |
| Solid    | `xlsx-solid-viewer`                | `SpreadsheetEditor`            |
| 原生 JS  | `xlsx-vanilla-viewer`              | `mountEditor`, `<xlsx-editor>` |

只有在无头使用（不带编辑器地解析、计算和保存）时，才需要单独安装 `@christophervr/xlsx-core`。

## 挂载编辑器

```tsx
import { useState } from 'react';
import { createWorkbook, SpreadsheetEditor } from '@christophervr/xlsx-react-viewer';

export function Editor() {
	const [workbook, setWorkbook] = useState(() => createWorkbook());
	return (
		<div style={{ height: '100vh' }}>
			<SpreadsheetEditor workbook={workbook} onWorkbookChange={setWorkbook} />
		</div>
	);
}
```

请为容器指定高度：网格会填满容器并在其内部滚动。要打开文件，可以传入文件字节（`bytes={uint8array}` 并配合 `fileName`）、URL（`src="/files/budget.xlsx"`），或调用组件句柄上的 `load(bytes, fileName)`。

## 保存

`save()` 返回一个 `.xlsx` 的 `Blob`，`saveBytes('csv')` 将当前工作表导出为 CSV，`download(name?)` 保存并触发浏览器下载。`dirty` 表示存在未保存的编辑；如果你自行持久化了保存得到的 Blob，请调用 `markClean()`。旧版 `.xls` 文件会保存为 `.xlsx`。

下一步：[框架绑定](/bindings)、[元素 API](/zh/api)、[主题](/theming)以及[功能与限制](/zh/features)。
