# React

React 适配器挂载共享的 `<xlsx-editor>` 元素，并通过 `onWorkbookChange` 转发工作簿更新。

```tsx
import { useRef, useState } from 'react';
import {
	createWorkbook,
	SpreadsheetEditor,
	type EditorHandle,
} from '@christophervr/xlsx-react-viewer';

export function Editor() {
	const [workbook, setWorkbook] = useState(() => createWorkbook());
	const editor = useRef<EditorHandle>(null);
	return (
		<>
			<button onClick={() => void editor.current?.download('Budget.xlsx')}>Download</button>
			<SpreadsheetEditor
				ref={editor}
				workbook={workbook}
				onWorkbookChange={setWorkbook}
				className="sheet"
			/>
		</>
	);
}
```

传入 `bytes`（配合 `fileName`）或 `src` 即可打开文件。ref 公开了[句柄](/zh/bindings#handle)：`load`、`save`、`saveBytes`、`download`、`select`，以及元素本身。

参见[完整的绑定契约](/zh/bindings)，或[试用 React 演示](/demo/)。
