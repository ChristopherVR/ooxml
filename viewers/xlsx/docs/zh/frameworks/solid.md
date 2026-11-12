# Solid

Solid 适配器是一个带有共享属性和回调的组件。

```tsx
import { createSignal } from 'solid-js';
import { createWorkbook, SpreadsheetEditor, type EditorHandle } from 'xlsx-solid-viewer';

export function Editor() {
	const [workbook, setWorkbook] = createSignal(createWorkbook());
	let editor: EditorHandle | undefined;
	return (
		<SpreadsheetEditor
			workbook={workbook()}
			onWorkbookChange={setWorkbook}
			editorRef={(handle) => (editor = handle)}
			class="sheet"
		/>
	);
}
```

`editorRef` 接收[句柄](/zh/bindings#handle)。组件卸载时，Solid 所有者会清理编辑器及其监听器。

参见[完整的绑定契约](/zh/bindings)，或[试用 Solid 演示](/demo-solid/)。
