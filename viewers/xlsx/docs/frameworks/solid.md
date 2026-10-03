# Solid

The Solid adapter is a component with the shared props and callbacks.

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

`editorRef` receives the [handle](/bindings#handle). The Solid owner cleans up the editor and its listeners when the component unmounts.

See the [complete binding contract](/bindings) or [try the Solid demo](/demo-solid/).
