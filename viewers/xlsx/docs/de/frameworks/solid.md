# Solid

Der Solid-Adapter ist eine Komponente mit den gemeinsamen Props und Callbacks.

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

`editorRef` erhält das [Handle](/de/bindings#handle). Der Besitzer in Solid räumt den Editor und seine Listener auf, wenn die Komponente entfernt wird.

Siehe die [vollständige Binding-Vertragsbeschreibung](/de/bindings) oder [probieren Sie die Solid-Demo aus](/demo-solid/).
