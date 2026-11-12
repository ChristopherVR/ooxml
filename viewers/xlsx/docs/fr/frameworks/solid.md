# Solid

L’adaptateur Solid est un composant doté des props et rappels communs.

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

`editorRef` reçoit la [poignée](/fr/bindings#handle). Le propriétaire Solid nettoie l’éditeur et ses écouteurs lorsque le composant est démonté.

Consultez le [contrat complet des liaisons](/fr/bindings) ou [essayez la démo Solid](/demo-solid/).
