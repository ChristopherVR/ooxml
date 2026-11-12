# Solid

El adaptador de Solid es un componente con las propiedades y callbacks compartidos.

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

`editorRef` recibe el [handle](/es/bindings#handle). El propietario de Solid limpia el editor y sus listeners cuando el componente se desmonta.

Consulte el [contrato completo de los enlaces](/es/bindings) o [pruebe la demo de Solid](/demo-solid/).
