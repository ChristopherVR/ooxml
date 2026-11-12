# React

El adaptador de React monta el elemento compartido `<xlsx-editor>` y reenvía las actualizaciones del libro a través de `onWorkbookChange`.

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

Pase `bytes` (con `fileName`) o `src` para abrir un archivo. La ref expone el [handle](/es/bindings#handle): `load`, `save`, `saveBytes`, `download`, `select` y el propio elemento.

Consulte el [contrato completo de los enlaces](/es/bindings) o [pruebe la demo de React](/demo/).
