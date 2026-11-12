# React

L’adaptateur React monte l’élément commun `<xlsx-editor>` et transmet les mises à jour du classeur par `onWorkbookChange`.

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

Passez `bytes` (avec `fileName`) ou `src` pour ouvrir un fichier. La ref expose la [poignée](/fr/bindings#handle) : `load`, `save`, `saveBytes`, `download`, `select` et l’élément lui-même.

Consultez le [contrat complet des liaisons](/fr/bindings) ou [essayez la démo React](/demo/).
