# React

Der React-Adapter bindet das gemeinsame Element `<xlsx-editor>` ein und leitet Aktualisierungen der Arbeitsmappe über `onWorkbookChange` weiter.

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

Übergeben Sie `bytes` (mit `fileName`) oder `src`, um eine Datei zu öffnen. Die Ref stellt das [Handle](/de/bindings#handle) bereit: `load`, `save`, `saveBytes`, `download`, `select` und das Element selbst.

Siehe die [vollständige Binding-Vertragsbeschreibung](/de/bindings) oder [probieren Sie die React-Demo aus](/demo/).
