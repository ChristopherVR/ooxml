# React

The React adapter mounts the shared `<xlsx-editor>` element and forwards workbook updates through `onWorkbookChange`.

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

Pass `bytes` (with `fileName`) or `src` to open a file. The ref exposes the [handle](/bindings#handle): `load`, `save`, `saveBytes`, `download`, `select` and the element itself.

See the [complete binding contract](/bindings) or [try the React demo](/demo/).
