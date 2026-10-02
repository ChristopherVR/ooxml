# React

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `@christophervr/xlsx-react-viewer` package; it needs only `react` next to it.
:::

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
