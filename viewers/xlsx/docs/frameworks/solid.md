# Solid

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `xlsx-solid-viewer` package; it needs only `solid-js` next to it.
:::

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
