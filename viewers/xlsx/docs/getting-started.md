# Getting started

## Run the demo from source

You need [Bun](https://bun.sh/) 1.3 and Node.js 24.

```bash
git clone https://github.com/ChristopherVR/xlsx-viewer.git
cd xlsx-viewer
bun install
bun run demo   # vanilla demo; add ?framework=react|vue|angular|svelte|solid
```

The landing page opens an `.xlsx`, `.xlsm`, legacy `.xls` or `.csv` file (drop it, or browse), starts a new workbook, or opens the sample workbook. Files are processed in the browser only.

## Pick a package

Install one self-contained editor package for your framework. Each bundles the `<xlsx-editor>` web component and its adapter, depends on the workbook model (`@christophervr/xlsx-core`), `ooxml-core` and `ooxml-ui`, and re-exports the model API (`createWorkbook`, `loadWorkbook`, `saveWorkbook`, ...), so a single install and a single import path are all an application needs.

| Framework | Package                            | Component                      |
| --------- | ---------------------------------- | ------------------------------ |
| React     | `@christophervr/xlsx-react-viewer` | `SpreadsheetEditor`            |
| Vue 3     | `xlsx-vue-viewer`                  | `SpreadsheetEditor`            |
| Angular   | `xlsx-angular-viewer`              | `SpreadsheetEditorComponent`   |
| Svelte 5  | `xlsx-svelte-viewer`               | default export, `/runtime`     |
| Solid     | `xlsx-solid-viewer`                | `SpreadsheetEditor`            |
| Vanilla   | `xlsx-vanilla-viewer`              | `mountEditor`, `<xlsx-editor>` |

Install `@christophervr/xlsx-core` on its own only for headless use (parsing, calculating and saving without an editor).

## Mount an editor

```tsx
import { useState } from 'react';
import { createWorkbook, SpreadsheetEditor } from '@christophervr/xlsx-react-viewer';

export function Editor() {
	const [workbook, setWorkbook] = useState(() => createWorkbook());
	return (
		<div style={{ height: '100vh' }}>
			<SpreadsheetEditor workbook={workbook} onWorkbookChange={setWorkbook} />
		</div>
	);
}
```

Give the container a height: the grid fills it and scrolls inside. To open a file, pass its bytes (`bytes={uint8array}` with `fileName`), a URL (`src="/files/budget.xlsx"`), or call `load(bytes, fileName)` on the component handle.

## Save

`save()` resolves to an `.xlsx` `Blob`, `saveBytes('csv')` exports the active sheet as CSV, and `download(name?)` saves and starts a browser download. `dirty` reports unsaved edits; call `markClean()` after you persist a saved Blob yourself. A legacy `.xls` file is saved as `.xlsx`.

Next: the [framework bindings](/bindings), the [element API](/api), [theming](/theming) and the [features and limitations](/features).
