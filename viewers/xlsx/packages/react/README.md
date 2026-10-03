# @christophervr/xlsx-react-viewer

Available on npm. Install the package using the command below.

The Excel spreadsheet editor for React. One self-contained package: the `<xlsx-editor>` web component, the legacy Excel 97-2003 `.xls` reader and the React adapter are bundled in, and the workbook model (`@christophervr/xlsx-core`: `createWorkbook`, `loadWorkbook`, `saveXlsx`, ...) installs with it and is re-exported, so one install and one import path are all an application needs.

```sh
npm install @christophervr/xlsx-react-viewer react
```

```tsx
import { useState } from 'react';
import { createWorkbook, SpreadsheetEditor } from '@christophervr/xlsx-react-viewer';

export function Editor() {
	const [workbook, setWorkbook] = useState(() => createWorkbook());
	return <SpreadsheetEditor workbook={workbook} onWorkbookChange={setWorkbook} />;
}
```

The shared option types and the `loadWorkbook` / `detectWorkbookFormat` helpers are exported from the package root.

Notes:

- `.xlsx`, `.xlsm`, legacy `.xls` and `.csv` files are opened; the format is sniffed from the bytes. A legacy `.xls` file is saved back as `.xlsx`.
- react is a peer dependency. `ooxml-core` and `ooxml-ui` are regular dependencies; never import `ooxml-ui` yourself.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<xlsx-editor>` element.
- An early editor: not Microsoft Excel parity, and saving is not lossless for unsupported features (pivot tables, slicers, sparklines and macros are kept but not editable). See the [features and limitations](https://christophervr.github.io/xlsx-viewer/features).

Documentation: [element API](https://christophervr.github.io/xlsx-viewer/api), [React guide](https://christophervr.github.io/xlsx-viewer/frameworks/react). Licensed under Apache-2.0.
