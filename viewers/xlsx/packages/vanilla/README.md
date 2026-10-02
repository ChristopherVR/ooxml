# xlsx-vanilla-viewer

The Excel spreadsheet editor for vanilla JavaScript. One self-contained package: the `<xlsx-editor>` web component, the legacy Excel 97-2003 `.xls` reader are bundled in, and the workbook model (`@christophervr/xlsx-core`: `createWorkbook`, `loadWorkbook`, `saveXlsx`, ...) installs with it and is re-exported, so one install and one import path are all an application needs.

> **Not published to npm yet.** Build from the repository to try it; the commands below are the intended API.

```sh
npm install xlsx-vanilla-viewer
```

```ts
import { mountEditor } from 'xlsx-vanilla-viewer';

const editor = mountEditor(container, {
	fileName: 'Budget.xlsx',
	onWorkbookChange: (workbook) => console.log(workbook.sheets.length),
	onWorkbookError: (error) => console.error(error),
});
await editor.load(bytes, 'Budget.xlsx');
const saved = await editor.save();
editor.destroy();
```

Or use the element directly: `defineXlsxEditor()` registers `<xlsx-editor>`.

The shared option types and the `loadWorkbook` / `detectWorkbookFormat` helpers are exported from the package root.

Notes:

- `.xlsx`, `.xlsm`, legacy `.xls` and `.csv` files are opened; the format is sniffed from the bytes. A legacy `.xls` file is saved back as `.xlsx`.
- `ooxml-core` and `ooxml-ui` are regular dependencies; never import `ooxml-ui` yourself.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<xlsx-editor>` element.
- An early editor: not Microsoft Excel parity, and saving is not lossless for unsupported features (pivot tables, slicers, sparklines and macros are kept but not editable). See the [features and limitations](https://christophervr.github.io/xlsx-viewer/features).

Documentation: [element API](https://christophervr.github.io/xlsx-viewer/api), [vanilla JavaScript guide](https://christophervr.github.io/xlsx-viewer/frameworks/vanilla). Licensed under Apache-2.0.
