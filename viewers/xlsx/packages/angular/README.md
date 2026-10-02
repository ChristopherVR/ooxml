# xlsx-angular-viewer

The Excel spreadsheet editor for Angular. One self-contained package: the `<xlsx-editor>` web component, the legacy Excel 97-2003 `.xls` reader and the Angular adapter are bundled in, and the workbook model (`@christophervr/xlsx-core`: `createWorkbook`, `loadWorkbook`, `saveXlsx`, ...) installs with it and is re-exported, so one install and one import path are all an application needs.

> **Not published to npm yet.** Build from the repository to try it; the commands below are the intended API.

```sh
npm install xlsx-angular-viewer @angular/core
```

```ts
import { Component } from '@angular/core';
import { createWorkbook, SpreadsheetEditorComponent } from 'xlsx-angular-viewer';

@Component({
	standalone: true,
	imports: [SpreadsheetEditorComponent],
	template: '<spreadsheet-editor [workbook]="workbook" (workbookChange)="workbook = $event" />',
})
export class EditorComponent {
	workbook = createWorkbook();
}
```

The shared option types and the `loadWorkbook` / `detectWorkbookFormat` helpers are exported from the package root.

Notes:

- `.xlsx`, `.xlsm`, legacy `.xls` and `.csv` files are opened; the format is sniffed from the bytes. A legacy `.xls` file is saved back as `.xlsx`.
- @angular/core is a peer dependency. `ooxml-core` and `ooxml-ui` are regular dependencies; never import `ooxml-ui` yourself.
- Use one editor package per application: each bundles its own copy of the editor and registers the `<xlsx-editor>` element.
- An early editor: not Microsoft Excel parity, and saving is not lossless for unsupported features (pivot tables, slicers, sparklines and macros are kept but not editable). See the [features and limitations](https://christophervr.github.io/xlsx-viewer/features).

Documentation: [element API](https://christophervr.github.io/xlsx-viewer/api), [Angular guide](https://christophervr.github.io/xlsx-viewer/frameworks/angular). Licensed under Apache-2.0.
