# xlsx-angular-viewer

[![npm version](https://img.shields.io/npm/v/xlsx-angular-viewer.svg)](https://www.npmjs.com/package/xlsx-angular-viewer)
[![license](https://img.shields.io/npm/l/xlsx-angular-viewer.svg)](https://github.com/ChristopherVR/xlsx-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/xlsx-angular-viewer.svg)](https://www.npmjs.com/package/xlsx-angular-viewer)

> A browser Excel spreadsheet editor for Angular, using one shared editor and the canonical OOXML document engine.

[Live demo](https://christophervr.github.io/xlsx-viewer/demo/) | [npm](https://www.npmjs.com/package/xlsx-angular-viewer) | [Full docs](https://christophervr.github.io/xlsx-viewer/) | [Source](https://github.com/ChristopherVR/xlsx-viewer)

## Install

```bash
npm install xlsx-angular-viewer
```

Use the framework peers declared by this package: `@angular/core` (`>=19`). Existing framework apps normally already provide them.

The package bundles its editor UI and adapter. The core model and shared controls
are installed as regular registry dependencies and the model API is re-exported.
One viewer package is enough; no separate core install is required.

## Quick start

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

Mount in the browser and give the host a height. Use one editor package per
application: each registers the same editor custom element.

## Features

| Feature    | Description                                                                         |
| ---------- | ----------------------------------------------------------------------------------- |
| Editing    | Shared grid, ribbon, formula bar, sheet tabs and workbook change events.            |
| Files      | XLSX, XLSM, XLTX, legacy XLS and CSV loading through core.                          |
| Model      | `createWorkbook`, `loadWorkbook`, `saveXlsx` and core types from the package entry. |
| Frameworks | The same `<xlsx-editor>` UI behind six thin adapters.                               |

## API

Shared option types, `loadWorkbook` and `detectWorkbookFormat` helpers are exported
from the package. Framework handles expose `load`, `save` and `download`. The
vanilla package supplies `mountEditor(container, options)` for framework-free
lifecycle management. `ooxml-core` and `ooxml-ui` are regular
dependencies. Parsing, formulas, editing and grid calculations live in core.

## Limitations

This is an early editor, without Microsoft Excel parity or guaranteed lossless
saving for unsupported features. Legacy XLS input saves as XLSX. Pivot tables,
slicers and sparklines are not editable; macros never execute. Review core
diagnostics and the feature guide before relying on preservation.

## Documentation

[Framework guide](https://christophervr.github.io/xlsx-viewer/frameworks/angular) |
[Bindings](https://christophervr.github.io/xlsx-viewer/bindings)

## License

Apache-2.0.
