# @christophervr/xlsx-core

[![npm version](https://img.shields.io/npm/v/%40christophervr%2Fxlsx-core.svg)](https://www.npmjs.com/package/@christophervr/xlsx-core)
[![license](https://img.shields.io/npm/l/%40christophervr%2Fxlsx-core.svg)](https://github.com/ChristopherVR/xlsx-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/%40christophervr%2Fxlsx-core.svg)](https://www.npmjs.com/package/@christophervr/xlsx-core)

> The DOM-free document model and file API, re-exported from the canonical OOXML core.

[Live demo](https://christophervr.github.io/xlsx-viewer/demo/) | [npm](https://www.npmjs.com/package/@christophervr/xlsx-core) | [Full docs](https://christophervr.github.io/xlsx-viewer/) | [Source](https://github.com/ChristopherVR/xlsx-viewer)

The framework-neutral Excel workbook model, parser, formula engine and serializer. A thin entry point: `@christophervr/xlsx-core` re-exports `ooxml-core/xlsx` and `@christophervr/xlsx-core/load` re-exports `ooxml-core/xlsx/load`; the logic lives in [ooxml-core](https://github.com/ChristopherVR/ooxml).

## Install

```bash
npm install @christophervr/xlsx-core
```

## Quick start

```ts
import { createWorkbook, saveXlsx } from '@christophervr/xlsx-core';
import { loadWorkbook } from '@christophervr/xlsx-core/load';

const workbook = createWorkbook();
const bytes = await saveXlsx(workbook);
const reopened = await loadWorkbook(bytes, { fileName: 'Book1.xlsx' });
```

## API and limitations

The main entry handles `.xlsx`, `.xlsm` and `.xltx`. The `/load` entry also opens legacy Excel 97-2003 `.xls` files (read only; they save as `.xlsx`) and `.csv` text, sniffing the format from the bytes. Saving is not lossless for features the model does not cover; unmodelled parts (VBA, pivot tables, charts) are carried through unchanged where the sheet still exists.

## Documentation

[Full docs](https://christophervr.github.io/xlsx-viewer/) | [Core source](https://github.com/ChristopherVR/ooxml)

## License

Apache-2.0.
