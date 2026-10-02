# @christophervr/xlsx-core

The framework-neutral Excel workbook model, parser, formula engine and serializer. A thin entry point: `@christophervr/xlsx-core` re-exports `ooxml-core/xlsx` and `@christophervr/xlsx-core/load` re-exports `ooxml-core/xlsx/load`; the logic lives in [ooxml-core](https://github.com/ChristopherVR/ooxml).

```sh
npm install @christophervr/xlsx-core
```

```ts
import { createWorkbook, saveXlsx } from '@christophervr/xlsx-core';
import { loadWorkbook } from '@christophervr/xlsx-core/load';

const workbook = createWorkbook();
const bytes = await saveXlsx(workbook);
const reopened = await loadWorkbook(bytes, { fileName: 'Book1.xlsx' });
```

The main entry handles `.xlsx`, `.xlsm` and `.xltx`. The `/load` entry also opens legacy Excel 97-2003 `.xls` files (read only; they save as `.xlsx`) and `.csv` text, sniffing the format from the bytes. Saving is not lossless for features the model does not cover; unmodelled parts (VBA, pivot tables, charts) are carried through unchanged where the sheet still exists.
