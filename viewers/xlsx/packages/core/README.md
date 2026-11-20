# @christophervr/xlsx-core

**The Excel workbook model without a browser: read, calculate, edit and write `.xlsx` workbooks in Node.js, workers and scripts.**

[![npm version](https://img.shields.io/npm/v/%40christophervr%2Fxlsx-core.svg)](https://www.npmjs.com/package/@christophervr/xlsx-core)
[![license](https://img.shields.io/npm/l/%40christophervr%2Fxlsx-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE)
[![docs](https://img.shields.io/badge/docs-christophervr.github.io-1f9d63.svg)](https://christophervr.github.io/ooxml/xlsx/)

[Live demo](https://christophervr.github.io/ooxml/xlsx/demo/) · [Documentation](https://christophervr.github.io/ooxml/xlsx/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/core/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

This is the model behind the Excel editors, as a DOM-free API. It is a thin entry point: `@christophervr/xlsx-core` re-exports `ooxml-core/xlsx`, and `@christophervr/xlsx-core/load` re-exports `ooxml-core/xlsx/load`. The code lives in [`ooxml-core`](https://www.npmjs.com/package/ooxml-core), which installs with this package.

## Install

```bash
npm install @christophervr/xlsx-core
```

You need this package only for headless work. Every Excel editor package (`@christophervr/xlsx-react-viewer`, `xlsx-vue-viewer` and the others) already depends on it and re-exports the same API.

## Quick start

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { createEditSession } from '@christophervr/xlsx-core';
import { loadWorkbook, saveWorkbook } from '@christophervr/xlsx-core/load';

const workbook = await loadWorkbook(await readFile('budget.xlsx'), { fileName: 'budget.xlsx' });

// Sheet, row and column are zero-based: this writes D1 on the first sheet and recalculates.
const session = createEditSession(workbook);
session.setCellInput(0, 0, 3, '=SUM(B1:C1)');

await writeFile('budget-edited.xlsx', await saveWorkbook(workbook, 'xlsx'));
```

`createWorkbook()` starts an empty workbook, and `saveXlsx(workbook)` writes one from the main entry without the loaders. To calculate an untrusted workbook under a time limit, use `createCalcEngine(workbook, { timeLimitMs: 10_000 })`, call `recalculateAll()` and read `engine.stopped`.

## Entry points

| Import                          | What it holds                                                                                                                                                                                                                   |
| ------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `@christophervr/xlsx-core`      | The workbook model, addresses, number formats, the formula parser and calculation engine, editing commands with undo, the `.xlsx` reader and writer (`saveXlsx`), pivots, slicers, workbook comparison and the DOM-free layout. |
| `@christophervr/xlsx-core/load` | `loadWorkbook` and `saveWorkbook` for every supported format, `detectWorkbookFormat`, the legacy `.xls` reader and CSV. Kept apart so code that never opens legacy files does not pull in the binary reader.                    |

## What you get

- **Every common format.** `.xlsx`, `.xlsm` (the VBA project is kept, never run), `.xltx`, password-protected workbooks (read, and written with `saveWorkbook(workbook, 'xlsx', { password })`), legacy `.xls` (read; saved as `.xlsx`) and `.csv` in both directions.
- **A calculation engine.** More than 500 functions, dynamic arrays, `LET` and `LAMBDA`, iterative calculation, and the workbook's automatic or manual mode.
- **Bounded work.** Each formula has its own work and text budget, and `CalcEngineOptions.timeLimitMs` and `workCap` bound a whole recalculation, which then stops between formulas and reports why in `engine.stopped`. `calculateOnOpen` runs the calculation a workbook asks for on open.
- **A formula policy.** `formulaPolicyIssue`, `cellInputPolicyIssue` and `RESTRICTED_FORMULA_FUNCTIONS` tell you whether a formula would reach outside the workbook (web and external data, macro functions, add-ins, AI functions, other files) before you write it.
- **Editing with undo.** `createEditSession` is the editing API the Excel editor itself drives: cell input, styles and borders, rows and columns, merges, sorting and more, each one undo step, with recalculation after every edit.
- **Careful saving.** Parts the model does not cover, such as VBA, pivot caches and chart detail it cannot carry, are written back unchanged where the sheet still exists, and a chart the model cannot carry keeps its source bytes and is reported.
- **Signatures and warnings.** `Workbook.signatures` reports the verified signature state, and `warnings` lists what the loader could not read.

## Supported environments

- **Node.js** 22.3 or newer.
- **Browsers and workers:** the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).

## Limitations

This is not Microsoft Excel. `.xlsb` and `.ods` are not read. External links, Power Query, OLAP and external-source PivotTables are kept but never refreshed, `WEBSERVICE` is not evaluated, and a few functions (`ASC`, `DBCS`, `JIS`, `PHONETIC`, `INFO`) are not implemented. Saving is not lossless for features the model does not cover; the [preservation guarantee](https://christophervr.github.io/ooxml/xlsx/preservation) and the [features page](https://christophervr.github.io/ooxml/xlsx/features) say what is kept, reported or dropped.

## Privacy and network use

The model makes no network request when it opens, calculates, edits or saves a workbook: linked pictures are not fetched, external workbook links and linked data are not refreshed, and `WEBSERVICE` is not evaluated. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

- The Excel editor for [React](https://www.npmjs.com/package/@christophervr/xlsx-react-viewer), [Vue](https://www.npmjs.com/package/xlsx-vue-viewer), [Angular](https://www.npmjs.com/package/xlsx-angular-viewer), [Svelte](https://www.npmjs.com/package/xlsx-svelte-viewer), [Solid](https://www.npmjs.com/package/xlsx-solid-viewer) and [plain JavaScript](https://www.npmjs.com/package/xlsx-vanilla-viewer).
- [`xlsx-viewer-mcp`](https://www.npmjs.com/package/xlsx-viewer-mcp): Excel tools for AI agents over MCP.
- [`ooxml-core`](https://www.npmjs.com/package/ooxml-core): the engine this package re-exports, with Word, PowerPoint and Visio next to Excel.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE)
