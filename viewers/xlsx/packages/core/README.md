# @christophervr/xlsx-core

**The Excel workbook model without a browser: read, calculate, edit and write `.xlsx`, `.xlsm`, `.xltx`, legacy `.xls` and `.csv` files in Node.js, workers and scripts.**

[![npm version](https://img.shields.io/npm/v/%40christophervr%2Fxlsx-core.svg)](https://www.npmjs.com/package/@christophervr/xlsx-core)
[![license](https://img.shields.io/npm/l/%40christophervr%2Fxlsx-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE)
[![types](https://img.shields.io/npm/types/%40christophervr%2Fxlsx-core.svg)](https://www.npmjs.com/package/@christophervr/xlsx-core)

[Documentation](https://christophervr.github.io/ooxml/xlsx/) · [Supported API](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md#the-supported-surface-of-docx-core-and-christophervrxlsx-core) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/core/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

This is the model behind the Excel editors as a DOM-free API, and a thin entry point: `@christophervr/xlsx-core` re-exports `ooxml-core/xlsx`, and `@christophervr/xlsx-core/load` re-exports `ooxml-core/xlsx/load`. The code lives in [`ooxml-core`](https://www.npmjs.com/package/ooxml-core), which installs with this package.

## Install

```bash
npm install @christophervr/xlsx-core
```

You need it only for headless work: every Excel editor package (`@christophervr/xlsx-react-viewer`, `xlsx-vue-viewer` and the others) already depends on it and re-exports the same API.

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

`createWorkbook()` starts an empty workbook, and `loadXlsx` / `saveXlsx` read and write `.xlsx` from the main entry without the other loaders. `loadWorkbook` takes `{ password }` for an encrypted workbook, and `saveWorkbook(workbook, 'xlsx', { password })` writes one; `saveWorkbook(workbook, 'csv')` writes the active sheet as UTF-8 CSV.

## API overview

| Import                          | Main entry points                                                                                                                                                                                                                       |
| ------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `@christophervr/xlsx-core`      | `createWorkbook`, `createWorksheet`, `getCell`, `putCell`, `parseAddress`, `formatAddress`; `loadXlsx`, `saveXlsx`; `createEditSession` (editing with undo); `createCalcEngine`; `compareWorkbooks`.                                    |
| `@christophervr/xlsx-core/load` | `loadWorkbook` and `saveWorkbook` for every supported format, `detectWorkbookFormat`, `DataIntegrityError`, and the legacy `.xls` and CSV readers. Kept apart so code that never opens legacy files does not pull in the binary reader. |

- **Bounded calculation.** `createCalcEngine(workbook, { timeLimitMs: 10_000 })` then `recalculateAll()` stops between formulas when the limit is reached and says why in `engine.stopped`; `workCap` stops it the same way at a formula that needs more work than the cap. Each formula also has its own work and text budget.
- **A formula policy for automated writers.** `formulaPolicyIssue`, `cellInputPolicyIssue` and `RESTRICTED_FORMULA_FUNCTIONS` tell you whether a formula would reach outside the workbook (web and external data, macro functions, add-ins, AI functions, other files) before you write it.
- **Host providers.** `WEBSERVICE`, `IMAGE`, the cube functions and the other integration functions run only through providers you pass in `CalcEngineOptions.providers` (`createFetchWebserviceProvider` and `createFetchImageProvider` are opt-in fetch-backed ones, with `allowedOrigins`). Without a provider, cached results are kept on load and newly entered formulas return `#NAME?`.
- **Load reports.** `workbook.warnings` lists what the loader could not read, and `workbook.signatures` the verified signature state of a signed package (saving removes signatures).

The supported surface under semver is listed in [docs/stability.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md#the-supported-surface-of-docx-core-and-christophervrxlsx-core); every other exported name (the `layout/` view logic, the `read/` and `write/` helpers) is internal and can change in a minor release.

## Status and limits

Excel is **Stable**: a frozen API under semver (checked by a committed public-surface snapshot), a [preservation guarantee](https://christophervr.github.io/ooxml/xlsx/preservation) and native Excel save and reopen checks. It does not match Microsoft Excel feature for feature.

- `.xlsb` and `.ods` are not read (`.xlsb` is detected and refused).
- External links, data connections, Power Query and OLAP or external-source PivotTables are kept with their cached values and never refreshed. Macros are kept, never run.
- Parts the model does not cover are written back unchanged where their sheet still exists, and a chart the model cannot carry keeps its source bytes and is reported; anything else that cannot be kept is reported in `warnings`, not silently dropped.
- The engine has limits Excel does not (for example very large `MMULT` arrays, or a `LAMBDA` recursing deeper than about 700 levels, return `#NUM!`).

See [features and limitations](https://christophervr.github.io/ooxml/xlsx/features), the [function list](https://christophervr.github.io/ooxml/xlsx/functions) and the [parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md).

## Privacy and network use

The model makes no network request when it opens, calculates, edits or saves a workbook. Linked pictures are not fetched and external workbook links and linked data are not refreshed. `WEBSERVICE` and `IMAGE` contact a URL only through a provider the host passes in `CalcEngineOptions.providers`, restricted with `allowedOrigins` if you use the built-in fetch providers; omit the provider to keep everything local. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#excel-in-cell-images-and-linked-data).

## Environments

- **Node.js** 22.3 or newer. ESM only.
- **Browsers and web workers:** the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)). The model needs no DOM, so it runs in a worker.

## Documentation

- [Documentation](https://christophervr.github.io/ooxml/xlsx/), with [features and limitations](https://christophervr.github.io/ooxml/xlsx/features), the [function list](https://christophervr.github.io/ooxml/xlsx/functions) and the [preservation guarantee](https://christophervr.github.io/ooxml/xlsx/preservation).
- The Excel editor for [React](https://www.npmjs.com/package/@christophervr/xlsx-react-viewer), [Vue](https://www.npmjs.com/package/xlsx-vue-viewer), [Angular](https://www.npmjs.com/package/xlsx-angular-viewer), [Svelte](https://www.npmjs.com/package/xlsx-svelte-viewer), [Solid](https://www.npmjs.com/package/xlsx-solid-viewer) and [plain JavaScript](https://www.npmjs.com/package/xlsx-vanilla-viewer).
- [`xlsx-viewer-mcp`](https://www.npmjs.com/package/xlsx-viewer-mcp): Excel tools for AI agents over MCP.
- [`ooxml-core`](https://www.npmjs.com/package/ooxml-core): the engine this package re-exports, and the [ooxml repository](https://github.com/ChristopherVR/ooxml).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE). See [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/NOTICE) for attributions.
