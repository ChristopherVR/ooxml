<div align="center">

# xlsx-viewer

**An Excel spreadsheet editor that runs entirely in the browser: one workbook model, one `<xlsx-editor>` web component, and thin packages for React, Vue, Angular, Svelte, Solid and plain JavaScript.**

[![npm version](https://img.shields.io/npm/v/%40christophervr%2Fxlsx-react-viewer.svg)](https://www.npmjs.com/package/@christophervr/xlsx-react-viewer)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![docs](https://img.shields.io/badge/docs-christophervr.github.io-1f9d63.svg)](https://christophervr.github.io/ooxml/xlsx/)
[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)

[**Live demo**](https://christophervr.github.io/ooxml/xlsx/demo/) &nbsp;&middot;&nbsp;
[**Documentation**](https://christophervr.github.io/ooxml/xlsx/) &nbsp;&middot;&nbsp;
[**Changelog**](CHANGELOG.md) &nbsp;&middot;&nbsp;
[**Report a bug**](https://github.com/ChristopherVR/ooxml/issues)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#readme) &middot; [Word](../docx#readme) &middot; **Excel** &middot; [PowerPoint](../pptx#readme) &middot; [Visio](../visio#readme) &middot; [OpenTeams](../teams#readme) &middot; [Core](../../src/core#readme)

![The Excel editor showing a workbook with the ribbon, formula bar, grid and sheet tabs](docs/public/hero-editor.webp)

</div>

## Why this editor

- **One editor, every framework.** `<xlsx-editor>` owns the grid, ribbon, formula bar, sheet tabs, selection, commands, history and styling. The framework packages only handle lifecycle and events, so a fix reaches all six at once.
- **A real workbook model.** The model, formula engine, reader and writer are the `xlsx` area of [`ooxml-core`](../../src/core#readme): DOM-free, usable on a server, and the same code the editor runs.
- **Safe with files you did not write.** Each formula has a work budget, a slow calculation on open runs in a worker, every package is read through a guarded zip reader, and macros never run.
- **Careful saving.** Parts the model does not cover (VBA, pivot caches, chart detail, custom XML) are carried through on save where the sheet still exists, and what cannot be kept is reported.
- **Honest limits.** Stable under the [stability policy](../../docs/stability.md), with a written [preservation guarantee](docs/preservation.md) and native Excel save and reopen checks, but not Microsoft Excel parity: the [feature list](docs/features.md) says what is and is not supported.

## Packages

| Package                                                     | What it is                                                         |
| ----------------------------------------------------------- | ------------------------------------------------------------------ |
| [`@christophervr/xlsx-react-viewer`](packages/react#readme) | React component.                                                   |
| [`xlsx-vue-viewer`](packages/vue#readme)                    | Vue 3 component.                                                   |
| [`xlsx-angular-viewer`](packages/angular#readme)            | Angular standalone component.                                      |
| [`xlsx-svelte-viewer`](packages/svelte#readme)              | Svelte 5 component.                                                |
| [`xlsx-solid-viewer`](packages/solid#readme)                | Solid component.                                                   |
| [`xlsx-vanilla-viewer`](packages/vanilla#readme)            | `mountEditor` and the plain `<xlsx-editor>` element, no framework. |
| [`@christophervr/xlsx-core`](packages/core#readme)          | The workbook model on its own, for servers and scripts.            |
| [`xlsx-viewer-mcp`](mcp#readme)                             | MCP server with Excel tools for AI agents.                         |

Each editor package is self-contained: it bundles the editor, depends on `@christophervr/xlsx-core`, `ooxml-core` and `ooxml-ui`, and re-exports the model API (`createWorkbook`, `loadWorkbook`, `saveWorkbook`, ...), so an application needs one install and one import path. The Excel packages keep the `@christophervr` scope where the unscoped names belong to others.

## Getting started

```bash
npm install @christophervr/xlsx-react-viewer react
```

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

Give the editor a height; the grid fills it. Pass `bytes` (with `fileName`) or `src` to open a file, and use the handle (`ref`) to `save()`, `saveBytes('csv')` or `download()`.

<details>
<summary><strong>Vue 3</strong></summary>

```vue
<script setup lang="ts">
import { shallowRef } from 'vue';
import { createWorkbook, SpreadsheetEditor } from 'xlsx-vue-viewer';

const workbook = shallowRef(createWorkbook());
</script>

<template>
	<SpreadsheetEditor :workbook="workbook" @workbook-change="workbook = $event" />
</template>
```

</details>

<details>
<summary><strong>Angular</strong></summary>

```ts
import { Component } from '@angular/core';
import { createWorkbook, SpreadsheetEditorComponent } from 'xlsx-angular-viewer';

@Component({
	selector: 'app-editor',
	standalone: true,
	imports: [SpreadsheetEditorComponent],
	template: '<spreadsheet-editor [workbook]="workbook" (workbookChange)="workbook = $event" />',
})
export class EditorComponent {
	workbook = createWorkbook();
}
```

</details>

<details>
<summary><strong>Svelte 5</strong></summary>

```svelte
<script>
	import XlsxEditor from 'xlsx-svelte-viewer';
	import { createWorkbook } from 'xlsx-svelte-viewer/runtime';

	let workbook = $state.raw(createWorkbook());
</script>

<XlsxEditor {workbook} onworkbookchange={(next) => (workbook = next)} />
```

</details>

<details>
<summary><strong>Solid</strong></summary>

```tsx
import { createSignal } from 'solid-js';
import { createWorkbook, SpreadsheetEditor } from 'xlsx-solid-viewer';

export function Editor() {
	const [workbook, setWorkbook] = createSignal(createWorkbook());
	return <SpreadsheetEditor workbook={workbook()} onWorkbookChange={setWorkbook} />;
}
```

</details>

<details>
<summary><strong>Plain JavaScript</strong></summary>

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

</details>

The [bindings guide](https://christophervr.github.io/ooxml/xlsx/bindings) lists every prop, event and handle member, and the [element API](https://christophervr.github.io/ooxml/xlsx/api) covers `<xlsx-editor>` itself. To build your own interface around the grid, see [building blocks](docs/building-blocks.md).

## Features

The areas below are covered by unit, browser and native Excel tests. Content the model does not cover is kept through a save or reported, as the [preservation guarantee](docs/preservation.md) describes.

|                   |                                                                                                                                                                                                                                                                                                                                                                                                                                |
| ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Editor**        | Ribbon with Tell me, Name Box, formula bar, grid with frozen panes, sheet tabs, status bar (sum, average, count), File backstage, undo and redo, Find and Replace, tables, outline groups, sheet and workbook protection, read-only mode.                                                                                                                                                                                      |
| **Formulas**      | More than 500 functions with dynamic arrays, `LET`, `LAMBDA` and `REGEX`; automatic or manual calculation, Calculate Now and Calculate Sheet, Show Formulas, iterative calculation saved with the workbook. Each formula has a work budget, and a slow calculation on open runs in a worker. Unknown functions show `#NAME?`; circular references are reported.                                                                |
| **Formatting**    | Fonts, fills, borders, alignment, number formats, merged cells, column widths, row heights, conditional formatting (rules, colour scales, data bars, icon sets), data validation lists.                                                                                                                                                                                                                                        |
| **Objects**       | Charts drawn as SVG from live values, inserted from the shared chart picker and edited from the Chart Design tab and Format pane. SmartArt with a text pane, shapes, connectors, groups and pictures. Form controls drawn with their linked cell state, and the interactive ones write to it. Sparklines, worksheet PivotTables and slicers. Ink drawn from the Draw tab saves as Excel ink; OLE objects show as placeholders. |
| **File formats**  | `.xlsx`, `.xlsm` (macros never run, kept on save), `.xltx`, password-protected workbooks, legacy Excel 97-2003 `.xls` (read; saved as `.xlsx`) and `.csv`.                                                                                                                                                                                                                                                                     |
| **Collaboration** | Real-time co-editing of cells, formats, sheets, sizes, merges, names, notes, hyperlinks, conditional formats, validation, tables and charts, between tabs of one browser or through a y-websocket compatible server you run. Pictures, other drawings and page setup stay local. See [collaboration](docs/collaboration.md).                                                                                                   |
| **Accessibility** | F6 regions and Excel's shortcuts, a grid that reads as a table to screen readers, charts named by their alt text, 3:1 focus rings and form controls, and forced-colours support. Checked in a real browser before every release.                                                                                                                                                                                               |
| **Localization**  | Interface in English, French, German, Spanish and Simplified Chinese through `locale`. Workbook content is never translated.                                                                                                                                                                                                                                                                                                   |
| **Not supported** | `.xlsb` and `.ods`. PivotTables from external or OLAP sources, calculated fields, label and value filters, external links and Power Query are kept on save but not shown or refreshed. Sparkline style galleries are not available. `WEBSERVICE`, `ASC`, `DBCS`, `JIS`, `PHONETIC` and `INFO` are not implemented.                                                                                                             |

## Supported environments

- **Browsers:** the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](../../docs/browser-support.md)).
- **Node.js:** 22.3 or newer, for the model, the MCP server, builds and server rendering.
- **Workers:** the calculation on open uses a module worker from the page's own origin; a Content Security Policy must allow `worker-src 'self'`.

## Privacy and network use

Screenshot and Screen Clipping use the browser’s surface picker and embed the result locally. See the [capture privacy details](../../PRIVACY.md#screenshots-and-screen-recordings).

Excel opens, renders, edits and saves workbooks locally by default. Workers, bundled equation code, the selected non-English language and chart or PivotTable tools load from the page's own origin on demand; English is bundled and is used if the selected language cannot load. Linked pictures are not fetched, external workbook links and linked data are not refreshed, `WEBSERVICE` is not evaluated, and fonts come from the device and the file unless the reader turns on the Trust Center font download (off by default) or the host sets `googleFonts`. Hyperlinks open only when the reader follows them, collaboration talks only to the server the host configures, and OneDrive only when the host connects it. Nothing collects analytics or telemetry. See [PRIVACY.md](../../PRIVACY.md).

Picture Format and Shape Format load bundled editing code and galleries on demand. Compress
Pictures loads a worker from the app's origin and processes the image bytes on the device.

## OneDrive and your own storage

`storage` sends File > Save and Open to your own backend, and `fileProviders` adds OneDrive (through `ooxml-core/embedding/onedrive` and a Microsoft sign-in your host supplies) or your own cloud places to File > Open. See the [integration reference](../../docs/onedrive-integration.md#excel-xlsx) for the load and save API, sharing permissions and conflict recovery, and the [testing guide](../../docs/onedrive-testing.md). OneDrive storage does not provide Microsoft Office coauthoring or automatic merging of documents.

## Development

```bash
bun install             # once, at the repository root (one workspace, one lockfile)
cd viewers/xlsx
bun run demo            # vanilla demo; add ?framework=react|vue|angular|svelte|solid
bun run typecheck       # tsc and svelte-check
bun run test            # Vitest (binding contract tests)
bun run test:scripts    # publish guards and the viewer script tests
bun run check:shared    # shared-code boundary check
bun run build:packages
bun run check:published # no tarball may import an internal package or ole2
bun run pack:smoke
bunx playwright install chromium
bun run test:browser    # Playwright contract tests
bun run fmt:check       # oxfmt
bun install --cwd docs && bun run --cwd docs docs:build   # site and all demos
```

`PLAYWRIGHT_CHROMIUM_EXECUTABLE` selects an explicit Chromium and `PLAYWRIGHT_PORT` (default 4180) the preview port. The workspace also holds two private packages that are bundled into each editor package and never published: `web-component` (the shared `<xlsx-editor>`) and `bindings` (the framework adapters). See the [release policy](docs/releasing.md).

## Documentation

[Getting started](https://christophervr.github.io/ooxml/xlsx/getting-started) &middot; [Architecture](https://christophervr.github.io/ooxml/xlsx/architecture) &middot; [Framework bindings](https://christophervr.github.io/ooxml/xlsx/bindings) &middot; [Element API](https://christophervr.github.io/ooxml/xlsx/api) &middot; [Theming](https://christophervr.github.io/ooxml/xlsx/theming) &middot; [Localization](https://christophervr.github.io/ooxml/xlsx/localization) &middot; [Collaboration](https://christophervr.github.io/ooxml/xlsx/collaboration) &middot; [Features and limitations](https://christophervr.github.io/ooxml/xlsx/features)

## License

[Apache License 2.0](LICENSE). See [`NOTICE`](NOTICE) for attributions.
