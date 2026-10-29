<div align="center">

# docx-viewer

[![npm version](https://img.shields.io/npm/v/docx-react-viewer.svg)](https://www.npmjs.com/package/docx-react-viewer)
[![license](https://img.shields.io/npm/l/docx-react-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE)
[![types](https://img.shields.io/npm/types/docx-react-viewer.svg)](https://www.npmjs.com/package/docx-react-viewer)

**A browser Word document editor with one document model, one web component and thin adapters for your framework.**
An early implementation: not Microsoft Word parity, and not lossless export.

[![docs](https://img.shields.io/badge/docs-christophervr.github.io-6366f1.svg)](https://christophervr.github.io/ooxml/docx/)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)

[**Live demo**](https://christophervr.github.io/ooxml/docx/demo/) &nbsp;&middot;&nbsp;
[**Documentation**](https://christophervr.github.io/ooxml/docx/) &nbsp;&middot;&nbsp;
[**Getting started**](#getting-started) &nbsp;&middot;&nbsp;
[**Packages**](#packages)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; **[Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme)** &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; [OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme) &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

![The docx-viewer editor showing a sample document with the Word-style ribbon](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/docx/docs/public/hero-editor.webp)

</div>

## Why docx-viewer?

- **One editor, every framework.** A ProseMirror-backed `<docx-editor>` web component owns rendering, selection, commands, history and styling. React, Vue, Angular, Svelte, Solid and vanilla adapters only handle lifecycle and events.
- **One document model.** A framework-neutral DOCX model, parser and serializer sit under the editor.
- **Careful preservation.** A no-op save returns the original bytes, supported edits keep package parts outside the model, and unsafe edits are rejected instead of silently dropped.
- **Honest limits.** Unsupported features are reported, not hidden. See the [support roadmap](docs/parity-roadmap.md) and [outstanding work](docs/outstanding-work.md).

## Feature tour

|                        |                                                                                                                                                                                                                                                                                   |
| ---------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Editor**             | Word-style ribbon, title bar, File backstage, status bar, undo/redo, find and replace, zoom and read-only mode in one web component.                                                                                                                                              |
| **Print Layout**       | Paginated view with line breaking, widow/orphan and keep rules, columns, table row splitting and printing. An approximation of Word, not Word's pagination.                                                                                                                       |
| **Document structure** | Styles, lists and numbering, tables, sections and columns, headers and footers, footnotes and fields such as TOC. Inline text boxes, header watermark and page borders. SmartArt and charts can be inserted and edited, and charts are drawn from their saved values (see below). |
| **Review**             | Tracked changes (accept, reject, navigate) and a comments pane with replies and resolve.                                                                                                                                                                                          |
| **File formats**       | DOCX, and legacy Word 97-2003 `.doc` via the shared `ole2` codecs inside `ooxml-core` (main-body text and constrained paragraph edits). Password-protected files are unsupported.                                                                                                 |
| **Collaboration**      | A transport-neutral protocol over validated ProseMirror steps with transient presence. Your application owns networking, identity, permissions and storage.                                                                                                                       |
| **Localization**       | Editor interface in English, French, German, Spanish and Simplified Chinese through a `locale` option. Document content is never translated.                                                                                                                                      |

## Getting started

### 1. Install

```bash
npm install docx-react-viewer
```

Choose the package for your framework in the package table below. The editor
installs its core and shared UI dependencies and re-exports its model API.
For source development, see [Development](#development).

### 2. Mount the editor

Install one self-contained editor package for your framework, for example `npm install docx-react-viewer react`. The editor packages are `docx-react-viewer`, `-vue-viewer`, `-angular-viewer`, `-svelte-viewer`, `-solid-viewer` and `-vanilla-viewer`. Each one bundles the editor UI and adapter, delegates layout and legacy `.doc` loading to core, brings the document model (`docx-core`) with it, and re-exports it, so a single install and a single import path are all an application needs. Install `docx-core` on its own only for headless use (parsing and serializing without an editor).

```tsx
import { createDocument, WordEditor } from 'docx-react-viewer';

export function Editor() {
	const [model, setModel] = useState(() => createDocument());
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

<details>
<summary><strong>Vue 3</strong></summary>

```vue
<script setup lang="ts">
import { ref } from 'vue';
import { createDocument, WordEditor } from 'docx-vue-viewer';

const model = ref(createDocument());
</script>
<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

</details>

<details>
<summary><strong>Angular</strong></summary>

```ts
import { createDocument, WordEditorComponent } from 'docx-angular-viewer';

@Component({
	standalone: true,
	imports: [WordEditorComponent],
	template: '<word-editor [documentModel]="model" (documentChange)="model = $event" />',
})
export class EditorComponent {
	model = createDocument();
}
```

</details>

<details>
<summary><strong>Svelte 5</strong></summary>

```svelte
<script>
  import WordEditor from 'docx-svelte-viewer';
  import { createDocument } from 'docx-svelte-viewer/runtime';
  let model = $state(createDocument());
</script>
<WordEditor documentModel={model} ondocumentchange={next => model = next} />
```

</details>

<details>
<summary><strong>SolidJS</strong></summary>

```tsx
import { createSignal } from 'solid-js';
import { createDocument, WordEditor } from 'docx-solid-viewer';

export function Editor() {
	const [model, setModel] = createSignal(createDocument());
	return <WordEditor documentModel={model()} onDocumentChange={setModel} />;
}
```

</details>

<details>
<summary><strong>Vanilla JavaScript</strong></summary>

```ts
import { createDocument, mountEditor } from 'docx-vanilla-viewer';

const editor = mountEditor(container, {
	documentModel: createDocument(),
	onDocumentChange: (model) => console.log(model),
	onDocumentError: (error) => console.error(error),
});
await editor.load(bytes);
const saved = await editor.save();
editor.destroy();
```

</details>

See the [bindings guide](docs/bindings.md) for props, events, saving and file commands.

## Packages

Seven packages are published on npm, each versioned independently:

| Package               | What it is                                                                                                            |
| --------------------- | --------------------------------------------------------------------------------------------------------------------- |
| `docx-core`           | DOCX model, parser and preserving serializer (`packages/core`, a thin entry point over `ooxml-core/docx`). DOCX only. |
| `docx-react-viewer`   | React component (`packages/react`).                                                                                   |
| `docx-vue-viewer`     | Vue 3 component (`packages/vue`).                                                                                     |
| `docx-angular-viewer` | Angular standalone component (`packages/angular`).                                                                    |
| `docx-svelte-viewer`  | Svelte 5 component (`packages/svelte`).                                                                               |
| `docx-solid-viewer`   | Solid component (`packages/solid`).                                                                                   |
| `docx-vanilla-viewer` | `mountEditor` and the plain `<docx-editor>` web component, no framework (`packages/vanilla`).                         |

Each `*-viewer` package is self-contained: the editor, layout engine, document loading and the legacy `.doc` reader come from `ooxml-core` (`/docx/layout`, `/docx/load`, which inlines the shared ole2 codecs). A package depends on `docx-core`, `ooxml-core`, `ooxml-ui` (the shared web components that draw SmartArt and the chart elements; installed for you, never imported by your code), the ProseMirror libraries and its framework peer, so `@christophervr/ole2` does not need to be published or installed. DOCX and legacy `.doc` files are both opened by every editor package.

The workspace also holds **private** packages that are never published and are inlined into each editor package at build time: `web-component` (a re-export of the shared `<docx-editor>`, including the browser text measurer behind Print Layout) and `bindings` (framework lifecycle and event adapters). The pagination engine, document loading and collaboration helpers are in `ooxml-core`. The `<docx-editor>` element itself lives in `ooxml-ui/docx` (`src/ui/src/docx`), which `web-component` re-exports; the document logic is in `ooxml-core`.

## SmartArt

A SmartArt graphic (inline or floating) is displayed from the drawing the producing application saved in the file (`dsp:drawing`), drawn as SVG by `<office-ui-smartart>`. Until it is edited, that drawing is a snapshot: its layout is **not recomputed**, and it is not claimed to look like Word. Preset outlines the renderer cannot draw appear as rectangles, gradient and pattern fills are approximated, transparency and 3D are dropped, and text is one line per shape without autofit; each case is listed in the diagram's info popover (the small `i` button) together with the file's own notices. An unedited diagram keeps its parts and markup byte for byte on save, including when other text is edited. The accessible name comes from the diagram's alt text, title or object name.

Selecting a body SmartArt shows the **SmartArt Design** tab: Add Shape (after, before, above, below), Promote, Demote, Move Up, Move Down, the shared **Text Pane** (`office-ui-smartart-text-pane`: type to change a node, Enter adds one, Tab and Shift+Tab demote and promote, Backspace in an empty line removes it) and **Change Layout**. **Insert > SmartArt** opens the shared picker (`office-ui-smartart-picker`) over 35 built-in layouts. Edits go through the core (`ooxml-core/docx` `editDocxDiagram`) and are undoable. An edited, inserted or re-laid-out diagram is drawn with an **approximate** layout computed by the editor from its nodes, not Word's; on save its data part is rewritten and its stale cached drawing is dropped, because Word ignores that drawing and lays the diagram out again when it opens the file (verified with desktop Word; see `docs/agnostic-core-plan.md` at the repository root, and the Word acceptance script `src/core/scripts/docx-word-acceptance.ps1`, which opens the edited documents in desktop Word). SmartArt Styles, Change Colors, picture layouts' pictures, per-shape formatting and SmartArt in headers, footers and notes are not editable here.

## Charts

Charts are drawn from the values cached in their part: column, bar, line, area, pie, doughnut, scatter, bubble, stock, radar and their combinations, and the waterfall, histogram, treemap and sunburst families from their chartex parts. Surface, volume-stock and the other chartex families show a labelled frame.

**Insert > Chart** opens the shared chart picker (`office-ui-chart-picker`) and inserts an inline chart of eight families (column, bar, line, area, pie, doughnut, scatter and radar) with Office's sample data. A selected chart shows the **Chart Design** tab: **Legend**, **Edit Data** (the shared data sheet, `office-ui-chart-data-sheet`, applied at once), **Change Chart Type** and **Format Selection**, which docks the shared Format pane (`office-ui-chart-format-pane`) beside the page. Format Chart Area and Edit Data are also on the context menu, double-clicking a chart opens Edit Data, and Delete removes it. The corner handle resizes a chart, with Shift to keep its aspect ratio. Each edit is one undo step.

Saving writes an edited chart part with the shared chart writer and updates the cells of its embedded workbook. An unedited chart keeps its parts byte for byte. Not offered: chart styles, colours and Quick Layout, Add Chart Element, Switch Row/Column and Select Data, and the Format pane's picture section. Admitted ChartEx charts are editable like classic charts (waterfall, histogram, treemap and sunburst are drawn; other families show a labelled frame), floating charts are not placed on Print Layout pages, and only the bottom-right corner handle resizes. Desktop Word checks the saved charts with `src/core/scripts/docx-word-chart-acceptance.ps1` (documents open without repair, Word reads back the series and values, and Edit Data opens each embedded workbook).

## Development

```bash
bun run typecheck       # tsc and svelte-check
bun run test            # Vitest
bun run check:shared    # shared-code boundary check
bun run build:packages
bun run check:published   # no tarball may import an internal package or ole2
bun run pack:smoke
bunx playwright install chromium
bun run test:browser    # Playwright contract tests
bun run fmt:check       # oxfmt
bun install --cwd docs && bun run --cwd docs docs:build   # site and all demos
```

`PLAYWRIGHT_CHROMIUM_EXECUTABLE` selects an explicit Chromium and `PLAYWRIGHT_PORT` (default 4180) the preview port. The packages are published to npm by the repository's shared release flow; see [releasing](docs/releasing.md).

## Documentation

[Architecture](docs/architecture.md) &middot; [Framework bindings](docs/bindings.md) &middot; [Editing text](docs/editing.md) &middot; [Collaboration](docs/collaboration.md) &middot; [Support roadmap](docs/parity-roadmap.md) &middot; [Outstanding work](docs/outstanding-work.md)

## License

[Apache License 2.0](LICENSE). See [`NOTICE`](NOTICE) for attributions.

Custom host controls: see [composable building blocks](docs/building-blocks.md).
