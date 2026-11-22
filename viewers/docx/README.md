<div align="center">

# Word viewer

**A Word document editor for the browser: one `<docx-editor>` element over one document model, with thin packages for React, Vue, Angular, Svelte, Solid and plain JavaScript.**

[![npm version](https://img.shields.io/npm/v/docx-react-viewer.svg)](https://www.npmjs.com/package/docx-react-viewer)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE)
[![docs](https://img.shields.io/badge/docs-christophervr.github.io-6366f1.svg)](https://christophervr.github.io/ooxml/docx/)
[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)

[**Live demo**](https://christophervr.github.io/ooxml/docx/demo/) &middot; [**Documentation**](https://christophervr.github.io/ooxml/docx/) &middot; [**Changelog**](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/CHANGELOG.md) &middot; [**Report a bug**](https://github.com/ChristopherVR/ooxml/issues)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#readme) &middot; **Word** &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; [OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme) &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

![The Word editor with its ribbon, showing a sample document in Print Layout](https://raw.githubusercontent.com/ChristopherVR/ooxml/main/viewers/docx/docs/public/hero-editor.webp)

</div>

## Why this editor

- **One editor, every framework.** The ProseMirror-based `<docx-editor>` owns rendering, selection, commands, history and styling. Each framework package only creates it, forwards props and re-emits its events, so a fix reaches every framework at once.
- **One document model.** A framework-neutral Word model, parser and serializer (`docx-core`, over `ooxml-core/docx`) sits under the editor and works on its own in Node.
- **Careful saving.** An unchanged document saves as its original bytes, supported edits keep the package parts the model does not cover, and an edit that cannot be saved safely is refused instead of silently dropped. See the [preservation guarantee](docs/preservation.md).
- **Stable.** Word is Stable under the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md): a semver-stable API checked against a snapshot, the preservation guarantee, and saved documents opened in desktop Word. Stable is not a claim of Word parity; the [support roadmap](docs/parity-roadmap.md) says what is and is not supported.

## Features

|                         |                                                                                                                                                                                                                          |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Editor**              | Word-style ribbon, title bar with Tell me, File backstage, status bar, context menus, undo and redo, find and replace, zoom, and Editing and Viewing modes in one element.                                               |
| **Print Layout**        | Paginated view with line breaking, widow and orphan control, keep rules, columns, table row splitting, page thumbnails and printing. The pagination is the editor's own and approximates Word's.                         |
| **Document structure**  | Styles, lists and numbering, tables, sections and columns, headers and footers, footnotes and endnotes, fields such as TOC, content controls and form fields, pictures, shapes, text boxes, watermarks and page borders. |
| **Review**              | Tracked changes (accept, reject, navigate) and a comments pane with replies and resolve.                                                                                                                                 |
| **Charts and SmartArt** | Insert and edit both through the shared Office pickers, data sheet and panes (see below).                                                                                                                                |
| **File formats**        | `.docx`; password-protected `.docx` through `loadDocument(bytes, { password })`; legacy Word 97-2003 `.doc` (main-body text, with paragraph text edits saved back as `.doc`).                                            |
| **Collaboration**       | A transport-neutral protocol over validated ProseMirror steps with presence. Your application owns networking, identity, permissions and storage.                                                                        |
| **Customization**       | Hide ribbon controls or whole tabs (the File tab included), add your own ribbon tabs, backstage pages and context menu commands, route Save and Open to your storage, theme with `--dve-*` tokens.                       |
| **Localization**        | English, French, German, Spanish and Simplified Chinese through `locale`. Document content is never translated.                                                                                                          |

## Getting started

Install the package for your framework. Each one brings the editor, `docx-core`, `ooxml-core` and `ooxml-ui` with it and re-exports the model API, so one install and one import path are all an application needs.

```bash
npm install docx-react-viewer
```

```tsx
import { useState } from 'react';
import { createDocument, WordEditor } from 'docx-react-viewer';

export function Editor() {
	const [model, setModel] = useState(createDocument);
	return <WordEditor className="editor" documentModel={model} onDocumentChange={setModel} />;
}
```

Give the editor's container a height; the editor fills it.

<details>
<summary><strong>Vue 3</strong> (<code>docx-vue-viewer</code>)</summary>

```vue
<script setup lang="ts">
import { shallowRef } from 'vue';
import { createDocument, WordEditor } from 'docx-vue-viewer';

const model = shallowRef(createDocument());
</script>

<template>
	<WordEditor class="editor" :document-model="model" @document-change="model = $event" />
</template>
```

Keep the model in a `shallowRef`, not a deep `ref`, so the editor recognises the model it emitted.

</details>

<details>
<summary><strong>Angular</strong> (<code>docx-angular-viewer</code>)</summary>

```ts
import { Component } from '@angular/core';
import { createDocument, WordEditorComponent, type DocumentModel } from 'docx-angular-viewer';

@Component({
	selector: 'app-editor',
	imports: [WordEditorComponent],
	template: `<word-editor
		class="editor"
		[documentModel]="model"
		(documentChange)="model = $event"
	/>`,
	styles: `
		.editor {
			display: block;
			height: 100vh;
		}
	`,
})
export class EditorComponent {
	model: DocumentModel = createDocument();
}
```

</details>

<details>
<summary><strong>Svelte 5</strong> (<code>docx-svelte-viewer</code>)</summary>

```svelte
<script lang="ts">
	import WordEditor from 'docx-svelte-viewer';
	import { createDocument } from 'docx-svelte-viewer/runtime';

	let model = $state.raw(createDocument());
</script>

<WordEditor documentModel={model} ondocumentchange={(next) => (model = next)} />
```

Keep the model in `$state.raw`, not `$state`, so the editor recognises the model it emitted.

</details>

<details>
<summary><strong>Solid</strong> (<code>docx-solid-viewer</code>)</summary>

```tsx
import { createSignal } from 'solid-js';
import { createDocument, WordEditor } from 'docx-solid-viewer';

export function Editor() {
	const [model, setModel] = createSignal(createDocument());
	return <WordEditor class="editor" documentModel={model()} onDocumentChange={setModel} />;
}
```

</details>

<details>
<summary><strong>Plain JavaScript</strong> (<code>docx-vanilla-viewer</code>)</summary>

```ts
import { createDocument, mountEditor } from 'docx-vanilla-viewer';

const editor = mountEditor(container, {
	documentModel: createDocument(),
	onDocumentChange: (model) => console.log(model),
	onDocumentError: (error) => console.error(error),
});
await editor.load(bytes); // .docx or legacy .doc
const saved = await editor.save(); // a .docx Blob
editor.destroy();
```

</details>

Every package's handle has `element`, `load(bytes)`, `save()`, `download(fileName?)`, `markClean()`, `openCloudFile(file)` and `dirty`. The [bindings guide](docs/bindings.md) has every prop and event, and each package's README its framework details.

## Packages

| Package                                                                    | What it is                                                                              |
| -------------------------------------------------------------------------- | --------------------------------------------------------------------------------------- |
| [`docx-react-viewer`](https://www.npmjs.com/package/docx-react-viewer)     | React component (`packages/react`).                                                     |
| [`docx-vue-viewer`](https://www.npmjs.com/package/docx-vue-viewer)         | Vue 3 component (`packages/vue`).                                                       |
| [`docx-angular-viewer`](https://www.npmjs.com/package/docx-angular-viewer) | Angular standalone component (`packages/angular`).                                      |
| [`docx-svelte-viewer`](https://www.npmjs.com/package/docx-svelte-viewer)   | Svelte 5 component (`packages/svelte`).                                                 |
| [`docx-solid-viewer`](https://www.npmjs.com/package/docx-solid-viewer)     | Solid component, with server rendering (`packages/solid`).                              |
| [`docx-vanilla-viewer`](https://www.npmjs.com/package/docx-vanilla-viewer) | `mountEditor` and the plain `<docx-editor>` element, no framework (`packages/vanilla`). |
| [`docx-core`](https://www.npmjs.com/package/docx-core)                     | The Word model, parser and preserving serializer for headless use (`packages/core`).    |
| [`docx-viewer-mcp`](https://www.npmjs.com/package/docx-viewer-mcp)         | An MCP server with Word tools for AI assistants (`mcp`).                                |

Each editor package is self-contained: it depends on `docx-core`, `ooxml-core`, `ooxml-ui` (the shared elements, which also bring the ProseMirror libraries) and its framework peer, and nothing else. Document loading, layout and the legacy `.doc` reader come from `ooxml-core` (`/docx/load`, `/docx/layout`), so the legacy binary codecs need no separate install. The `<docx-editor>` element lives in `ooxml-ui/docx`; the private `web-component` and `bindings` packages here are bundled into each editor package at build time.

## SmartArt

A SmartArt graphic is shown from the drawing the producing application saved in the file, drawn as SVG by `<office-ui-smartart>`. Until it is edited that drawing is a snapshot: its layout is not recomputed and is not claimed to look like Word. Preset outlines the renderer cannot draw appear as rectangles, gradient and pattern fills are approximated, transparency and 3D are dropped, and text is one line per shape without autofit; the diagram's info popover lists each case. An unedited diagram keeps its parts and markup byte for byte on save.

Selecting a SmartArt in the body shows the **SmartArt Design** tab: Add Shape, Promote, Demote, Move Up and Move Down, the shared **Text Pane**, **Change Layout**, and the SmartArt Styles and Change Colors galleries. **Insert > SmartArt** opens the shared picker over 35 built-in layouts. Edits go through the core and are undoable. An edited, inserted or re-laid-out diagram is drawn with an approximate layout computed from its nodes, not Word's; on save its data part is rewritten and the stale cached drawing is dropped, because Word lays the diagram out again when it opens the file. Picture layouts' pictures, per-shape formatting and SmartArt in headers, footers and notes are not editable.

## Charts

Charts are drawn from the values cached in their part: column, bar, line, area, pie, doughnut, scatter, bubble, stock, radar and their combinations, and the waterfall, histogram, treemap and sunburst families of the newer chart format. Surface, volume-stock and the other newer families show a labelled frame.

**Insert > Chart** opens the shared chart picker with the column, bar, line, area, pie, doughnut, scatter and radar types and the newer families, and inserts an inline chart with Office's sample data. A selected chart shows the **Chart Design** tab: **Legend**, **Edit Data** (the shared data sheet), **Change Chart Type** and **Format Selection**, which docks the shared Format pane beside the page. Double-clicking a chart opens Edit Data, the corner handle resizes it (Shift keeps its aspect ratio), and each edit is one undo step.

Saving writes an edited chart part with the shared chart writer and updates only the cells of its embedded workbook that the edit changed; an unedited chart keeps its parts byte for byte. Not offered yet: chart styles, colours and Quick Layout, Add Chart Element, Switch Row/Column, Select Data, and the Format pane's picture section. Floating charts are not placed on Print Layout pages, and only the bottom-right handle resizes.

## Privacy and network use

Screenshot and Screen Clipping use the browser’s surface picker and embed the result locally. See the [capture privacy details](../../PRIVACY.md#screenshots-and-screen-recordings).

3D Models and Object read the file you choose and embed it in your local document. In an active Word collaboration session, the selected file and its preview are shared with that session's peers through its configured connection, like inserted pictures. For local-only editing, leave the session.

By default the Word packages make no network request when opening, rendering, editing or saving a document: linked pictures are not fetched, linked data is not refreshed, and fonts come from the device and the file. Missing fonts are requested from Google Fonts, or a mirror the host chooses, only when the host sets the element's `googleFonts` property or the reader turns on font downloads in File > Options > Trust Center, which is off by default. Hyperlinks open only when the reader follows them. Small preferences, including that Trust Center choice, stay in the browser's localStorage. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Supported environments

Browsers: the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)). Node.js 22.3 or newer for builds, tests and `docx-core`.

## Development

Run these from `viewers/docx`:

```bash
bun run typecheck         # tsc and svelte-check
bun run test              # Vitest: unit and binding contract tests
bun run check:shared      # docx-core stays a logic-free re-export
bun run build:packages    # build the publishable packages into packages/*/dist
bun run check:published   # no tarball imports an internal package or ole2
bun run pack:smoke        # pack, install into a clean consumer, load DOCX and .doc
bunx playwright install chromium
bun run test:browser      # Playwright against the built demo (PLAYWRIGHT_PORT, default 4180)
bun run fmt:check         # oxfmt
bun install --cwd docs && bun run --cwd docs docs:build   # the docs site and every demo
```

`PLAYWRIGHT_CHROMIUM_EXECUTABLE` selects an explicit Chromium. The packages are released by the repository's shared release flow; see [releasing](docs/releasing.md).

## Documentation

[Getting started](docs/getting-started.md) &middot; [Architecture](docs/architecture.md) &middot; [Framework bindings](docs/bindings.md) &middot; [Building blocks](docs/building-blocks.md) &middot; [Editing text](docs/editing.md) &middot; [Theming](docs/theming.md) &middot; [Collaboration](docs/collaboration.md) &middot; [OneDrive](docs/onedrive.md) &middot; [Support roadmap](docs/parity-roadmap.md) &middot; [Preservation guarantee](docs/preservation.md) &middot; [Outstanding work](docs/outstanding-work.md)

## License

[Apache License 2.0](LICENSE). See [`NOTICE`](NOTICE) for attributions.
