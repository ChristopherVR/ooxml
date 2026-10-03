# visio-vanilla-viewer

[![npm version](https://img.shields.io/npm/v/visio-vanilla-viewer.svg)](https://www.npmjs.com/package/visio-vanilla-viewer)
[![license](https://img.shields.io/npm/l/visio-vanilla-viewer.svg)](https://github.com/ChristopherVR/visio-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/visio-vanilla-viewer.svg)](https://www.npmjs.com/package/visio-vanilla-viewer)

> A browser Visio `.vsdx` viewer for a browser with ES modules and a bundler such as Vite. This package ships the functional viewer and its framework adapter.

[Live demo](https://christophervr.github.io/visio-viewer/demo/) | [npm](https://www.npmjs.com/package/visio-vanilla-viewer) | [Full docs](https://christophervr.github.io/visio-viewer/) | [Source](https://github.com/ChristopherVR/visio-viewer)

## Install

```bash
npm install visio-vanilla-viewer
```

## Quick start

```js
import { mountViewer } from 'visio-vanilla-viewer';

const host = document.createElement('div');
host.style.height = '600px';
const input = document.createElement('input');
input.type = 'file';
input.accept = '.vsdx';
document.body.append(input, host);

const viewer = mountViewer(host, { showToolbar: true });
input.addEventListener('change', async () => {
	const file = input.files?.[0];
	if (file) await viewer.load(file);
});
// Call viewer.destroy() when the host is removed.
```

## Features

| Feature    | Description                                                         |
| ---------- | ------------------------------------------------------------------- |
| Viewing    | One SVG renderer, page selection, zoom, layers and text search.     |
| Editing    | Core-validated text and geometry operations with bounded undo/redo. |
| Export     | Source-backed VSDX copies, static SVG and print snapshots.          |
| Frameworks | One shared `<visio-viewer>` UI behind six adapters.                 |

## API

The shared properties are `document` (a parsed `VisioDocument`, or `null` to clear),
`pageIndex` (zero-based), `zoom` (`1` means 100%), `showToolbar` and `events`.
The callback map uses `document-load`, `document-change`, `document-error`,
`page-change`, `zoom-change` and `shape-select`. File loading uses the imperative
`load(File | Blob | Uint8Array | ArrayBuffer)` handle; `document` is not a file URL.
Provide a height for the viewer host. Mount and load files in the browser.

The handle supports `fit()`, `replacePlainText(pageId, shapeId, text)`,
`applyEdits(edits)`, `undo()`, `redo()` and `exportVsdx()`.
`exportVsdx()` returns `{ bytes, dirty, diagnostics }`; saving or downloading those
bytes is the application's responsibility. Load original bytes through `load()`
to enable source-backed editing and export. Handles are usable after mount.

Every viewer package includes the same browser UI and re-exports the
`visio-core` document API. Parsing and editing belong to `ooxml-core/visio`;
framework bindings only manage lifecycle, properties and events.

## Limitations

Rendering and editing are beta features. Supported text and geometry edits are
bounded by core validation; unsupported operations reject and diagnostics report
limitations. Native Visio layout parity and lossless export are not established.
Files stay in the browser unless your application sends them elsewhere.

## Documentation

[Viewer guide](https://christophervr.github.io/visio-viewer/docs/) |
[Demo](https://christophervr.github.io/visio-viewer/demo/) |
[Source](https://github.com/ChristopherVR/visio-viewer)

For direct custom-element use, call `registerVisioViewer()` from this package,
then create a `<visio-viewer>` element and call its `load()` method. Registration
is explicit; importing the package does not register the element.

## License

Apache-2.0.
