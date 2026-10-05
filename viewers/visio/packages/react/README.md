# visio-react-viewer

[![npm version](https://img.shields.io/npm/v/visio-react-viewer.svg)](https://www.npmjs.com/package/visio-react-viewer)
[![license](https://img.shields.io/npm/l/visio-react-viewer.svg)](https://github.com/ChristopherVR/visio-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/visio-react-viewer.svg)](https://www.npmjs.com/package/visio-react-viewer)

> A browser Visio `.vsdx` viewer for React 18 or later. This package ships the functional viewer and its framework adapter.

[Live demo](https://christophervr.github.io/visio-viewer/demo/) | [npm](https://www.npmjs.com/package/visio-react-viewer) | [Full docs](https://christophervr.github.io/visio-viewer/) | [Source](https://github.com/ChristopherVR/visio-viewer)

## Install

```bash
npm install visio-react-viewer
```

## Quick start

```tsx
import { useRef } from 'react';
import { VisioViewer, type ViewerHandle } from 'visio-react-viewer';

export function Diagram() {
	const viewer = useRef<ViewerHandle>(null);
	return (
		<>
			<input
				type="file"
				accept=".vsdx"
				onChange={async (event) => {
					const file = event.currentTarget.files?.[0];
					if (file) await viewer.current?.load(file);
				}}
			/>
			<VisioViewer ref={viewer} showToolbar style={{ height: 600 }} />
		</>
	);
}
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

## License

Apache-2.0.
