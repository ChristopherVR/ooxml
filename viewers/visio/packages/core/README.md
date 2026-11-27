# visio-core

**The DOM-free Visio document API: read, inspect, create and edit `.vsdx` drawings in Node.js, browsers and workers.**

[![npm version](https://img.shields.io/npm/v/visio-core.svg)](https://www.npmjs.com/package/visio-core)
[![license](https://img.shields.io/npm/l/visio-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE)
[![types](https://img.shields.io/npm/types/visio-core.svg)](https://www.npmjs.com/package/visio-core)

[Live demo](https://christophervr.github.io/ooxml/visio/demo/) · [Documentation](https://christophervr.github.io/ooxml/visio/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/core/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

`visio-core` re-exports `ooxml-core/visio`, the Visio area of [`ooxml-core`](https://www.npmjs.com/package/ooxml-core). It adds no parser or editor of its own, so it always matches the `ooxml-core` it installs, and it is the API every Visio viewer package and the MCP server use.

## Install

```bash
npm install visio-core
```

`ooxml-core` is installed with it. There are no peer dependencies.

## Quick start

```js
import { readFile, writeFile } from 'node:fs/promises';
import { createVsdx, editVsdx, parseVsdx } from 'visio-core';

// Read: pages, shapes and diagnostics for what was approximated or left out.
const bytes = await readFile('flow.vsdx');
const drawing = await parseVsdx(bytes);
for (const page of drawing.pages) console.log(page.name, page.shapes.length, 'shapes');
for (const note of drawing.diagnostics) console.warn(note.code, note.message);

// Edit: one atomic batch; a refused edit changes nothing.
const page = drawing.pages[0];
const result = await editVsdx(bytes, [
	{ type: 'replace-plain-text', pageId: page.id, shapeId: page.shapes[0].id, text: 'Approved' },
]);
await writeFile('flow-approved.vsdx', result.bytes);

// Create: a blank landscape Letter page (dimensions are inches; the default is portrait Letter).
await writeFile('blank.vsdx', await createVsdx({ width: 11, height: 8.5 }));
```

`parseVsdx` and `editVsdx` take a `Uint8Array` or `ArrayBuffer`. Coordinates are drawing inches with the origin at the bottom left of the page. `editVsdx` returns `{ bytes, changedParts, diagnostics }`.

## API overview

| Entry point                                                                     | What it does                                                                                                                                                                  |
| ------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `parseVsdx(bytes, options?)`                                                    | Reads `.vsdx`, `.vsdm`, `.vstx` and `.vstm` into a typed scene: pages, shapes, groups, connectors, master inheritance, geometry, text, styles, themes, layers and shape data. |
| `loadVisio(bytes, options?)`                                                    | Detects the format; legacy binary `.vsd` becomes a read-only preview (`parseLegacyVsd`).                                                                                      |
| `editVsdx(bytes, edits, options?)`                                              | Applies a batch of `VisioEdit` commands atomically: text, formatting, shapes, connectors, containers, pages, layers, shape data, data and charts.                             |
| `createVsdx(options?)`                                                          | Writes a blank, editable one-page drawing.                                                                                                                                    |
| `readVisioStencil(bytes)`, `loadVisioStencilEditor()`                           | Reads `.vssx` and `.vssm` stencils; lazily loads the stencil authoring API.                                                                                                   |
| `convertVsdxToTemplate(bytes)`                                                  | Converts a drawing (`.vsdx`, `.vsdm`) into a template (`.vstx`, `.vstm`).                                                                                                     |
| `parseVsdxPages`, `mergeVisioPages`                                             | Re-reads only the pages that changed.                                                                                                                                         |
| `captureVisioClipboard`, `serializeVisioClipboard`, `deserializeVisioClipboard` | Bounded, versioned shape clipboard text (experimental; never touches the system clipboard).                                                                                   |
| `parseVisioFormula`, `evaluateVisioFormula`                                     | The bounded ShapeSheet formula interpreter (an AST, never JavaScript evaluation).                                                                                             |
| `visioTextExtents`, `setVisioTextMeasurer`                                      | Text measuring used to keep text-sized shapes in step; a UI can supply its own measurer.                                                                                      |
| `VisioPackageError`                                                             | Thrown for unsafe, malformed or over-limit packages, with a `code` such as `LIMIT_XML_MEMORY`.                                                                                |

The full edit catalogue, limits and module map are in the [area overview](https://github.com/ChristopherVR/ooxml/blob/main/src/core/visio/README.md); the viewer-facing contract is the [viewer API](https://christophervr.github.io/ooxml/visio/api).

### Limits

Every read and edit has a ten-second cooperative deadline and one memory budget, `maxXmlHeapBytes`, which by default follows the host's heap (480 MiB where the runtime reports none). Pass `limits` to tighten it:

```js
import { parseVsdx, VisioPackageError } from 'visio-core';

try {
	await parseVsdx(bytes, { limits: { maxXmlHeapBytes: 256 * 1024 * 1024 } });
} catch (error) {
	if (error instanceof VisioPackageError) console.error(error.code, error.message);
	else throw error;
}
```

Pages of up to about 10,000 shapes as Visio saves them fit the defaults; such a page needs a Node heap limit of about 1,000 MiB (`--max-old-space-size=820` or more). When a drawing is too large to edit, `parseVsdx` reports an `edit-size-limit` diagnostic first.

## Status and limits

**Beta.** The API changes only with a changelog note.

- Untouched package parts are kept byte for byte and edited parts semantically, but there is no written preservation guarantee yet; signatures are removed on edit (as Visio does) and a VBA project is carried, never run.
- Each edit command has its own admission rules; anything outside them is refused with a diagnostic, never half applied.
- There is no general ShapeSheet recalculation: only the bounded dependencies of admitted edits are recomputed.
- Plain-text editing and the shape clipboard are experimental.
- Most EMF records, EMF+, WMF and OLE objects, VDX files and older binary Visio versions are not supported.

See the [capability ledger](https://christophervr.github.io/ooxml/visio/parity) and the [parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md).

## Privacy and network use

The package makes no network request: it reads and writes the bytes you pass. Linked pictures are not fetched, data record sets and linked chart data are never refreshed, macros never run, and fonts are never downloaded (only the viewers can, when the host or reader turns Google Fonts on). Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#word-excel-and-visio).

## Environments

- Node.js 22.3 or newer.
- Browsers and web workers: the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)). Run large reads and edits in a worker to keep a page responsive.

## Documentation

- [Visio documentation](https://christophervr.github.io/ooxml/visio/) and the [viewer API](https://christophervr.github.io/ooxml/visio/api).
- [Area overview](https://github.com/ChristopherVR/ooxml/blob/main/src/core/visio/README.md) of `ooxml-core/visio`.
- Browser viewers that re-export this API: [`visio-react-viewer`](https://www.npmjs.com/package/visio-react-viewer), [`visio-vue-viewer`](https://www.npmjs.com/package/visio-vue-viewer), [`visio-angular-viewer`](https://www.npmjs.com/package/visio-angular-viewer), [`visio-svelte-viewer`](https://www.npmjs.com/package/visio-svelte-viewer), [`visio-solid-viewer`](https://www.npmjs.com/package/visio-solid-viewer) and [`visio-vanilla-viewer`](https://www.npmjs.com/package/visio-vanilla-viewer); MCP tools: [`visio-viewer-mcp`](https://www.npmjs.com/package/visio-viewer-mcp).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE). See [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/NOTICE) for attributions.
