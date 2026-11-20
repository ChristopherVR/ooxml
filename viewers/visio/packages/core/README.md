# visio-core

Read, inspect and edit Visio drawings in Node or the browser, with no DOM. The document API behind every Visio viewer package.

[![npm version](https://img.shields.io/npm/v/visio-core.svg)](https://www.npmjs.com/package/visio-core)
[![license](https://img.shields.io/npm/l/visio-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE)
[![types](https://img.shields.io/npm/types/visio-core.svg)](https://www.npmjs.com/package/visio-core)
[![status: beta](https://img.shields.io/badge/status-beta-orange.svg)](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)

[Live demo](https://christophervr.github.io/ooxml/visio/demo/) · [Docs](https://christophervr.github.io/ooxml/visio/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/core/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

> **Beta.** The API changes only with a changelog note. Edits are validated and bounded, but saving is not yet guaranteed to keep everything a file contains, and rendering does not match Microsoft Visio. The [capability ledger](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/docs/parity.md) lists what works and what is missing.

`visio-core` re-exports `ooxml-core/visio`. It adds no parser or editor of its own, so it always matches the `ooxml-core` it installs.

## Install

```bash
npm install visio-core
```

## Quick start

```js
import { readFile, writeFile } from 'node:fs/promises';
import { editVsdx, parseVsdx } from 'visio-core';

const bytes = await readFile('flow.vsdx');
const drawing = await parseVsdx(bytes);

for (const page of drawing.pages) console.log(page.name, page.shapes.length, 'shapes');
for (const note of drawing.diagnostics) console.warn(note.code, note.message);

const page = drawing.pages[0];
const result = await editVsdx(bytes, [
	{ type: 'replace-plain-text', pageId: page.id, shapeId: page.shapes[0].id, text: 'Approved' },
]);
await writeFile('flow-approved.vsdx', result.bytes);
```

`parseVsdx` and `editVsdx` take a `Uint8Array` or `ArrayBuffer`. `loadVisio` detects the format and also reads legacy binary `.vsd` files, read-only.

## What you get

- **A document model.** Pages, background pages, shapes, groups and connectors with master inheritance, geometry, text, styles, themes, layers, shape data, hyperlinks, comments and embedded pictures, plus diagnostics for what was approximated or left out.
- **Validated edits.** `editVsdx` applies a batch of edits atomically and returns `{ bytes, changedParts, diagnostics }`: text, formatting, create, move, resize, rotate and delete, ordering, pages, page setup, shape data, layers, containers, connectors and charts. Unsupported or unsafe edits are refused, never half applied.
- **New drawings.** `createVsdx` writes a blank drawing.
- **Formats.** `.vsdx` and its macro-enabled and template kinds (`.vsdm`, `.vstx`, `.vstm`); `readVisioStencil` reads stencils; `parseLegacyVsd` previews Visio 2003 `.vsd`. An edit keeps a VBA project byte for byte and never runs it.
- **Bounded by default.** Every read and edit has a time limit and one memory budget sized from the host's heap. Pages of up to about 10,000 shapes as Visio saves them fit the default limits; a file that would exhaust memory is refused with `LIMIT_XML_MEMORY`, and `parseVsdx` reports `edit-size-limit` first when a drawing is too large to edit.

**Not supported yet:** general ShapeSheet recalculation (only bounded numeric dependencies of admitted edits are recomputed); most EMF records, EMF+, WMF and OLE objects; VDX files; older binary Visio versions. Plain-text editing and the shape clipboard are experimental.

## Limits

Pass `limits` to tighten what one read or edit may use:

```js
import { parseVsdx, VisioPackageError } from 'visio-core';

try {
	await parseVsdx(bytes, { limits: { maxXmlHeapBytes: 256 * 1024 * 1024 } });
} catch (error) {
	if (error instanceof VisioPackageError) console.error(error.code, error.message);
	else throw error;
}
```

By default `maxXmlHeapBytes` follows the host's heap (the limit Node or a Chromium page reports, less 256 MiB, divided by 1.6; 480 MiB where the runtime reports none). Set it lower when the same heap holds other large data or runs several reads at once. A page of 10,000 shapes needs a Node heap limit of about 1,000 MiB (`--max-old-space-size=820` or more).

## Supported environments

- Node.js 22.3 or newer.
- Browsers and web workers: the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).

## Privacy and network use

The Visio packages make no network request when opening, editing or saving a drawing: linked pictures are not fetched and linked data is not refreshed. This package never downloads fonts; the viewers do so only when the host or the reader turns Google Fonts on. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

| Package                                                                                                                                                                                                                                                                                                                                                                                                                                                          | What it is                                                       |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| [`visio-react-viewer`](https://www.npmjs.com/package/visio-react-viewer), [`visio-vue-viewer`](https://www.npmjs.com/package/visio-vue-viewer), [`visio-angular-viewer`](https://www.npmjs.com/package/visio-angular-viewer), [`visio-svelte-viewer`](https://www.npmjs.com/package/visio-svelte-viewer), [`visio-solid-viewer`](https://www.npmjs.com/package/visio-solid-viewer), [`visio-vanilla-viewer`](https://www.npmjs.com/package/visio-vanilla-viewer) | The browser viewer for each framework; each re-exports this API. |
| [`visio-viewer-mcp`](https://www.npmjs.com/package/visio-viewer-mcp)                                                                                                                                                                                                                                                                                                                                                                                             | Visio inspection and edits as MCP tools for agents.              |
| [`ooxml-core`](https://www.npmjs.com/package/ooxml-core)                                                                                                                                                                                                                                                                                                                                                                                                         | The Office logic this package re-exports.                        |

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme), with viewers for [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme), [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme), [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) and [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE). See [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/NOTICE) for attributions.
