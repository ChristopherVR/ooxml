# pptx-viewer-core

**The headless PowerPoint engine: read, build, edit, save and convert `.pptx` (and legacy `.ppt`) files from TypeScript in Node.js, workers or the browser, with no UI and no native dependencies.**

[![npm version](https://img.shields.io/npm/v/pptx-viewer-core.svg)](https://www.npmjs.com/package/pptx-viewer-core)
[![license](https://img.shields.io/npm/l/pptx-viewer-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)
[![types](https://img.shields.io/npm/types/pptx-viewer-core.svg)](https://www.npmjs.com/package/pptx-viewer-core)

[Documentation](https://christophervr.github.io/ooxml/pptx/core/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/core/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

It is a thin entry point over `ooxml-core/pptx` and the engine behind the [React](https://www.npmjs.com/package/pptx-react-viewer), [Vue](https://www.npmjs.com/package/pptx-vue-viewer), [Angular](https://www.npmjs.com/package/pptx-angular-viewer), [Svelte](https://www.npmjs.com/package/pptx-svelte-viewer) and [vanilla](https://www.npmjs.com/package/pptx-vanilla-viewer) viewers.

## Install

```bash
npm install pptx-viewer-core
```

Reading, editing, saving and password protection need nothing else (`ooxml-core` is its one dependency; encryption runs on the Web Crypto API). Node signature signing and verification through `pptx-viewer-core/signature-node` need the optional peers `node-forge` 1, `xml-crypto` 6 and `@xmldom/xmldom` 0.8 or 0.9.

## Quick start

```ts
import { readFile, writeFile } from 'node:fs/promises';

import { PptxHandler } from 'pptx-viewer-core';

const file = await readFile('deck.pptx');
const handler = new PptxHandler();
// load() takes an ArrayBuffer: .pptx, .pptm, .ppsx, .potx, legacy .ppt or pptx-viewer-json
const data = await handler.load(
	file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength),
);

console.log(`${data.slides.length} slides, ${data.width}x${data.height} px`);
for (const element of data.slides[0]?.elements ?? []) {
	if (element.type === 'text') console.log(element.text);
}

await writeFile('copy.pptx', await handler.save(data.slides)); // Uint8Array of a .pptx
```

For whole-deck edits, `Presentation` wraps the handler:

```ts
import { ChartBuilder, Presentation, ThemePresets } from 'pptx-viewer-core';

const pptx = await Presentation.create({ title: 'Sales Report', theme: ThemePresets.MODERN_BLUE });
pptx
	.addSlide('Title Slide')
	.addText('Q4 Sales Report', { fontSize: 44, bold: true, x: 100, y: 200, width: 800, height: 80 });
pptx
	.addSlide('Blank')
	.addBuilderElement(
		ChartBuilder.create('bar')
			.categories(['North', 'South', 'East', 'West'])
			.addSeries('2026', [210, 150, 180, 120], '#2563EB')
			.title('Revenue ($M)')
			.bounds(50, 100, 860, 420),
	);
pptx.replaceText('2026', 'FY2026');
const bytes = await pptx.save();
```

`Presentation.load(buffer)` opens an existing deck the same way. Positions and sizes in the model are pixels at 96 DPI; the exact EMU values read from the file are kept beside them (`xEmu`, `widthEmu` and so on) for the save.

## API overview

| Import                                                            | Main entry points                                                                                                                                                                                                                                                                                                                            |
| ----------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `pptx-viewer-core`                                                | `PptxHandler` (`load`, `save`, `saveEncrypted`), `Presentation`, the element builders (`TextBuilder`, `ShapeBuilder`, `ChartBuilder`, `TableBuilder`, `ImageBuilder`, `ConnectorBuilder`, `MediaBuilder`, `GroupBuilder`), `ThemePresets`, `SlideSizes`, unit helpers (`inches`, `cm`, `mm`, `pt`), the converters and the `PptxData` types. |
| `pptx-viewer-core/converter`                                      | `PptxMarkdownConverter` and `SvgExporter` on their own.                                                                                                                                                                                                                                                                                      |
| `pptx-viewer-core/signature-node`                                 | Node signature inspection (`inspectPptxDigitalSignatures`) and signing.                                                                                                                                                                                                                                                                      |
| `pptx-viewer-core/chart`, `/geometry`, `/color`, `/text`, `/math` | The shared chart, geometry, colour, text and maths helpers.                                                                                                                                                                                                                                                                                  |
| `pptx-viewer-core/smartart-layouts`                               | The built-in SmartArt layout definitions.                                                                                                                                                                                                                                                                                                    |

Converters: `new PptxMarkdownConverter(outputDir, options, fileSystem?).convert(data)` writes Markdown (pass a file system adapter to extract media), `PptxJsonConverter.toJson(data)` writes the self-contained `pptx-viewer-json` model that `load()` reads back, and `SvgExporter` draws slides as SVG. The package also has a `pptx` command (`info`, `export-svg`, `export-md`, `merge`, `find`, `replace`, `create`, `diff`); run it by package name, `npx pptx-viewer-core info deck.pptx`, since `npx pptx` fetches an unrelated package. See the [builder](https://christophervr.github.io/ooxml/pptx/core/builder), [converter](https://christophervr.github.io/ooxml/pptx/core/converter), [SVG export](https://christophervr.github.io/ooxml/pptx/core/svg-export) and [CLI](https://christophervr.github.io/ooxml/pptx/core/cli) pages.

## Status and limits

PowerPoint carries no stability label yet: it has not been assessed against the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md).

- Untouched parts are carried through a save, but saving an edited slide can still rewrite minor markup; it is not a lossless round trip.
- SmartArt without its cached drawing is laid out by an approximate engine, and some constructs load as approximations; both are reported on `data.warnings`.
- For decks you do not trust, `load(buffer, { limits })` bounds the input, uncompressed bytes, entries and slides.
- Signature verification reports `verified-trusted` only for a certificate that chains to a root you configure (`PPTX_VIEWER_TRUST_ROOTS_PEM` or `PPTX_VIEWER_TRUST_ROOTS_FILE`); no roots are built in.

See the [limitations page](https://christophervr.github.io/ooxml/pptx/guide/limitations) and the [Office parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md#powerpoint).

## Privacy and network use

Parsing, editing and saving happen locally and never upload the file. Nothing collects analytics or telemetry.

- Linked (external) pictures and media keep their addresses but are not fetched unless you pass `allowExternalImages: true` to `load`.
- Node signature inspection sends an OCSP request to the responder named in each certificate that verified a signature value; pass `{ checkRevocation: false }` to make no request.

See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#digital-signature-inspection-node-only).

## Environments

- Node.js 22.3 or newer, Bun, Deno, web workers and serverless runtimes; ESM and CommonJS builds.
- Browsers: the last two major versions of Chrome, Edge, Firefox and Safari, Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- `pptx-viewer-core/signature-node` is Node only.

## Documentation

- The [core guide](https://christophervr.github.io/ooxml/pptx/core/): [loading](https://christophervr.github.io/ooxml/pptx/core/loading), [editing](https://christophervr.github.io/ooxml/pptx/core/editing), [saving](https://christophervr.github.io/ooxml/pptx/core/saving) and [encryption and signatures](https://christophervr.github.io/ooxml/pptx/core/encryption).
- [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp) exposes this engine as MCP tools; [`ooxml-core`](https://www.npmjs.com/package/ooxml-core) holds the Word, Excel, PowerPoint and Visio logic it re-exports.
- Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE). Keep the [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/NOTICE) file with redistributions.
