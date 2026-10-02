<div align="center">

# ooxml-core

**Read, edit, validate and write Office Open XML documents in TypeScript.**
One package, one XML model, every format: Word, PowerPoint and the shared building blocks beneath them. Spreadsheets, Visio and more are planned.

[![license](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](CODE_OF_CONDUCT.md)

[**Try the apps**](https://christophervr.github.io/ooxml/) &nbsp;&middot;&nbsp;
[**Packages and areas**](#one-package-many-areas) &nbsp;&middot;&nbsp;
[**Install**](#install) &nbsp;&middot;&nbsp;
[**Examples**](#examples) &nbsp;&middot;&nbsp;
[**Roadmap**](#roadmap) &nbsp;&middot;&nbsp;
[**Contributing**](CONTRIBUTING.md)

</div>

## Why ooxml-core?

- **No UI, no framework.** Everything runs in browsers, Node.js, Bun, workers and serverless functions. The viewer apps ([docx-viewer](https://github.com/ChristopherVR/docx-viewer), [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer)) are thin interfaces on top of it, and so can your application be.
- **One package, one structure.** All formats share the same XML model, packaging layer, units, colours and geometry. Each format is an _area_ of this one package, not a separate dependency to keep in step.
- **Round-trips without losing what it does not understand.** Documents are loaded into a model, edited, and written back. Untouched parts stay byte-for-byte, unknown markup is preserved, and edits that would damage unsupported content are rejected instead of silently dropped.
- **Strict by default.** New code is strict TypeScript with branded measurement units, the shared `xml` area parses strictly (no DTD or entity expansion), and the `docx` area is checked against the ECMA-376 schemas in the test suite.
- **Honest about its limits.** Unsupported features are reported, never hidden, and nothing here claims Office parity or lossless export without evidence.

## One package, many areas

`ooxml-core` is a **single published package**. Every area is a subpath import, so you only load what you use, and the formats are symmetrical: `docx` and `pptx` are each imported through their own subpaths, and the root entry groups only the shared building blocks by namespace.

```ts
import { parseXml } from 'ooxml-core/xml'; // one area
import { xml, opc } from 'ooxml-core'; // or by namespace (shared building blocks)
```

| Area       | What it is                                                                                                                                                                                                                                                          |
| ---------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `xml`      | The shared XML model: strict DOM parsing and serialization, namespaces, namespace-aware helpers.                                                                                                                                                                    |
| `opc`      | Open Packaging Conventions: relationships, content types, part paths, zip helpers, safe hyperlinks.                                                                                                                                                                 |
| `units`    | Branded EMU, twip and point types, constants and conversions.                                                                                                                                                                                                       |
| `color`    | Hex, RGB, HSL and linear colour primitives, OOXML percent and angle parsing.                                                                                                                                                                                        |
| `geometry` | DrawingML preset shapes, connection sites, clip paths, callouts and boolean shape operations.                                                                                                                                                                       |
| `diagram`  | SmartArt (DiagramML), format-neutral: data model, layout/colour/quick-style parts, the cached `dsp:drawing` shape tree, relationship resolution and a loader. Used by `docx`; `pptx` re-imports its parsers.                                                        |
| `collab`   | Format-neutral real-time collaboration on Yjs: session and provider lifecycle, awareness/presence, transport-neutral sync (WebSocket, in-memory, any byte channel), update codecs, ordering helpers, asset sync and the product adapter seam. See `docs/collab-area.md`. |
| `docx`     | WordprocessingML: model, parser, preserving serializer, editing, validation. Also `/docx/embedded`. `/docx/layout` is the DOM-free pagination engine (an approximation of Word, not parity) and `/docx/load` detects and loads DOCX and legacy .doc (ole2 inlined). |
| `pptx`     | PresentationML: model, parser, serializer, editing, charts, SmartArt, converters, CLI, signatures. Subpaths `/pptx/converter`, `/pptx/cli`, `/pptx/signature-node`.                                                                                                 |

Legacy binary formats (`.doc`, `.xls`, `.ppt`) and the compound-file container live in the sibling package [`ole2`](https://github.com/ChristopherVR/ole2); this package never contains binary codecs, and `ole2` never contains modern OOXML.

## Install

```bash
npm install ooxml-core
```

Optional peer dependencies enable specific features: `node-forge` and `xml-crypto` for digital signatures (`/pptx/signature-node`), and `@napi-rs/canvas` for server-side rasterisation.

## Examples

**Open a Word document, change it, save it**

```ts
import { loadDocx } from 'ooxml-core/docx';

const loaded = await loadDocx(bytes); // Uint8Array | ArrayBuffer
loaded.model.blocks; // paragraphs and tables
const edited = await loaded.save(); // original package parts are preserved
```

**Build or edit a PowerPoint deck**

```ts
import { PptxHandler } from 'ooxml-core/pptx';

const { handler, data, createSlide } = await PptxHandler.create({ title: 'Quarterly Review' });
data.slides.push(
	createSlide().addText('Hello World', { x: 100, y: 100, width: 600, height: 80 }).build(),
);
const bytes = await handler.save(data.slides); // a valid .pptx
```

**Work at the package level**

```ts
import { parseXml } from 'ooxml-core/xml';
import { parseRelationships, resolvePartPath } from 'ooxml-core/opc';

const rels = parseRelationships(relsXml);
const part = resolvePartPath('word/document.xml', rels.get('rId5')!.target);
```

## Roadmap

The shared areas come first, then more formats on the same foundation:

- **Shared layers:** DrawingML (fills, lines, effects, text, theme), charts, diagrams (SmartArt), maths, encryption primitives and schema-generated types, each as an area, written once for every format.
- **More formats:** `xlsx` (SpreadsheetML), Visio (`.vsdx`, also an OPC package), and further Office Open XML parts as they are needed. Each arrives as its own area.
- **One XML model:** the `pptx` area still uses its own XML object model and is compiled with relaxed TypeScript flags while it is migrated onto the shared `xml` area and tightened. New code is strict.
- **Collaboration:** the Yjs and sync protocol that the viewers share will live here as a `collab` area.

## Development

You need [Bun](https://bun.sh/) and Node.js 22 or newer.

```bash
bun install
bun run typecheck      # strict project and the pptx project
bun run test
bun run build
bun run test:package   # packs the build and imports every entry point from a clean install
```

The working agreements are in [AGENTS.md](AGENTS.md), and [PROVENANCE.md](PROVENANCE.md) records where each module came from.

## Related projects

- [docx-viewer](https://github.com/ChristopherVR/docx-viewer) and [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer): the editors and viewers built on this package.
- [OOXML Office](https://christophervr.github.io/ooxml/): the suite's launcher page (`site/`), which opens the demos those viewers deploy to their own GitHub Pages sites.
- [ole2](https://github.com/ChristopherVR/ole2): the compound-file container and legacy binary Office codecs.

## License

[Apache-2.0](LICENSE). Third-party notices are in [NOTICE](NOTICE) and [THIRD-PARTY-LICENSES](THIRD-PARTY-LICENSES).
