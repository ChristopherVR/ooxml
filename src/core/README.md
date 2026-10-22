<div align="center">

# ooxml-core

[![npm version](https://img.shields.io/npm/v/ooxml-core.svg)](https://www.npmjs.com/package/ooxml-core)
[![license](https://img.shields.io/npm/l/ooxml-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/ooxml-core.svg)](https://www.npmjs.com/package/ooxml-core)

**Read, edit, validate and write Office Open XML documents in TypeScript.**
One package, one XML model, every format: Word, PowerPoint, Excel and the shared building blocks beneath them. Visio VSDX parsing and bounded editing are available through `ooxml-core/visio`.

[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](../../CODE_OF_CONDUCT.md)

[**Try the apps**](https://christophervr.github.io/ooxml/) &nbsp;&middot;&nbsp;
[**Packages and areas**](#features-and-api-one-package-many-areas) &nbsp;&middot;&nbsp;
[**Install**](#install) &nbsp;&middot;&nbsp;
[**Examples**](#quick-start) &nbsp;&middot;&nbsp;
[**Contributing**](../../CONTRIBUTING.md)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme) &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; [OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme) &middot; **[Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)**

</div>

## Why ooxml-core?

- **No UI, no framework.** Everything runs in browsers, Node.js, Bun, workers and serverless functions. The viewer apps ([docx](https://christophervr.github.io/ooxml/docx/), [xlsx](https://christophervr.github.io/ooxml/xlsx/), [visio](https://christophervr.github.io/ooxml/visio/), [teams](https://christophervr.github.io/ooxml/teams/) and [pptx](https://christophervr.github.io/ooxml/pptx/)) are thin interfaces on top of it, and so can your application be.
- **One package, one structure.** All formats share the same XML model, packaging layer, units, colours and geometry. Each format is an _area_ of this one package, not a separate dependency to keep in step.
- **Round-trips without losing what it does not understand.** Documents are loaded into a model, edited, and written back. Untouched parts stay byte-for-byte, unknown markup is preserved, and edits that would damage unsupported content are rejected instead of silently dropped.
- **Strict by default.** New code is strict TypeScript with branded measurement units, the shared `xml` area parses strictly (no DTD or entity expansion), and the `docx` area is checked against the ECMA-376 schemas in the test suite.
- **Honest about its limits.** Unsupported features are reported, never hidden, and nothing here claims Office parity or lossless export without evidence.

## Features and API: one package, many areas

`ooxml-core` is a **single published package**. Every area is a subpath import, so you only load what you use, and the formats are symmetrical: `docx`, `pptx` and `xlsx` are each imported through their own subpaths, and the root entry groups only the shared building blocks by namespace.

```ts
import { parseXml } from 'ooxml-core/xml'; // one area
import { xml, opc } from 'ooxml-core'; // or by namespace (shared building blocks)
```

| Area         | What it is                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| ------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `xml`        | The shared XML model: strict DOM parsing and serialization, namespaces, namespace-aware helpers.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| `opc`        | Open Packaging Conventions: relationships, content types, part paths, zip helpers, the hyperlink safety policy (`hyperlinkPolicy`) and document properties (core, app and custom, `opc/properties`) and digital signatures (`signature`: detection, the parts a writer strips, XML-DSig parsing and digest checks; the Node verifier is `/pptx/signature-node`).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| `units`      | Branded EMU, twip and point types, constants and conversions.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `color`      | Hex, RGB, HSL and linear colour primitives, OOXML percent and angle parsing.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `geometry`   | DrawingML preset shapes, connection sites, clip paths, callouts and boolean shape operations.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `diagram`    | SmartArt (DiagramML), format-neutral: data model, layout/colour/quick-style parts, the cached `dsp:drawing` shape tree, relationship resolution and a loader. Used by `docx`; `pptx` re-imports its parsers.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `drawingml`  | Format-neutral DrawingML: colour, fill, text-body and geometry readers, the colour and fill writers and the theme model that docx and xlsx read their themes through.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| `chart`      | Format-neutral chart algorithms and formatting helpers: data calculations, axes, number formats, the Change Colors palettes and the SVG chart painter.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `text`       | Shared text rules: script-category segmentation, tab leaders and styled-run wrapping.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| `i18n`       | Locale normalisation, dictionary merging and the translator every product UI uses.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `ribbon`     | DOM-free ribbon contracts every product shares, today the add-in tab descriptor a host uses to add its own tabs.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| `digest`     | Synchronous pure-TypeScript SHA-1, SHA-256, SHA-384 and SHA-512, ECMA-376 digest name normalisation (`SHA-512`, `SHA512`, `sha_512`, ...) and the agile password hash Office protection elements store (`hashPassword`, `verifyPasswordHash`). No Web Crypto, so it works on `http://`. Used by `xlsx` sheet and workbook protection.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| `crypto`     | ECMA-376 package encryption ([MS-OFFCRYPTO]) for every format: `decryptOoxmlPackage`, `encryptOoxmlPackage` (agile AES-256/SHA-512 by default, or Standard), `isEncryptedOoxmlPackage`, and typed errors with a `code` (`password-required`, `incorrect-password`, `data-integrity`). The compound-file container comes from ole2 (inlined); not in the root entry.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| `collab`     | Format-neutral real-time collaboration on Yjs: session and provider lifecycle, awareness/presence, transport-neutral sync (WebSocket, in-memory, any byte channel), update codecs, ordering helpers, asset sync and the product adapter seam. See `docs/collab-area.md`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `docx`       | WordprocessingML: model, parser, preserving serializer, editing, validation. Also `/docx/embedded`. `/docx/layout` is the DOM-free pagination engine (an approximation of Word, not parity) and `/docx/load` detects and loads DOCX (also password-protected, `{ password }`) and legacy .doc (ole2 inlined).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `pptx`       | PresentationML: model, parser, serializer, editing, charts, SmartArt, converters, CLI, signatures. Subpaths `/pptx/converter`, `/pptx/cli`, `/pptx/signature-node`, `/pptx/node`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| `xlsx`       | SpreadsheetML: workbook model, `.xlsx`/`.xlsm` reader and preserving writer, formula engine (dependency graph, dynamic arrays, 490+ functions), Excel number formats, editing commands with undo, and the DOM-free grid layout (sizes, colours, conditional formats, charts as SVG, drawing anchors). SmartArt is shown from the cached drawing (not re-laid out) and kept verbatim on save. `/xlsx/load` detects and loads `.xlsx`, `.xlsm` (also password-protected: `loadWorkbook(bytes, { password })`, `saveWorkbook(workbook, 'xlsx', { password })`), legacy `.xls` (ole2 inlined) and CSV, and refuses `.xlsb` with a typed error. `/xlsx/collab` shares a workbook over Yjs (cells, formats, rows, columns, merges, sheets and defined names; see `docs/collab-area.md`). Digital signatures are detected and reported, never kept on save. An approximation of Excel, not parity. |
| `visio`      | Visio: `.vsdx` parsing, loading and bounded editing, page layers and visibility, and a conservative preview of legacy `.vsd` (through ole2). See `docs/visio-legacy-vsd.md`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `math`       | Office Math: OMML to and from LaTeX and MathML, with colour and size helpers.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `teams`      | The logic of the OpenTeams workspace: channels and chat on Yjs, presence, WebRTC calls with pluggable signalling, server configuration and the `createTeamsClient` store every UI binds to. Built on `collab`. See `docs/teams-area.md`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `automation` | Headless document operations and filesystem execution shared by the MCP servers (`/automation`, `/automation/node`).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |

Legacy binary formats (`.doc`, `.xls`, `.ppt`) and the compound-file container live in the sibling package [`ole2`](https://github.com/ChristopherVR/ole2); this package never contains binary codecs, and `ole2` never contains modern OOXML.

The `chart` subpath provides shared chart data calculations (regression, quartiles,
blank values and stacked series), plus Office's 17 Change Colors palettes through
`CHART_COLOR_PALETTES` and `chartPaletteColors(palette, seriesCount, scheme)`.
Palette interpolation reuses the shared DrawingML color resolver. Native Excel
comparisons expose remaining one-channel-step rounding differences for some
themes and large series counts; this is not a claim of exact color parity.
The `text` subpath provides script-category
segmentation primitives. These areas are DOM-free and do not import a product model.
SVG curve flattening is available from `geometry`. See the
[PowerPoint reuse audit](../../docs/pptx-shared-logic-audit.md) for extraction evidence
and remaining candidates.

CSV import (`csvToWorkbook`, or `loadWorkbook(bytes, { csv })` from `ooxml-core/xlsx/load`)
types fields the way typing them into a cell does, which approximates Excel opening a CSV: a
number past 15 significant digits keeps the nearest double where Excel keeps the first 15
digits, and General columns read numeric dates month first. It takes Excel's own controls to
change that: the automatic data conversion switches (`conversions: { leadingZeros: false }`
keeps `00123`), and the Text Import Wizard's column formats
(`columns: { 0: 'text', 2: 'DMY', 5: 'skip' }`), decimal and thousands separators and trailing
minus. Each column whose data a conversion changed gets one entry in `workbook.warnings`, with a
count and the first cell.

## Install

```bash
npm install ooxml-core
```

Optional peer dependencies enable specific features: `node-forge` and `xml-crypto` for digital signatures (`/pptx/signature-node`), and `@napi-rs/canvas` for server-side rasterisation. The `.ppt` writer's Node rasteriser for SVG pictures lives in the Node-only `/pptx/node` entry, so `ooxml-core/pptx` never references `@napi-rs/canvas` and browser builds of it never see that package. A `.ppt` save rasterises an SVG picture that has no raster fallback with the DOM. In Node.js, where there is no DOM, pass a rasteriser (`save(slides, { outputFormat: 'ppt', rasterizeSvg: rasterizeSvgInNode })`, with `rasterizeSvgInNode` from `ooxml-core/pptx/node`); without one the picture is written as a placeholder with a `ppt-image-format-unsupported` warning. Earlier releases loaded `@napi-rs/canvas` from `ooxml-core/pptx` itself for this.

## Quick start

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
	createSlide('Title Slide')
		.addPlaceholder('ctrTitle', 'Quarterly Review')
		.addPlaceholder('subTitle', 'Results and next steps')
		.build(),
);
const bytes = await handler.save(data.slides); // a valid .pptx
```

`addPlaceholder(type, text, { idx? })` binds text to the selected layout and inherits
its position, size, and text styles. Use `{ idx: 1 }` or `{ idx: 2 }` for repeated
content slots in Two Content. Optional `x`, `y`, `width`, and `height` override the
inherited frame. If a template's master frame needs a compatible-role fallback,
only the newly created placeholder receives an explicit frame. A missing slot
requires a complete explicit frame; otherwise the builder throws. Standalone `new SlideBuilder(number, layoutPath, name, handler)`
requires a loaded handler for this method. Use `addText` for a free text box.

**Open a very large workbook**

```ts
import { DEFAULT_XLSX_LOAD_LIMITS, isWorkbookTooLargeError, loadXlsx } from 'ooxml-core/xlsx';

// The default limits guard against zip bombs; raise only what the expected files need.
const workbook = await loadXlsx(bytes, {
	limits: { maxTotalBytes: 4 * DEFAULT_XLSX_LOAD_LIMITS.maxTotalBytes, maxXmlElements: 20_000_000 },
});
// Over a limit, loading throws an error that isWorkbookTooLargeError(error) recognises (code
// 'workbook-too-large'): `limit` names the limit and `maximum` is its value. loadWorkbook
// ('ooxml-core/xlsx/load') takes `limits` too.
```

**Compare two versions of a workbook**

```ts
import { compareWorkbooks, comparisonText, loadXlsx } from 'ooxml-core/xlsx';

const result = compareWorkbooks(await loadXlsx(v1), await loadXlsx(v2));
console.log(comparisonText(result)); // values, formulas, formats, sheets, names
```

Cells are matched by address; see [workbook comparison](../../docs/xlsx-compare.md) for the categories and limits.

**Work at the package level**

```ts
import { parseXml } from 'ooxml-core/xml';
import { parseRelationships, resolvePartPath } from 'ooxml-core/opc';

const rels = parseRelationships(relsXml);
const part = resolvePartPath('word/document.xml', rels.get('rId5')!.target);
```

## Contributing and roadmap

`ooxml-core` is developed in the [ChristopherVR/ooxml](https://github.com/ChristopherVR/ooxml) repository, which also holds [`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui), the MCP servers and the Word, Excel, Visio and OpenTeams viewers. Its README covers the repository layout, the development workflow and releases, and the [roadmap](https://github.com/ChristopherVR/ooxml/blob/main/docs/roadmap.md) says what comes next; the working agreements are in [AGENTS.md](../../AGENTS.md) and [CONTRIBUTING.md](../../CONTRIBUTING.md), and [PROVENANCE.md](../../PROVENANCE.md) records where each module came from.

## License

[Apache-2.0](LICENSE). Third-party notices are in [NOTICE](NOTICE) and [THIRD-PARTY-LICENSES](THIRD-PARTY-LICENSES).
