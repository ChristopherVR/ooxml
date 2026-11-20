<div align="center">

# ooxml-core

**Read, edit, validate and write Word, Excel, PowerPoint and Visio files in TypeScript, in the browser or on a server.**

[![npm version](https://img.shields.io/npm/v/ooxml-core.svg)](https://www.npmjs.com/package/ooxml-core)
[![license](https://img.shields.io/npm/l/ooxml-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/ooxml-core.svg)](https://www.npmjs.com/package/ooxml-core)
[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)

[**Live demo**](https://christophervr.github.io/ooxml/) &nbsp;&middot;&nbsp;
[**Docs**](https://christophervr.github.io/ooxml/docs/) &nbsp;&middot;&nbsp;
[**Changelog**](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md) &nbsp;&middot;&nbsp;
[**Report a bug**](https://github.com/ChristopherVR/ooxml/issues/new/choose)

</div>

`ooxml-core` is the engine behind the [OOXML](https://github.com/ChristopherVR/ooxml#readme) Word, Excel, PowerPoint, Visio and OpenTeams apps. It has no UI and no framework: one package, one XML model and one packaging layer, with every format as an area you import on its own. It runs in browsers, Node.js, Bun, workers and serverless functions, and files never leave the process.

- **Every format, one structure.** WordprocessingML, SpreadsheetML, PresentationML and Visio share the same XML, OPC, units, colour, geometry, DrawingML, chart and SmartArt code.
- **Preserving round trips.** Untouched parts stay byte for byte, unknown markup is kept, and an edit that would damage content the model does not cover is rejected instead of silently dropped.
- **Built for files you did not write.** Every package goes through one guarded zip reader, the shared XML parser refuses DTDs and entity expansion, and the work a chart, a formula or a Visio drawing may cost is bounded, so a crafted file is refused instead of exhausting memory.
- **Signatures that mean something.** Digital signatures of every format are verified with pure TypeScript crypto in the browser and in Node, and every loader reports the same verified `SignatureState`.
- **A real Excel engine.** A dependency-graph formula engine with dynamic arrays, `LET` and `LAMBDA` and more than 480 functions, Excel number formats, editing commands with undo, and the DOM-free grid layout the Excel editor paints.
- **Honest about its limits.** Unsupported features are reported, never hidden. Nothing here claims Office parity or lossless export; open gaps are listed in the [parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md).

## Install

```bash
npm install ooxml-core
```

Optional peer dependencies turn on Node-only features:

| Peer              | Needed for                                                                                                 |
| ----------------- | ---------------------------------------------------------------------------------------------------------- |
| `node-forge`      | Signing PowerPoint packages and checking signer trust in Node (`ooxml-core/pptx/signature-node`)           |
| `@napi-rs/canvas` | Rasterising SVG pictures when a `.ppt` is saved in Node (`rasterizeSvgInNode` from `ooxml-core/pptx/node`) |

Verifying signatures needs neither: `ooxml-core/opc` verifies them in any runtime.

## Quick start

**Open a Word document, change it, save it**

```ts
import { loadDocx } from 'ooxml-core/docx';

const loaded = await loadDocx(bytes); // Uint8Array | ArrayBuffer
loaded.model.blocks; // paragraphs and tables
const saved = await loaded.save(); // parts the model does not cover are kept
```

**Build a PowerPoint deck**

```ts
import { PptxHandler } from 'ooxml-core/pptx';

const { handler, data, createSlide } = await PptxHandler.create({ title: 'Quarterly Review' });
data.slides.push(
	createSlide('Title Slide')
		.addPlaceholder('ctrTitle', 'Quarterly Review')
		.addPlaceholder('subTitle', 'Results and next steps')
		.build(),
);
const pptx = await handler.save(data.slides);
```

`addPlaceholder(type, text, { idx? })` binds text to the chosen layout and inherits its position, size and text styles; use `{ idx: 1 }` or `{ idx: 2 }` for the repeated content slots of Two Content, and `addText` for a free text box.

**Read an Excel workbook, including very large ones**

```ts
import {
	DEFAULT_XLSX_LOAD_LIMITS,
	isWorkbookTooLargeError,
	loadXlsx,
	readXlsxRows,
} from 'ooxml-core/xlsx';

// The default limits guard against zip bombs; raise only what your files need.
const workbook = await loadXlsx(bytes, {
	limits: { maxTotalBytes: 4 * DEFAULT_XLSX_LOAD_LIMITS.maxTotalBytes, maxXmlElements: 20_000_000 },
});

// Or stream one sheet's rows without building the workbook: memory does not grow with the sheet.
for await (const { row, cells } of readXlsxRows(bytes, { sheet: 'Orders' })) {
	for (const { col, value, formula, iso } of cells) {
		// zero-based row and col; iso is set for a number shown as a date or time
	}
}
```

Over a limit, loading throws an error that `isWorkbookTooLargeError(error)` recognises, with the `limit` it hit and its `maximum`. `loadWorkbook` from `ooxml-core/xlsx/load` also opens password-protected workbooks (`{ password }`), legacy `.xls` and CSV.

**Open a Visio drawing**

```ts
import { parseVsdx } from 'ooxml-core/visio';

const drawing = await parseVsdx(bytes);
for (const page of drawing.pages) console.log(page.name, page.shapes.length);
console.log(drawing.diagnostics); // what was approximated or left out
```

**Check a document's signatures**

```ts
import { signature } from 'ooxml-core/opc';
import { signatureStateText } from 'ooxml-core/i18n';

const summary = await signature.verifyPackageSignatureSummary(bytes); // any Office package
if (signature.isVerifiedSignatureState(summary.state)) {
	console.log(signatureStateText(summary.state, 'en')?.label);
}
```

**Compare two versions**

```ts
import { compareDocuments } from 'ooxml-core/docx';
import { compareWorkbooks, comparisonText, loadXlsx } from 'ooxml-core/xlsx';

const { bytes: redline } = await compareDocuments(v1, v2, { author: 'Reviewer' }); // tracked changes
console.log(comparisonText(compareWorkbooks(await loadXlsx(a), await loadXlsx(b))));
```

See [document comparison](https://github.com/ChristopherVR/ooxml/blob/main/docs/docx-compare.md) and [workbook comparison](https://github.com/ChristopherVR/ooxml/blob/main/docs/xlsx-compare.md) for what is compared.

## Areas

Every area is a subpath import, so you load only what you use. The document formats are imported through their own subpaths; the root entry groups the shared building blocks by namespace and never pulls in a format.

```ts
import { parseXml } from 'ooxml-core/xml'; // one area
import { opc, units } from 'ooxml-core'; // or the shared building blocks by namespace
```

| Area         | Subpaths                                                                                                                                                  | What it holds                                                                                                                                                                                                                                                             |
| ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `docx`       | `/docx`, `/docx/embedded`, `/docx/layout`, `/docx/load`, `/docx/ui`                                                                                       | WordprocessingML: model, parser, preserving writer, editing, review and comparison, validation against the ECMA-376 schemas in tests. `/docx/layout` is the DOM-free pagination engine (an approximation of Word); `/docx/load` adds encrypted `.docx` and legacy `.doc`. |
| `xlsx`       | `/xlsx`, `/xlsx/load`, `/xlsx/collab`, `/xlsx/ui`                                                                                                         | SpreadsheetML: workbook model, reader and preserving writer, formula engine, number formats, editing with undo, grid layout, pivots, slicers, auditing and comparison. `/xlsx/load` adds encrypted workbooks, `.xls` and CSV; `/xlsx/collab` shares a workbook over Yjs.  |
| `pptx`       | `/pptx`, `/pptx/node`, `/pptx/signature-node`, `/pptx/converter`, `/pptx/cli`, `/pptx/automation`, `/pptx/smartart-layouts`, `/pptx/ui`, `/pptx/editor/*` | PresentationML: model, parser, writer, the slide SDK, charts, SmartArt, media, collaboration mapping, converters, a CLI, signing and Node trust checks, and the PowerPoint automation tools.                                                                              |
| `visio`      | `/visio`, `/visio/collab`, `/visio/ui`                                                                                                                    | Visio: `.vsdx` parsing and bounded editing, layers, data graphics, stencils, a memory budget sized from the host heap, and a conservative preview of legacy `.vsd`.                                                                                                       |
| `opc`        | `/opc`                                                                                                                                                    | Open Packaging Conventions: relationships, content types, part names, the guarded zip reader, package comparison, document properties, the hyperlink and resource URL policies, and signature verification (`signature`).                                                 |
| `xml`        | `/xml`                                                                                                                                                    | The shared XML model: strict parsing without DTDs or entity expansion, serialization and namespace-aware helpers.                                                                                                                                                         |
| `drawingml`  | `/drawingml`                                                                                                                                              | Format-neutral DrawingML: colour, fill, outline, effect, text-body and geometry readers and writers, and the theme model.                                                                                                                                                 |
| `chart`      | `/chart`, `/chart/region-map`                                                                                                                             | Charts for every format: model, parser, writer, ChartEx, edits, embedded workbooks, the SVG painter, style presets and Office's Change Colors palettes.                                                                                                                   |
| `diagram`    | `/diagram`, `/diagram/builtin`                                                                                                                            | SmartArt: the part model, layout engine, node edits, part writers and the built-in layout, colour and style definitions.                                                                                                                                                  |
| `crypto`     | `/crypto`                                                                                                                                                 | ECMA-376 package encryption for every format (`decryptOoxmlPackage`, `encryptOoxmlPackage`, `isEncryptedOoxmlPackage`), agile AES-256 by default, with typed errors.                                                                                                      |
| `collab`     | `/collab`                                                                                                                                                 | Real-time collaboration on Yjs: sessions, providers, awareness, transport-neutral sync, bounded update decoding and share links.                                                                                                                                          |
| `teams`      | `/teams`                                                                                                                                                  | The OpenTeams workspace logic: channels and chat on Yjs, presence, WebRTC calls with pluggable signaling, and the `createTeamsClient` store every UI binds to.                                                                                                            |
| `automation` | `/automation`, `/automation/node`                                                                                                                         | Headless document tools shared by the MCP servers and assistants, and `registerAutomationTools` for an MCP server confined to one folder.                                                                                                                                 |
| `embedding`  | `/embedding`, `/embedding/onedrive`                                                                                                                       | Host contracts for embedding the editors (capabilities, customization, storage), and an optional OneDrive transport whose authentication the host supplies.                                                                                                               |
| `ai`         | `/ai`                                                                                                                                                     | The assistant policy a deployment sets: `normalizeAiPolicy`, `aiFeatureEnabled`, `checkAiEndpoint`, `userConnectionRefusal`. It makes no request.                                                                                                                         |
| Shared       | `/units`, `/color`, `/geometry`, `/text`, `/math`, `/i18n`, `/ribbon`, `/digest`                                                                          | Branded EMU, twip and point units; colour and WCAG contrast; preset shapes and boolean shape operations; text rules; Office Math to and from LaTeX and MathML; locales and translation; ribbon add-in contracts; synchronous digests and Office password hashes.          |

Legacy binary formats (`.doc`, `.xls`, `.ppt`) and the compound-file container come from the sibling package [`ole2`](https://github.com/ChristopherVR/ole2), inlined at build time into the subpaths that read legacy files (`/docx/load`, `/xlsx/load`, `/pptx`, `/visio`); the main `xlsx` entry never imports it.

## Format notes

### Signatures

`signature.verifyPackageSignatures` and `signature.verifyPackageSignatureSummary` from `ooxml-core/opc` verify every signature of a Word, Excel, PowerPoint or Visio package in a browser or in Node: Canonical XML, RSA PKCS#1 v1.5 (SHA-1 to SHA-512) and the digests are pure TypeScript. They check the signature value, each signed reference and that the signature covers the content, but trust no signer, so an intact signature is `verified-untrusted`. A method, key type or transform they do not implement reports `unsupported`, never invalid, and so does a signature past the verifier's work limits (an RSA key under 1024 or over 16384 bits, more than 10,000 references or 100 MB of referenced data, a signature part over 16 MB). The loaders set the result as `signatures` on every signed document, and `signatureStateText` in `ooxml-core/i18n` words each state in seven languages.

`inspectPptxDigitalSignatures` from `ooxml-core/pptx/signature-node` adds trust in Node. It reports `verified-trusted` only when every check above passes and the signer's certificate permits document signing and chains to a root you configured with `PPTX_VIEWER_TRUST_ROOTS_PEM` or `PPTX_VIEWER_TRUST_ROOTS_FILE`; no roots are built in. `coverage-incomplete` means a valid signature that leaves document content unsigned, listed in each detail's `uncoveredParts`. Integrity is decided first, so no revocation, timestamp or policy outcome can hide a failed check. Timestamp tokens are not validated and report `not-checked`. Revocation checks contact the OCSP responder a certificate names; pass `{ checkRevocation: false }` to make no request.

Relationships parts of a signed deck are read with one strict reader by both the verifier and `PptxHandler.load`; a signed deck the strict reader rejects opens like an unsigned one and is reported as not verified. The Word and Excel loaders read relationships with the tolerant `parseRelationships`, so a verified report for a `.docx` or `.xlsx` describes what the verifier read.

### Limits

Every loader has defaults that refuse a zip bomb or a pathological file before it costs much: `loadXlsx` and `loadWorkbook` take `limits`, `loadDocx` takes `limits` (`DEFAULT_DOCX_LOAD_LIMITS`), `parseVsdx` takes package limits including `maxXmlHeapBytes`, and `diffOpcPackages` takes a third argument. Each formula has a work and a text budget, and a formula past it returns `#NUM!` while the rest of the workbook still computes; `calculateOnOpen(workbook, { timeLimitMs })` runs a workbook's calculation on open under a time limit.

### CSV

CSV import (`csvToWorkbook`, or `loadWorkbook(bytes, { csv })`) types fields the way typing them into a cell does, which approximates Excel opening a CSV. It takes Excel's own controls: the automatic data conversion switches (`conversions: { leadingZeros: false }` keeps `00123`) and the Text Import Wizard's column formats (`columns: { 0: 'text', 2: 'DMY', 5: 'skip' }`), separators and trailing minus. Each column a conversion changed gets one entry in `workbook.warnings`.

### Charts

`ooxml-core/chart` keeps the point counts a file declares through viewing, editing and saving, while the blanks a view materialises come from a bounded budget (`CHART_POINTS_LIMITED` reports a shortened view). A chart edit writes only the embedded workbook cells it changed. Office's 17 Change Colors palettes are available through `CHART_COLOR_PALETTES` and `chartPaletteColors(palette, seriesCount, scheme)`; some themes and large series counts still differ from Excel by one colour step.

## Supported environments

- Node.js 22.3 or newer (`"engines": { "node": ">=22.3" }`), and Bun.
- The last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least; see [browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md).
- Every subpath ships ESM and CommonJS, with TypeScript declarations.

## Privacy and network use

`ooxml-core` parses, edits and writes documents locally and never uploads them. It contacts a server only when you ask it to: a source URL you pass, collaboration or Teams servers you configure, OneDrive when you set it up, and, during Node signature inspection, the OCSP responder named in a signing certificate (turn it off with `{ checkRevocation: false }`). Linked pictures and media are not fetched unless the load allows linked content. Nothing collects analytics or telemetry. Every request is listed in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

- [`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui): the browser elements and the Word, Excel, PowerPoint, Visio and Teams editors built on this package.
- The editors for your framework: [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme), [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme), [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme), [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) and [OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme), for React, Vue, Angular, Svelte, Solid and plain JavaScript.
- [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp): one MCP server that gives an AI agent the `ooxml-core/automation` tools.
- [`docx-core`](https://www.npmjs.com/package/docx-core) and [`@christophervr/xlsx-core`](https://www.npmjs.com/package/@christophervr/xlsx-core): re-exports of `ooxml-core/docx` and `ooxml-core/xlsx` with their own stability snapshots.

## Contributing

`ooxml-core` is developed in [ChristopherVR/ooxml](https://github.com/ChristopherVR/ooxml), together with `ooxml-ui`, the MCP servers and the viewers. Start with [CONTRIBUTING.md](https://github.com/ChristopherVR/ooxml/blob/main/CONTRIBUTING.md); the [roadmap](https://github.com/ChristopherVR/ooxml/blob/main/docs/roadmap.md) says what comes next, and [PROVENANCE.md](https://github.com/ChristopherVR/ooxml/blob/main/PROVENANCE.md) records where each module came from. Report security problems privately as [SECURITY.md](https://github.com/ChristopherVR/ooxml/blob/main/SECURITY.md) describes.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/LICENSE). Third-party notices are in [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/src/core/NOTICE) and [THIRD-PARTY-LICENSES](https://github.com/ChristopherVR/ooxml/blob/main/src/core/THIRD-PARTY-LICENSES).
