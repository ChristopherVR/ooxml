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

`ooxml-core` is the headless engine behind the [ooxml](https://github.com/ChristopherVR/ooxml#readme) Word, Excel, PowerPoint, Visio and OpenTeams editors. It is for applications and services that open, change or produce Office files without Office: one package with no UI, no framework and no native code, where every format is an area you import on its own. It runs in browsers, Node.js, Bun, workers and serverless functions, and files never leave the process.

- **One structure for every format.** WordprocessingML, SpreadsheetML, PresentationML and Visio share one XML model, one packaging layer, and the same units, colour, geometry, DrawingML, chart and SmartArt code.
- **Round trips that keep what they do not understand.** Untouched parts stay byte for byte, unknown markup is kept, and an edit that would damage content the model does not cover is refused instead of silently dropped.
- **Safe on files you did not write.** One guarded zip reader, an XML parser that refuses DTDs and entity expansion, and bounded work for formulas, charts and Visio drawings, so a crafted file is refused instead of exhausting memory.
- **A real Excel engine.** A dependency-graph formula engine with dynamic arrays, `LET`, `LAMBDA` and more than 500 functions ([function list](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/docs/functions.md)), number formats, and editing commands with undo.
- **Signatures that are verified, not just detected**, in pure TypeScript in every runtime.

## Install

```bash
npm install ooxml-core
```

Optional peer dependencies turn on Node-only features; nothing else needs them.

| Peer              | Needed for                                                                                                 |
| ----------------- | ---------------------------------------------------------------------------------------------------------- |
| `node-forge`      | Signing PowerPoint packages and checking signer trust in Node (`ooxml-core/pptx/signature-node`)           |
| `@napi-rs/canvas` | Rasterising SVG pictures when a `.ppt` is saved in Node (`rasterizeSvgInNode` from `ooxml-core/pptx/node`) |

## Quick start

**Open a Word document, change it and save it**

```ts
import { loadDocx } from 'ooxml-core/docx';

const loaded = await loadDocx(bytes); // Uint8Array | ArrayBuffer
loaded.model.blocks; // paragraphs and tables
const saved = await loaded.save(); // parts the model does not cover are kept
```

**Fill a Word template**

```ts
import { loadDocx } from 'ooxml-core/docx';
import { mergeFields } from 'ooxml-core/docx/mail-merge';

const template = await loadDocx(bytes);
const { model, results } = await mergeFields(
	template.model,
	{ values: { Name: 'Ada' }, regions: { Items: [{ Item: 'Pen' }, { Item: 'Ink' }] } },
	{ unlink: true },
);
const filled = await template.save(model); // MERGEFIELD values, TableStart/TableEnd regions, HTML
```

See [filling a template](https://github.com/ChristopherVR/ooxml/blob/main/docs/docx-mail-merge.md#filling-a-template).

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

**Read an Excel workbook, or stream one sheet of a very large one**

```ts
import { loadXlsx, readXlsxRows } from 'ooxml-core/xlsx';

const workbook = await loadXlsx(bytes);

for await (const { row, cells } of readXlsxRows(bytes, { sheet: 'Orders' })) {
	for (const { col, value, formula } of cells) {
		// zero-based row and col; memory does not grow with the sheet
	}
}
```

`loadWorkbook` from `ooxml-core/xlsx/load` also opens password-protected workbooks (`{ password }`), legacy `.xls` and CSV.

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

const summary = await signature.verifyPackageSignatureSummary(bytes); // any Office package
signature.isVerifiedSignatureState(summary.state);
```

**Compare two versions**

```ts
import { compareDocuments } from 'ooxml-core/docx';
import { compareWorkbooks, comparisonText, loadXlsx } from 'ooxml-core/xlsx';

const { bytes: redline } = await compareDocuments(v1, v2, { author: 'Reviewer' }); // tracked changes
console.log(comparisonText(compareWorkbooks(await loadXlsx(a), await loadXlsx(b))));
```

See [document comparison](https://github.com/ChristopherVR/ooxml/blob/main/docs/docx-compare.md) and [workbook comparison](https://github.com/ChristopherVR/ooxml/blob/main/docs/xlsx-compare.md).

## Areas

Every area is a subpath, so you load only what you use. The root entry groups the shared building blocks by namespace (`import { opc, units } from 'ooxml-core'`) and never pulls in a document format.

| Area         | Subpaths                                                                                                                                                  | What it holds                                                                                                                                                                                                           |
| ------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `docx`       | `/docx`, `/docx/embedded`, `/docx/layout`, `/docx/load`, `/docx/mail-merge`, `/docx/pdf-import`, `/docx/pdf-convert`, `/docx/ui`                          | WordprocessingML: model, parser, preserving writer, editing, review and comparison. `/docx/layout` is the DOM-free pagination engine (an approximation of Word); `/docx/load` adds encrypted `.docx` and legacy `.doc`. |
| `xlsx`       | `/xlsx`, `/xlsx/load`, `/xlsx/collab`, `/xlsx/ui`                                                                                                         | SpreadsheetML: model, reader and preserving writer, formula engine, number formats, editing with undo, grid layout, pivots, slicers, auditing and comparison. `/xlsx/load` adds encrypted workbooks, `.xls` and CSV.    |
| `pptx`       | `/pptx`, `/pptx/node`, `/pptx/signature-node`, `/pptx/converter`, `/pptx/cli`, `/pptx/automation`, `/pptx/smartart-layouts`, `/pptx/ui`, `/pptx/editor/*` | PresentationML: model, parser, writer, the slide builder, charts, SmartArt, media, converters, a CLI, signing and Node trust checks, and the PowerPoint automation tools.                                               |
| `visio`      | `/visio`, `/visio/collab`, `/visio/ui`                                                                                                                    | Visio: `.vsdx` parsing and bounded editing, layers, data graphics, stencils, and a conservative preview of legacy `.vsd`.                                                                                               |
| `opc`        | `/opc`                                                                                                                                                    | Packaging: relationships, content types, the guarded zip reader, package comparison, document properties, URL policies and signature verification (`signature`).                                                        |
| `xml`        | `/xml`                                                                                                                                                    | The shared XML model: strict parsing, serialization and namespace-aware helpers.                                                                                                                                        |
| `drawingml`  | `/drawingml`                                                                                                                                              | Format-neutral DrawingML: colour, fill, outline, effect, text-body and geometry readers and writers, and the theme model.                                                                                               |
| `chart`      | `/chart`, `/chart/region-map`                                                                                                                             | Charts for every format: model, parser, writer, ChartEx, edits, embedded workbooks, the SVG painter and style presets.                                                                                                  |
| `diagram`    | `/diagram`, `/diagram/builtin`                                                                                                                            | SmartArt: the part model, layout engine, node edits, part writers and the built-in definitions.                                                                                                                         |
| `pdf`        | `/pdf`                                                                                                                                                    | A DOM-free PDF reader (`openPdf`) that never runs the file's scripts or actions; `/docx/pdf-import` turns its output into an editable, approximate Word document.                                                       |
| `crypto`     | `/crypto`                                                                                                                                                 | ECMA-376 package encryption for every format (`decryptOoxmlPackage`, `encryptOoxmlPackage`, `isEncryptedOoxmlPackage`).                                                                                                 |
| `collab`     | `/collab`                                                                                                                                                 | Real-time collaboration on Yjs: sessions, providers, awareness, transport-neutral sync and bounded update decoding.                                                                                                     |
| `teams`      | `/teams`                                                                                                                                                  | The OpenTeams workspace logic: channels and chat on Yjs, presence, WebRTC calls with pluggable signaling, and the `createTeamsClient` store.                                                                            |
| `automation` | `/automation`, `/automation/node`, `/automation/add-ins`                                                                                                  | Headless document tools shared by the MCP servers and assistants; `registerAutomationTools` registers them on an MCP server confined to one folder; `/automation/add-ins` reads Office Add-in manifests offline.        |
| `embedding`  | `/embedding`, `/embedding/onedrive`                                                                                                                       | Host contracts for embedding the editors (capabilities, customization, storage) and an optional OneDrive transport whose sign-in the host supplies.                                                                     |
| `ai`         | `/ai`                                                                                                                                                     | The assistant policy a deployment sets (`normalizeAiPolicy`, `aiFeatureEnabled`, `checkAiEndpoint`, `userConnectionRefusal`). It makes no request.                                                                      |
| Shared       | `/units`, `/color`, `/geometry`, `/text`, `/math`, `/i18n`, `/ribbon`, `/digest`                                                                          | Branded EMU, twip and point units; colour and contrast; preset shapes; text rules; Office Math to and from LaTeX; locales and translation; ribbon add-in contracts; digests and Office password hashes.                 |

Legacy `.doc`, `.xls` and `.ppt` and the compound-file container come from the sibling package [`ole2`](https://github.com/ChristopherVR/ole2), inlined into the subpaths that read legacy files (`/docx/load`, `/xlsx/load`, `/pptx`, `/visio`). The main `xlsx` entry never imports it.

## Safety and limits

- **Load limits.** `loadDocx` (`DEFAULT_DOCX_LOAD_LIMITS`), `loadXlsx` and `loadWorkbook` (`DEFAULT_XLSX_LOAD_LIMITS`) and `parseVsdx` take `limits` whose defaults refuse a zip bomb before it costs much; a workbook over a limit throws an error that `isWorkbookTooLargeError` recognises. Raise only what your files need.
- **Formula budgets.** A formula past its work or text budget returns `#NUM!` while the rest of the workbook still computes; `calculateOnOpen(workbook, { timeLimitMs })` bounds calculation on open.
- **Signatures.** `ooxml-core/opc` checks the signature value, every signed reference and that the signature covers the content, but trusts no signer, so an intact signature is `verified-untrusted`; a method it does not implement reports `unsupported`, never invalid. `inspectPptxDigitalSignatures` from `ooxml-core/pptx/signature-node` adds trust in Node, against roots you supply through `PPTX_VIEWER_TRUST_ROOTS_PEM` or `PPTX_VIEWER_TRUST_ROOTS_FILE` (none are built in).
- **Security problems:** report them privately, as [SECURITY.md](https://github.com/ChristopherVR/ooxml/blob/main/SECURITY.md) describes.

## Status

`ooxml-core` follows the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md): its public API is the subpaths in `exports`. The formats are implemented independently from the published specifications and from files Office saves. They do not claim Office parity or lossless export of content a model does not cover, and unsupported features are reported rather than hidden. Open gaps are listed in the [parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md).

## Privacy and network use

`ooxml-core` parses, edits and writes documents locally and collects no analytics or telemetry. It contacts another server only when you ask it to:

- a source URL you pass instead of bytes;
- collaboration or Teams servers you configure, and OneDrive when you set it up (an active Word collaboration session also shares inserted pictures, 3D models and Office objects with the room);
- Excel `IMAGE` and `WEBSERVICE` formulas, only through a provider you supply (for example `createFetchImageProvider({ fetch, allowedOrigins })` as `providers.image`, or `createFetchWebserviceProvider` as `providers.webservice`); without one `IMAGE` returns `#BLOCKED!` and `WEBSERVICE` `#NAME?`, and the Excel editor passes no `WEBSERVICE` provider;
- in Node, the OCSP responder a signing certificate names, during `inspectPptxDigitalSignatures`; pass `{ checkRevocation: false }` to make no request.

In a web build, a few lazy pieces (Excel's non-English dictionaries, the XPath code for `FILTERXML`, Visio's TIFF decoder and embedded chart reader) load from the app's own origin and carry no document data. Linked pictures and media are not fetched unless the load allows linked content. Every request is listed in [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Environments

- Node.js 22.3 or newer, and Bun.
- The last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- Every subpath ships ESM and CommonJS with TypeScript declarations.

## Documentation

- The [documentation hub](https://christophervr.github.io/ooxml/docs/) and the API guides for [Word](https://christophervr.github.io/ooxml/docx/api), [Excel](https://christophervr.github.io/ooxml/xlsx/api), [PowerPoint](https://christophervr.github.io/ooxml/pptx/core/) and [Visio](https://christophervr.github.io/ooxml/visio/api).
- [`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui): the browser elements and editors built on this package. Every framework package is listed in the [ooxml README](https://github.com/ChristopherVR/ooxml#packages).
- [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp): the `ooxml-core/automation` tools for AI agents.
- [`docx-core`](https://www.npmjs.com/package/docx-core) and [`@christophervr/xlsx-core`](https://www.npmjs.com/package/@christophervr/xlsx-core): re-exports of `ooxml-core/docx` and `ooxml-core/xlsx` with their own stability snapshots.
- Contributing: [CONTRIBUTING.md](https://github.com/ChristopherVR/ooxml/blob/main/CONTRIBUTING.md), the [roadmap](https://github.com/ChristopherVR/ooxml/blob/main/docs/roadmap.md) and [PROVENANCE.md](https://github.com/ChristopherVR/ooxml/blob/main/PROVENANCE.md).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/LICENSE). Third-party notices are in [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/src/core/NOTICE) and [THIRD-PARTY-LICENSES](https://github.com/ChristopherVR/ooxml/blob/main/src/core/THIRD-PARTY-LICENSES).
