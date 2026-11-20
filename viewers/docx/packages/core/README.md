# docx-core

**The Word document model, parser and preserving serializer: read, inspect, edit and write `.docx` files in Node or the browser, with no DOM and no editor.**

[![npm version](https://img.shields.io/npm/v/docx-core.svg)](https://www.npmjs.com/package/docx-core)
[![license](https://img.shields.io/npm/l/docx-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE)
[![types](https://img.shields.io/npm/types/docx-core.svg)](https://www.npmjs.com/package/docx-core)

[Live demo](https://christophervr.github.io/ooxml/docx/demo/) · [Documentation](https://christophervr.github.io/ooxml/docx/) · [API](https://christophervr.github.io/ooxml/docx/api) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/core/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

`docx-core` is a stable entry point over the `docx` area of [`ooxml-core`](https://www.npmjs.com/package/ooxml-core): `docx-core` re-exports `ooxml-core/docx`, and `docx-core/embedded` re-exports `ooxml-core/docx/embedded`. It is the same model every Word editor package uses.

## Install

```bash
npm install docx-core
```

It depends only on `ooxml-core`, which installs with it.

## Quick start

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { loadDocx, plainText } from 'docx-core';

const loaded = await loadDocx(await readFile('report.docx'));
console.log(plainText(loaded.model));

// Change the model as a value, then save in the original format.
const [first] = loaded.model.blocks;
const model =
	first?.type === 'paragraph'
		? {
				...loaded.model,
				blocks: [
					{ ...first, runs: [{ text: 'Quarterly report' }] },
					...loaded.model.blocks.slice(1),
				],
			}
		: loaded.model;
await writeFile('report-edited.docx', await loaded.save(model));
```

A new document starts from `createDocument()` and is written with `saveDocx(model)`, which resolves to the `.docx` bytes.

`loaded.save()` writes only what changed: an unchanged model returns the original bytes, package parts the model does not cover are kept, and an edit that cannot be saved safely rejects instead of silently degrading the file. Treat the model as a value: build a new object for each change instead of changing one in place.

## What you get

- **Model.** `loadDocx`, `saveDocx`, `createDocument` and the `DocumentModel` types: paragraphs and runs, styles and numbering, tables, sections and columns, headers and footers, notes, fields, content controls, drawings, charts and SmartArt.
- **Safe loading.** Every package goes through a guarded zip reader with size and part limits. `loadDocx(bytes, { limits })` sets them (`DocxLoadLimits`); the defaults are `DEFAULT_DOCX_LOAD_LIMITS`.
- **Review.** `listRevisions`, `acceptRevision`, `rejectRevision`, `acceptAllRevisions` and `rejectAllRevisions` for tracked changes, and `compareDocuments` to produce a redline of two documents as tracked changes.
- **Forms and protection.** `listFormFields` and `fillFormFields` for content controls and legacy form fields, and `enforceProtection` and `verifyProtectionPassword` for editing restrictions.
- **Fields and contents.** `updateTableOfContents`, `refreshComputedFields`, `documentStats` and `plainText`.
- **Charts and SmartArt.** `editDocxChart`, `newDocxChart`, `editDocxDiagram` and `createDocxDiagram` run the shared chart and SmartArt edits against a Word document.
- **Inspection.** `inspectDocument` and `removeDocumentInformation` for Inspect Document, and `validateDocumentModel` to check a model before saving.
- **Signatures.** A signed document carries `model.signatures`, the verified state of its signatures (`verified`, `verified-untrusted`, `invalid` or `unsupported`). Only a signature that verifies counts as signed.
- **Embedded documents.** `docx-core/embedded` has `readOleDocumentParagraphs` and `writeOleDocumentParagraphEdit`, the narrow API another product uses for a Word document embedded in its file.

The [API reference](https://christophervr.github.io/ooxml/docx/api) lists the rest, and [model units](https://christophervr.github.io/ooxml/docx/model-units) explains the twips, half-points and EMUs the model uses.

## Known limits

- `docx-core` reads and writes `.docx`. Legacy Word 97-2003 `.doc` and password-protected `.docx` files open through `loadDocument` from `ooxml-core/docx/load`, which every Word editor package re-exports.
- Saving is careful, not lossless for features the model does not support; the [preservation guarantee](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/docs/preservation.md) says exactly what is kept.
- There is no layout here: pagination and line breaking live in `ooxml-core/docx/layout` and the editor packages.
- The open gaps are listed in the [support roadmap](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/docs/parity-roadmap.md) and the [parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md).

## Supported environments

- Node.js 22.3 or newer.
- Browsers: the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least. See [browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md).
- API stability: `docx-core` is Stable under the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md), and its public surface is checked against a snapshot on every change.

## Privacy and network use

`docx-core` makes no network request: it parses and writes the bytes you give it. Linked pictures are not fetched and linked data is not refreshed. Font downloads are a feature of the editor packages, off by default there. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

- The Word editor for your framework, which re-exports this API: [`docx-react-viewer`](https://www.npmjs.com/package/docx-react-viewer), [`docx-vue-viewer`](https://www.npmjs.com/package/docx-vue-viewer), [`docx-angular-viewer`](https://www.npmjs.com/package/docx-angular-viewer), [`docx-svelte-viewer`](https://www.npmjs.com/package/docx-svelte-viewer), [`docx-solid-viewer`](https://www.npmjs.com/package/docx-solid-viewer) and [`docx-vanilla-viewer`](https://www.npmjs.com/package/docx-vanilla-viewer).
- [`docx-viewer-mcp`](https://www.npmjs.com/package/docx-viewer-mcp): Word tools for AI assistants over MCP.
- [`ooxml-core`](https://www.npmjs.com/package/ooxml-core): all the Office logic, including Excel, PowerPoint and Visio. See the [ooxml repository](https://github.com/ChristopherVR/ooxml#readme).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE). See [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/NOTICE) for attributions.
