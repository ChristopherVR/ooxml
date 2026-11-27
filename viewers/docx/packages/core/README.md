# docx-core

**The Word document model, parser and preserving serializer: read, inspect, edit and write `.docx` files in Node or the browser, with no DOM and no editor.**

[![npm version](https://img.shields.io/npm/v/docx-core.svg)](https://www.npmjs.com/package/docx-core)
[![license](https://img.shields.io/npm/l/docx-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE)
[![types](https://img.shields.io/npm/types/docx-core.svg)](https://www.npmjs.com/package/docx-core)

[Documentation](https://christophervr.github.io/ooxml/docx/) · [API](https://christophervr.github.io/ooxml/docx/api) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/core/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

`docx-core` re-exports the `docx` area of [`ooxml-core`](https://www.npmjs.com/package/ooxml-core), and `docx-core/embedded` re-exports `ooxml-core/docx/embedded`. It is the same model every Word editor package uses, with no logic of its own.

## Install

```bash
npm install docx-core
```

It depends only on `ooxml-core`, which installs with it.

## Quick start

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { createDocument, loadDocx, plainText, saveDocx } from 'docx-core';

const loaded = await loadDocx(await readFile('report.docx'));
console.log(plainText(loaded.model));

// The model is a value: build a new object for each change, never change one in place.
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

// A new document:
await writeFile('blank.docx', await saveDocx(createDocument()));
```

`loaded.save()` writes only what changed: an unchanged model returns the original bytes, package parts the model does not cover are kept, and an edit that cannot be saved safely rejects instead of silently degrading the file. `saveDocx(model)` writes a document from the model alone. Both resolve to the `.docx` bytes.

## API overview

| Area                | Main entry points                                                                                                                                        |
| ------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Load and save       | `loadDocx(bytes, { limits })` (a guarded zip reader; defaults in `DEFAULT_DOCX_LOAD_LIMITS`), `saveDocx`, `createDocument` and the `DocumentModel` types |
| Review              | `listRevisions`, `acceptRevision`, `rejectRevision`, `acceptAllRevisions`, `rejectAllRevisions`, `revisionSpans`, and `compareDocuments` for a redline   |
| Forms and fields    | `listFormFields`, `fillFormFields`, `updateTableOfContents`, `refreshComputedFields`, `refreshStyleRefFields`                                            |
| Protection          | `enforceProtection`, `verifyProtectionPassword`                                                                                                          |
| Contents and checks | `plainText`, `documentStats`, `inspectDocument`, `removeDocumentInformation`, `validateDocumentModel`                                                    |
| Embedded documents  | `docx-core/embedded`: `readOleDocumentParagraphs`, `writeOleDocumentParagraphEdit`                                                                       |

A signed document carries `model.signatures` with each signature's state (`verified`, `verified-untrusted`, `invalid` or `unsupported`); only one that verifies counts as signed.

The package re-exports a whole core area, so it exports more names than it supports. The **supported** surface, covered by semver, is listed in the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md#the-supported-surface-of-docx-core-and-christophervrxlsx-core); every other exported name is internal and may change in a minor release. The [API reference](https://christophervr.github.io/ooxml/docx/api) and [model units](https://christophervr.github.io/ooxml/docx/model-units) (twips, half-points and EMUs) have the details.

## Status and limits

Word is **Stable**: the exported surface is checked against a committed snapshot on every change, and saving follows the [preservation guarantee](https://christophervr.github.io/ooxml/docx/preservation). That is not a lossless round trip for features the model does not support.

- `docx-core` reads and writes `.docx` only. Legacy Word 97-2003 `.doc` and password-protected `.docx` files open through `loadDocument` from `ooxml-core/docx/load`, which every Word editor package re-exports.
- There is no layout here: pagination and line breaking live in `ooxml-core/docx/layout` and the editor packages.
- Some fields (MERGEFIELD, INCLUDETEXT, INDEX, CITATION and others) keep their cached result instead of being computed.
- Canvases, legacy VML text boxes and WordArt are kept but not modelled for drawing, and ruby text has no model.

The open gaps are in the [support roadmap](https://christophervr.github.io/ooxml/docx/parity-roadmap) and the [parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md).

## Privacy and network use

`docx-core` makes no network request: it parses and writes the bytes you give it. Linked pictures are not fetched and linked data is not refreshed. Font downloads are a feature of the editor packages, off by default there. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#word-excel-and-visio).

## Environments

- Node.js 22.3 or newer.
- Browsers: the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
- It needs no DOM, so it can also run in a Web Worker.

## Documentation

- [Documentation](https://christophervr.github.io/ooxml/docx/), the [API reference](https://christophervr.github.io/ooxml/docx/api) and the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md).
- The Word editor for your framework, which re-exports this API: [`docx-react-viewer`](https://www.npmjs.com/package/docx-react-viewer), [`docx-vue-viewer`](https://www.npmjs.com/package/docx-vue-viewer), [`docx-angular-viewer`](https://www.npmjs.com/package/docx-angular-viewer), [`docx-svelte-viewer`](https://www.npmjs.com/package/docx-svelte-viewer), [`docx-solid-viewer`](https://www.npmjs.com/package/docx-solid-viewer) and [`docx-vanilla-viewer`](https://www.npmjs.com/package/docx-vanilla-viewer).
- [`docx-viewer-mcp`](https://www.npmjs.com/package/docx-viewer-mcp): Word tools for AI assistants over MCP.
- [`ooxml-core`](https://www.npmjs.com/package/ooxml-core) and the [ooxml repository](https://github.com/ChristopherVR/ooxml#readme): all the Office logic, including Excel, PowerPoint and Visio.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE). See [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/NOTICE) for attributions.
