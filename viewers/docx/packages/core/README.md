# docx-core

[![npm version](https://img.shields.io/npm/v/docx-core.svg)](https://www.npmjs.com/package/docx-core)
[![license](https://img.shields.io/npm/l/docx-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE)
[![types](https://img.shields.io/npm/types/docx-core.svg)](https://www.npmjs.com/package/docx-core)

> The DOM-free document model and file API, re-exported from the canonical OOXML core.

[Live demo](https://christophervr.github.io/ooxml/docx/demo/) | [npm](https://www.npmjs.com/package/docx-core) | [Full docs](https://christophervr.github.io/ooxml/docx/) | [Source](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; **[Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme)** &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; [OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme) &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

The framework-neutral DOCX document model, parser and preserving serializer. A thin entry point: `docx-core` re-exports `ooxml-core/docx` and `docx-core/embedded` re-exports `ooxml-core/docx/embedded`; the logic lives in [ooxml-core](https://github.com/ChristopherVR/ooxml).

## Install

```bash
npm install docx-core
```

## Quick start

```ts
import { createDocument, loadDocx, saveDocx } from 'docx-core';

const model = createDocument();
const bytes = await saveDocx(model);
const loaded = await loadDocx(bytes);
```

## API and limitations

It handles modern `.docx` only. To open legacy `.doc` files use one of the editor packages (`docx-react-viewer`, `-vue-viewer`, `-angular-viewer`, `-svelte-viewer`, `-solid-viewer`, `-vanilla-viewer`), whose `loadDocument` detects both formats. Saving is not lossless for unsupported features.

## Documentation

[Full docs](https://christophervr.github.io/ooxml/docx/) | [Core source](https://github.com/ChristopherVR/ooxml)

## Privacy and network use

The Word packages make no network request when opening, rendering, editing or saving a document: linked pictures are not fetched, linked data is not refreshed, and fonts come from the device only (nothing is downloaded from Google Fonts or any other font service). Hyperlinks open only when the reader follows them. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## License

Apache-2.0.

## OneDrive integration

Use the optional shared `ooxml-core/embedding/onedrive` connection with host-supplied Microsoft
sign-in. See the [Office package and integration reference](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md#word-docx)
for this product's load/save API, native storage wiring, sharing permissions and conflict recovery.
The [testing guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-testing.md)
separates test-local transports from live Microsoft acceptance. No simulation service is shipped;
OneDrive storage does not provide Microsoft Office coauthoring or automatic document merging.
