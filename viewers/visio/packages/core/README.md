# visio-core

[![npm version](https://img.shields.io/npm/v/visio-core.svg)](https://www.npmjs.com/package/visio-core)
[![license](https://img.shields.io/npm/l/visio-core.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE)
[![types](https://img.shields.io/npm/types/visio-core.svg)](https://www.npmjs.com/package/visio-core)

> The DOM-free Visio document API. This package is a thin re-export of `ooxml-core/visio`; it contains no separate parser, editing engine or viewer.

[Live demo](https://christophervr.github.io/ooxml/visio/demo/) | [npm](https://www.npmjs.com/package/visio-core) | [Full docs](https://christophervr.github.io/ooxml/visio/) | [Source](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme) &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; **[Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme)** &middot; [OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme) &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

## Install

```bash
npm install visio-core
```

## Quick start

```js
import { readFile } from 'node:fs/promises';
import { parseVsdx } from 'visio-core';

const document = await parseVsdx(await readFile('diagram.vsdx'));
console.log(document.pages);
```

## API and limitations

Exports include the document model, `parseVsdx`, layer helpers and the core's
supported edit and geometry APIs. The installed `ooxml-core` version defines the
exact API. Review returned diagnostics and validation errors; native Visio parity
and lossless export are not established.

For a browser UI, install `visio-react-viewer`, `visio-vue-viewer`,
`visio-angular-viewer`, `visio-svelte-viewer`, `visio-solid-viewer` or
`visio-vanilla-viewer`. Each already re-exports this document API.
For headless MCP access, use `visio-viewer-mcp`.

## Documentation

[Viewer guide](https://christophervr.github.io/ooxml/visio/docs/) |
[Core source](https://github.com/ChristopherVR/ooxml)

## License

Apache-2.0.
