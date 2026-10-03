# ooxml-mcp

[![npm version](https://img.shields.io/npm/v/ooxml-mcp.svg)](https://www.npmjs.com/package/ooxml-mcp)
[![license](https://img.shields.io/npm/l/ooxml-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/ooxml-mcp.svg)](https://www.npmjs.com/package/ooxml-mcp)

> Headless MCP access to document operations owned by OOXML core.

[Try the apps](https://christophervr.github.io/ooxml/) | [npm](https://www.npmjs.com/package/ooxml-mcp) | [Full docs](https://christophervr.github.io/ooxml/) | [Source](https://github.com/ChristopherVR/ooxml)

Combined MCP view of the format tools owned by the viewer repositories.
No schemas, document algorithms or mutations are duplicated here.

Install `ooxml-mcp` and the format packages you want to expose:
`docx-viewer-mcp`, `xlsx-viewer-mcp`, `visio-viewer-mcp` and
`pptx-viewer-mcp` (the version exposing `registerTools`).

## Install

```bash
npm install ooxml-mcp docx-viewer-mcp xlsx-viewer-mcp visio-viewer-mcp pptx-viewer-mcp
```

## Quick start

Run `ooxml-tools /path/to/documents` for all four formats, or
`ooxml-tools /path/to/documents docx,xlsx,visio` for a selected set.
A missing selected package is a startup error, so tools are never silently omitted.
The format servers require the published `ooxml-core` 0.11.0 or later.

Each standalone package exports `registerTools(server, { rootDir })`.
The combined server loads those exact functions. Core alone owns parsing,
document models, editing, recalculation, preservation and serialization.
PowerPoint retains its existing unprefixed tool names. The combined server
passes the same document root to all four tool registrations.
Other tools use format prefixes (`docx_`, `xlsx_`, `visio_`).

## Features

Composes the actual Word, Excel, Visio and PowerPoint tool registrations. The viewer
packages own schemas and registration; OOXML core owns document algorithms.

## Limitations

Selected format packages must be installed. Unsupported operations report errors or
core diagnostics. Office parity and lossless export are not established.

## API

Programmatic use: `await createServer({ rootDir, formats: ['docx', 'xlsx'] })`.
Importing this package does not connect stdio. All documents remain local.
This package ships JavaScript source and needs no separate build.

## Documentation

[Source and tools](https://github.com/ChristopherVR/ooxml/tree/main/mcp) |
[Core automation](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md)

## License

Apache-2.0.
