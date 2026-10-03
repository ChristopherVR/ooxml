# xlsx-viewer-mcp

[![npm version](https://img.shields.io/npm/v/xlsx-viewer-mcp.svg)](https://www.npmjs.com/package/xlsx-viewer-mcp)
[![license](https://img.shields.io/npm/l/xlsx-viewer-mcp.svg)](https://github.com/ChristopherVR/xlsx-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/xlsx-viewer-mcp.svg)](https://www.npmjs.com/package/xlsx-viewer-mcp)

> Headless MCP access to document operations owned by OOXML core.

[Live demo](https://christophervr.github.io/xlsx-viewer/demo/) | [npm](https://www.npmjs.com/package/xlsx-viewer-mcp) | [Full docs](https://christophervr.github.io/xlsx-viewer/) | [Source](https://github.com/ChristopherVR/xlsx-viewer)

Repository-owned MCP schemas and server wiring. All document loading, editing,
calculation and serialization delegate to `ooxml-core/automation`.
Requires `ooxml-core` 0.11.0 or later, installed as a dependency.

## Install

```bash
npm install xlsx-viewer-mcp
```

## Quick start

```json
{
	"mcpServers": {
		"xlsx": {
			"command": "npx",
			"args": ["-y", "--package=xlsx-viewer-mcp", "xlsx-tools", "/path/to/documents"]
		}
	}
}
```

Paths are scoped to the configured root (the current directory by default).
Creation and save-as never overwrite existing files. Edit tools update the source
unless `outputPath` names a new file. Files are written atomically after core
validation. Unsupported features and preservation limitations are returned by core.

## Features

Workbook inspection, creation, bounded range reading and cell edits with recalculation.
All document algorithms live in OOXML core. Importing the package does not start stdio.

## Limitations

Unsupported operations report errors or core diagnostics. Office parity and lossless
export are not established. Updates overwrite the source unless `outputPath` is supplied.

## API

Programmatic integration: `createServer({ rootDir })` or
`registerTools(existingMcpServer, { rootDir })`. Importing the package does not
start a transport. The OOXML combined MCP calls this same registration function.

This package ships JavaScript source and needs no separate build step.

Package releases use npm trusted publishing with provenance.

## Documentation

[Source and tools](https://github.com/ChristopherVR/xlsx-viewer/tree/main/mcp) |
[Core automation](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md)

## License

Apache-2.0.
