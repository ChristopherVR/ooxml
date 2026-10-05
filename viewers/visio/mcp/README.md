# visio-viewer-mcp

[![npm version](https://img.shields.io/npm/v/visio-viewer-mcp.svg)](https://www.npmjs.com/package/visio-viewer-mcp)
[![license](https://img.shields.io/npm/l/visio-viewer-mcp.svg)](https://github.com/ChristopherVR/visio-viewer/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/visio-viewer-mcp.svg)](https://www.npmjs.com/package/visio-viewer-mcp)

> Headless MCP access to document operations owned by OOXML core.

[Live demo](https://christophervr.github.io/visio-viewer/demo/) | [npm](https://www.npmjs.com/package/visio-viewer-mcp) | [Full docs](https://christophervr.github.io/visio-viewer/) | [Source](https://github.com/ChristopherVR/visio-viewer)

Repository-owned MCP schemas and server wiring. All document loading, editing
and serialization delegate to `ooxml-core/automation`.
Requires `ooxml-core` 0.11.0 or later. It is installed as a dependency.

## Install

```bash
npm install visio-viewer-mcp
```

## Quick start

```json
{
	"mcpServers": {
		"visio": {
			"command": "npx",
			"args": ["-y", "--package=visio-viewer-mcp", "visio-tools", "/path/to/documents"]
		}
	}
}
```

Paths are scoped to the configured root (the current directory by default).
Save-as never overwrites existing files. Edit tools update the source
unless `outputPath` names a new file. Files are written atomically after core
validation. Unsupported features and preservation limitations are returned by core.

## Features

VSDX inspection and batches of supported text and geometry edits. All document
algorithms live in OOXML core. Importing the package does not start stdio.

## Limitations

Unsupported operations report errors or core diagnostics. Office parity and lossless
export are not established. Updates overwrite the source unless `outputPath` is supplied.

## API

Programmatic integration: `createServer({ rootDir })` or
`registerTools(existingMcpServer, { rootDir })`. Importing the package does not
start a transport. The OOXML combined MCP calls this same registration function.

This package ships JavaScript source and needs no separate build step.

The package exposes `visio_inspect` and `visio_edit` for `.vsdx` files.
`visio_edit` supports core-validated plain text replacement, move, resize, delete
and rectangle creation. Coordinates use inches with a bottom-left origin.
There is no standalone create-document tool. Unsupported edits reject; inspect
and edit results include diagnostics about fidelity limits.

## Documentation

[Source and tools](https://github.com/ChristopherVR/visio-viewer/tree/main/mcp) |
[Core automation](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md)

## License

Apache-2.0.
