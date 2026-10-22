# docx-viewer-mcp

[![npm version](https://img.shields.io/npm/v/docx-viewer-mcp.svg)](https://www.npmjs.com/package/docx-viewer-mcp)
[![license](https://img.shields.io/npm/l/docx-viewer-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE)
[![types](https://img.shields.io/npm/types/docx-viewer-mcp.svg)](https://www.npmjs.com/package/docx-viewer-mcp)

> Headless MCP access to document operations owned by OOXML core.

[Live demo](https://christophervr.github.io/ooxml/docx/demo/) | [npm](https://www.npmjs.com/package/docx-viewer-mcp) | [Full docs](https://christophervr.github.io/ooxml/docx/) | [Source](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; **[Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme)** &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; [OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme) &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

A thin MCP wrapper: the tools themselves (names, schemas, bounds and operations) are
`docxTools` in `ooxml-core/automation`, registered with `registerAutomationTools` from
`ooxml-core/automation/node`. Requires `ooxml-core` 1.22.0 or later, installed as a dependency.

## Install

```bash
npm install docx-viewer-mcp
```

## Quick start

```json
{
	"mcpServers": {
		"docx": {
			"command": "npx",
			"args": ["-y", "--package=docx-viewer-mcp", "docx-tools", "/path/to/documents"]
		}
	}
}
```

Paths are scoped to the configured root (the current directory by default).
Creation and save-as never overwrite existing files. Edit tools update the source
unless `outputPath` names a new file. Files are written atomically after core
validation. Unsupported features and preservation limitations are returned by core.

## Features

Word inspection, form field listing and filling, document creation, supported run text
replacement and the same replacement suggested as tracked changes. All
document algorithms live in OOXML core. Importing the package does not start stdio.

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

[Source and tools](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx/mcp) |
[Core automation](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md)

## License

Apache-2.0.
