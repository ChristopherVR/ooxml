# xlsx-viewer-mcp

**Excel workbooks for AI agents: an MCP server that inspects, reads, creates and edits `.xlsx` files on disk, with bounded reads and reviewable edits.**

[![npm version](https://img.shields.io/npm/v/xlsx-viewer-mcp.svg)](https://www.npmjs.com/package/xlsx-viewer-mcp)
[![license](https://img.shields.io/npm/l/xlsx-viewer-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE)
[![docs](https://img.shields.io/badge/docs-christophervr.github.io-1f9d63.svg)](https://christophervr.github.io/ooxml/xlsx/)

[Documentation](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

The tools themselves (names, schemas, bounds and operations) are `xlsxTools` in `ooxml-core/automation`, the same definitions the OOXML Office assistant and the combined `ooxml-mcp` server use. This package registers them on an MCP server, confined to one folder, and ships the `xlsx-tools` stdio command.

## Install

```bash
npm install xlsx-viewer-mcp
```

Or run it straight from an MCP client without installing.

## Quick start

Add the server to your MCP client's configuration and give it the folder it may work in:

```json
{
	"mcpServers": {
		"xlsx": {
			"command": "npx",
			"args": ["-y", "--package=xlsx-viewer-mcp", "xlsx-tools", "--root", "/path/to/workbooks"]
		}
	}
}
```

Without `--root` (or a bare folder as the first argument) the server uses the current working directory.

## Tools

| Tool                     | Kind   | What it does                                                                                                    |
| ------------------------ | ------ | --------------------------------------------------------------------------------------------------------------- |
| `xlsx_inspect`           | read   | Sheets with their state and used range, tables, the active sheet and defined names. Start here.                 |
| `xlsx_read_range`        | read   | Values, formulas, displayed text and number formats of a range (at most 10,000 cells).                          |
| `xlsx_read_filled_cells` | read   | The filled cells of one sheet or all of them, up to a cap (default 2,000, at most 10,000), with omitted counts. |
| `xlsx_find`              | read   | Searches as Excel's Find does, with wildcards and options; at most 1,000 matches plus the total.                |
| `xlsx_read_table`        | read   | A table by name: range, header and totals rows, column names and cells.                                         |
| `xlsx_create`            | create | Creates a workbook with the named sheets. Never overwrites a file.                                              |
| `xlsx_set_cells`         | edit   | Writes up to 10,000 cell inputs and recalculates.                                                               |
| `xlsx_propose_edits`     | read   | Returns the diff of a list of edits and a sealed proposal, without saving.                                      |
| `xlsx_apply_proposal`    | edit   | Applies an approved proposal, refusing if the workbook changed in between.                                      |

Every read is bounded, reads never return file bytes, and every tool reports the core's warnings.

## How it keeps files safe

- **Confined to the root.** Paths are judged on their text before any filesystem call. A path outside the root, a UNC or device path, and on Windows a drive-relative path, an alternate data stream or a device name such as `NUL` all get the same error. Use a path relative to the root, or an absolute path inside it.
- **No accidental overwrites.** Creation never overwrites. Edits replace the source atomically after validation, unless `outputPath` names a new file.
- **Formulas stay inside the workbook.** `xlsx_set_cells`, `xlsx_propose_edits` and `xlsx_apply_proposal` refuse a formula that could fetch data, run code, call an add-in, send cell data to an AI service or refer to another file when the workbook is opened in a spreadsheet application. The error names the cell, defined names are checked too, and no tool parameter turns the check off.
- **Bounded work.** Workbooks are read through a guarded zip reader and calculated under a per-formula work budget, so a crafted file cannot exhaust memory or stall a call.

## Use it from code

```ts
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer, registerTools } from 'xlsx-viewer-mcp';

// A ready server:
await createServer({ rootDir: '/path/to/workbooks' }).connect(new StdioServerTransport());

// Or add the Excel tools to a server you already have:
const server = new McpServer({ name: 'my-tools', version: '1.0.0' });
registerTools(server, { rootDir: '/path/to/workbooks' });
```

Importing the package never starts a transport. Add `@modelcontextprotocol/sdk` to your own dependencies when you import it, as above. A host whose callers are all trusted can allow every formula by registering `createXlsxTools({ formulas: 'unrestricted' })` from `ooxml-core/automation` with `registerAutomationTools` from `ooxml-core/automation/node` instead of calling `registerTools`.

## Supported environments

Node.js 22.3 or newer, on Windows, macOS and Linux. The package ships plain JavaScript and needs no build step.

## Limitations

The tools cover the operations listed above, not all of Excel. Unsupported operations are refused with an error, and preservation limits are reported in the core's warnings. Saving is not lossless for features the model does not cover; see the [preservation guarantee](https://christophervr.github.io/ooxml/xlsx/preservation).

## Privacy and network use

The server makes no network request of its own: it reads and writes files under the root it is given and answers over stdio. On Windows, a path naming a network share can make the system contact that host while the path is checked, before it is refused. Formulas that would make a spreadsheet application fetch data, open links or call an AI service are refused, so a workbook the tools write does not phone out when it is opened. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

The server does not sign in to Microsoft or sync OneDrive. An application that already has access can fetch the bytes through `ooxml-core/embedding/onedrive` and hand the tools a local file; see the [OneDrive integration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md).

## Related packages

- [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp): one server with the Word, Excel, PowerPoint and Visio tools.
- [`@christophervr/xlsx-core`](https://www.npmjs.com/package/@christophervr/xlsx-core): the workbook model for your own headless code.
- The Excel editor for [React](https://www.npmjs.com/package/@christophervr/xlsx-react-viewer), [Vue](https://www.npmjs.com/package/xlsx-vue-viewer), [Angular](https://www.npmjs.com/package/xlsx-angular-viewer), [Svelte](https://www.npmjs.com/package/xlsx-svelte-viewer), [Solid](https://www.npmjs.com/package/xlsx-solid-viewer) and [plain JavaScript](https://www.npmjs.com/package/xlsx-vanilla-viewer).
- [`ooxml-core`](https://www.npmjs.com/package/ooxml-core): the engine and the tool definitions.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/LICENSE)
