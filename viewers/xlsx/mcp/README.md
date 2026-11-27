# xlsx-viewer-mcp

**An MCP server that lets AI agents inspect, read, create and edit Excel workbooks on disk, confined to one folder, with bounded reads and reviewable edits.**

[![npm version](https://img.shields.io/npm/v/xlsx-viewer-mcp.svg)](https://www.npmjs.com/package/xlsx-viewer-mcp)
[![license](https://img.shields.io/npm/l/xlsx-viewer-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/LICENSE)
[![types](https://img.shields.io/npm/types/xlsx-viewer-mcp.svg)](https://www.npmjs.com/package/xlsx-viewer-mcp)

[MCP guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) · [Documentation](https://christophervr.github.io/ooxml/xlsx/) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

The tools (names, schemas, bounds and operations) are `xlsxTools` in `ooxml-core/automation`, the same definitions the combined [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp) server and the OOXML Office assistant use. This package registers them on an MCP server and ships the `xlsx-tools` stdio command.

## Install

```bash
npm install xlsx-viewer-mcp
```

Or let your MCP client run it with `npx`, as below. Node.js 22.3 or newer, on Windows, macOS and Linux; the package is plain JavaScript and needs no build step.

## Client configuration

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

The command is `xlsx-tools`. Its one option is `--root <dir>` (or `--root=<dir>`, or the folder as a bare first argument): the folder the tools may read and write. Without it the server uses the current working directory and says so on stderr.

From code, `createServer({ rootDir })` returns a ready server and `registerTools(server, { rootDir })` adds the Excel tools to a server you already have. Importing the package never starts a transport; add `@modelcontextprotocol/sdk` to your own dependencies when you use it:

```ts
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer, registerTools } from 'xlsx-viewer-mcp';

await createServer({ rootDir: '/path/to/workbooks' }).connect(new StdioServerTransport());

// Or, on your own server:
const server = new McpServer({ name: 'my-tools', version: '1.0.0' });
registerTools(server, { rootDir: '/path/to/workbooks' });
```

## Tools

Every tool takes a `filePath` (relative to the root, or absolute inside it); edit tools also take an optional `outputPath`. Reads accept `.xlsx` and `.xlsm`.

| Tool                     | Kind   | What it does                                                                                                                              |
| ------------------------ | ------ | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `xlsx_inspect`           | read   | Sheets (state, used range and size, tables), the active sheet, defined names, properties and warnings. Start here.                        |
| `xlsx_read_range`        | read   | Value, formula, displayed text and number format of each cell in an A1 range (at most 10,000 cells).                                      |
| `xlsx_read_filled_cells` | read   | The filled cells of one sheet or all, up to `maxCells` (default 2,000, at most 10,000), with omitted counts.                              |
| `xlsx_find`              | read   | Searches as Excel's Find does (wildcards on by default); at most `maxResults` matches (default 100, up to 1,000) and the total.           |
| `xlsx_read_table`        | read   | A table by name: sheet, range, header and totals rows, column names and cells (at most 10,000 cells).                                     |
| `xlsx_create`            | create | Creates an `.xlsx` with the named sheets (up to 100). Never overwrites a file.                                                            |
| `xlsx_set_cells`         | edit   | Writes up to 10,000 cell inputs and recalculates.                                                                                         |
| `xlsx_propose_edits`     | read   | Returns the diff of a list of edits (cells, clear, format, merge, rows and columns, sheets, names) and a sealed proposal, without saving. |
| `xlsx_apply_proposal`    | edit   | Applies an approved proposal; refused, changing nothing, if the workbook or the proposal changed in between.                              |

Reads never return file bytes, and every tool reports the core's warnings. Each tool carries a schema version in `_meta.schemaVersion`. See the [MCP guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md).

## Safety and limits

- **Confined to the root.** Paths are judged on their text before any filesystem call, then again after resolving links. A path outside the root, a UNC or device path, and on Windows a drive-relative path, an alternate data stream or a device name such as `NUL` all get the same error.
- **No accidental overwrites.** `xlsx_create` never overwrites. Edits replace the source atomically after validation, unless `outputPath` names another file.
- **Formulas stay inside the workbook.** The edit tools refuse a formula that could fetch data or open links (`WEBSERVICE`, `IMAGE`, `HYPERLINK`, `RTD`, cube functions), call an AI service (`COPILOT`, `TRANSLATE`), run a macro function or call an add-in, or refer to another file. Defined names are checked too, and no tool parameter turns the check off. A host whose callers are all trusted can register `createXlsxTools({ formulas: 'unrestricted' })` from `ooxml-core/automation` with `registerAutomationTools` from `ooxml-core/automation/node` instead.
- **Bounded work.** Files over 64 MiB are refused, workbooks are read through a guarded zip reader and calculated under a per-formula work budget, so a crafted file cannot exhaust memory or stall a call.
- **Not all of Excel.** Only the operations above exist; anything else is refused. Saving is not lossless for features the model does not cover; see the [preservation guarantee](https://christophervr.github.io/ooxml/xlsx/preservation).

The tool names, inputs and kinds are part of the Stable API under the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md).

## Privacy and network use

The server makes no network request of its own: it reads and writes files under its root and answers over stdio. On Windows, a path naming a network share can make the system contact that host while the path is checked, before it is refused. Because formulas that would fetch data, open links or call an AI service are refused, a workbook the tools write does not contact anything when it is later opened in a spreadsheet application. The server does not sign in to Microsoft or sync OneDrive; an application with access can fetch the bytes through `ooxml-core/embedding/onedrive` and hand the tools a local file ([OneDrive integration](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md)). Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#mcp-servers-mcp-and-the-viewers-mcp-packages).

## Documentation

- The [MCP guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) and the [Excel documentation](https://christophervr.github.io/ooxml/xlsx/).
- [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp): one server with the Word, Excel, PowerPoint and Visio tools.
- [`@christophervr/xlsx-core`](https://www.npmjs.com/package/@christophervr/xlsx-core): the workbook model for your own headless code.
- The Excel editor for [React](https://www.npmjs.com/package/@christophervr/xlsx-react-viewer), [Vue](https://www.npmjs.com/package/xlsx-vue-viewer), [Angular](https://www.npmjs.com/package/xlsx-angular-viewer), [Svelte](https://www.npmjs.com/package/xlsx-svelte-viewer), [Solid](https://www.npmjs.com/package/xlsx-solid-viewer) and [plain JavaScript](https://www.npmjs.com/package/xlsx-vanilla-viewer).
- Source: the [ooxml repository](https://github.com/ChristopherVR/ooxml).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/LICENSE).
