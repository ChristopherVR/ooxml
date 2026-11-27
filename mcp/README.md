<div align="center">

# ooxml-mcp

**One MCP server that lets an AI agent read and edit Word, Excel, PowerPoint and Visio files on your machine.**

[![npm version](https://img.shields.io/npm/v/ooxml-mcp.svg)](https://www.npmjs.com/package/ooxml-mcp)
[![license](https://img.shields.io/npm/l/ooxml-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/LICENSE)
[![types](https://img.shields.io/npm/types/ooxml-mcp.svg)](https://www.npmjs.com/package/ooxml-mcp)

[**Live demo**](https://christophervr.github.io/ooxml/) &nbsp;&middot;&nbsp;
[**Docs**](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) &nbsp;&middot;&nbsp;
[**Changelog**](https://github.com/ChristopherVR/ooxml/blob/main/mcp/CHANGELOG.md) &nbsp;&middot;&nbsp;
[**Report a bug**](https://github.com/ChristopherVR/ooxml/issues/new/choose)

</div>

`ooxml-mcp` gives any [Model Context Protocol](https://modelcontextprotocol.io) client the document tools of [ooxml](https://github.com/ChristopherVR/ooxml#readme) through one stdio server. It composes the four format servers (`docx-viewer-mcp`, `xlsx-viewer-mcp`, `visio-viewer-mcp` and `pptx-viewer-mcp`) and adds no logic of its own: every tool, schema and bound is defined once in [`ooxml-core/automation`](https://www.npmjs.com/package/ooxml-core), the same code the editors' assistants run.

## Install

```bash
npm install ooxml-mcp docx-viewer-mcp xlsx-viewer-mcp visio-viewer-mcp pptx-viewer-mcp
```

The format packages are optional peers: install the ones you want to expose. `ooxml-core` comes with this package, so every format server shares one copy of it.

## Client configuration

Point your MCP client at the `ooxml-tools` command and a folder of documents. Most desktop clients take this shape:

```json
{
	"mcpServers": {
		"office": {
			"command": "npx",
			"args": [
				"-y",
				"-p",
				"ooxml-mcp",
				"-p",
				"docx-viewer-mcp",
				"-p",
				"xlsx-viewer-mcp",
				"-p",
				"visio-viewer-mcp",
				"-p",
				"pptx-viewer-mcp",
				"ooxml-tools",
				"--root",
				"/path/to/documents"
			]
		}
	}
}
```

Or run it yourself:

```bash
ooxml-tools --root /path/to/documents                          # all four formats
ooxml-tools --root /path/to/documents --formats docx,xlsx      # only these
```

`--name=value` works too, and the older `ooxml-tools /path/to/documents docx,xlsx` form is still accepted. Without `--root` the server uses the current working directory and says so on stderr. A selected format whose package is missing or too old stops the server at startup with an installation message, so tools are never silently left out.

## Tools

| Format     | Tools                                                                                                                                                                                          |
| ---------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Word       | `docx_inspect`, `docx_list_form_fields`, `docx_create`, `docx_set_run_text`, `docx_suggest_run_text` (the same edit as tracked changes), `docx_fill_form_fields`                               |
| Excel      | `xlsx_inspect`, `xlsx_read_range`, `xlsx_read_filled_cells`, `xlsx_find`, `xlsx_read_table`, `xlsx_create`, `xlsx_set_cells`, `xlsx_propose_edits` and `xlsx_apply_proposal` (a reviewed edit) |
| Visio      | `visio_inspect`, `visio_edit` (an atomic batch of text and conservative geometry edits)                                                                                                        |
| PowerPoint | 74 tools under their established unprefixed names, such as `get_slide`, `add_slide` and `batch_update_elements`; reads accept `.pptx` and `.ppt`, edits write `.pptx`                          |

Excel reads are bounded (at most 10,000 cells per read and 1,000 matches per find), and `xlsx_set_cells` recalculates the workbook with the core's formula engine. The full list, bounds and result shapes are in [MCP tools](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md), and the review flow behind `xlsx_propose_edits` is in [agent edit review](https://github.com/ChristopherVR/ooxml/blob/main/docs/agent-edit-review.md).

## Safety and limits

- **Confined to one folder.** Every tool takes a path inside `--root`. Paths that leave it, UNC and device paths, and names Windows would reinterpret are refused on their text, before the disk is touched.
- **Bounded reads, careful writes.** Reads never return file bytes and cap what they return; creates never overwrite; edits replace the file atomically or write to `outputPath`.
- **Safe for untrusted prompts.** Excel writes refuse formulas that could send workbook data out of the file (web, RTD, cube, macro and AI functions, external references), and no tool parameter turns that check off.
- **Versioned schemas.** Each tool carries an integer schema version, sent to clients as `_meta.schemaVersion`, so a breaking schema change is always visible.
- **Honest results.** An unsupported edit returns an error or the core's diagnostics instead of a damaged file. These tools are not Office and claim neither Office parity nor lossless export; open gaps are in the [parity checklist](https://github.com/ChristopherVR/ooxml/blob/main/docs/parity-checklist.md).

## Use it from code

```js
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer } from 'ooxml-mcp';

const server = await createServer({ rootDir: '/path/to/documents', formats: ['docx', 'xlsx'] });
await server.connect(new StdioServerTransport());
```

`createServer` returns an MCP SDK `McpServer` with the selected registrations and never opens a transport itself. To change a policy, such as letting trusted callers write any formula with `createXlsxTools({ formulas: 'unrestricted' })` from `ooxml-core/automation`, register the tools on your own server with `registerAutomationTools` from `ooxml-core/automation/node`.

## Privacy and network use

The server talks to its client over stdio only and makes no network request of its own. It reads and writes files under its root directory, and linked pictures in a deck are not fetched. On Windows, a path naming a network share can make the system contact that host while the path is checked, before it is refused. Nothing collects analytics or telemetry. It never acquires Microsoft credentials or synchronises OneDrive; an application that holds them can fetch and save Office bytes with [`ooxml-core/embedding/onedrive`](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md) and hand them to its own tool workflow. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#mcp-servers-mcp-and-the-viewers-mcp-packages).

## Environments

Node.js 22.3 or newer on Windows, macOS and Linux. The package ships JavaScript source and declarations and needs no build step.

## Documentation

- [MCP tools](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md): every tool, its bounds and result shape, and how the servers are composed.
- [`docx-viewer-mcp`](https://www.npmjs.com/package/docx-viewer-mcp), [`xlsx-viewer-mcp`](https://www.npmjs.com/package/xlsx-viewer-mcp), [`visio-viewer-mcp`](https://www.npmjs.com/package/visio-viewer-mcp) and [`pptx-viewer-mcp`](https://www.npmjs.com/package/pptx-viewer-mcp): one format each, with their own `docx-tools`, `xlsx-tools`, `visio-tools` and `pptx-tools` commands.
- [`ooxml-core`](https://www.npmjs.com/package/ooxml-core): the library that defines every tool and owns parsing, editing, recalculation and saving.
- The editors your users open the results in: [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme), [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme), [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) and [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/LICENSE).
