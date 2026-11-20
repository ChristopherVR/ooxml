# visio-viewer-mcp

Give an AI agent safe, local access to Visio drawings: an MCP server that inspects `.vsdx` files and applies validated edits, confined to one folder.

[![npm version](https://img.shields.io/npm/v/visio-viewer-mcp.svg)](https://www.npmjs.com/package/visio-viewer-mcp)
[![license](https://img.shields.io/npm/l/visio-viewer-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE)
[![types](https://img.shields.io/npm/types/visio-viewer-mcp.svg)](https://www.npmjs.com/package/visio-viewer-mcp)
[![status: beta](https://img.shields.io/badge/status-beta-orange.svg)](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)

[Docs](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/mcp/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

> **Beta.** The tool names and schemas change only with a changelog note. Edits are experimental: they are validated before anything is written, but a saved drawing is not yet guaranteed to keep everything the file contains.

## Install

```bash
npm install visio-viewer-mcp
```

`ooxml-core` is installed with it. The package ships plain JavaScript and needs no build step.

## Quick start

Add the server to your MCP client, with the folder it may use:

```json
{
	"mcpServers": {
		"visio": {
			"command": "npx",
			"args": ["-y", "--package=visio-viewer-mcp", "visio-tools", "--root", "/path/to/drawings"]
		}
	}
}
```

The `visio-tools` command serves over stdio. Without `--root` (or a bare folder argument) it uses the current directory and says so on stderr.

## Tools

| Tool            | What it does                                                                                                                                                                                                                          |
| --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `visio_inspect` | Reads a `.vsdx` file: pages, shape ids, text, geometry and the core's diagnostics. Never writes.                                                                                                                                      |
| `visio_edit`    | Applies a batch of 1 to 1,000 edits atomically: `replace-plain-text`, `move-shape`, `resize-shape`, `delete-shape` and `create-rectangle`. Coordinates are in inches from the bottom-left of the page. Unsupported edits are refused. |

Each tool takes a `filePath` under the root. `visio_edit` replaces the source atomically unless `outputPath` names a new file with the same extension; it never overwrites an existing file there. Results and errors come back as JSON text, with diagnostics that describe any fidelity limits. There is no tool that creates a drawing from nothing.

## Safety

- **Confined to the root.** Paths are checked on their text before any filesystem call, then again after links are resolved. UNC and device paths are refused everywhere, and on Windows so are names the system would reinterpret (`:`, `<`, `>`, `"`, `|`, `?`, `*`, a trailing dot or space, device names such as `NUL`).
- **Bounded.** Files are limited to 64 MiB in and out, and every read or edit has one memory budget sized from the heap Node reports. A drawing that does not fit is refused with `LIMIT_XML_MEMORY`. A page of 10,000 shapes needs a heap limit of about 1,000 MiB, so start Node with `--max-old-space-size=820` or more if your agents open drawings that large.
- **Atomic writes.** Output is written to a temporary file and moved into place only after the core has validated every edit in the batch.

## Use it from code

```js
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createServer, registerTools } from 'visio-viewer-mcp';

// A new server with the Visio tools:
const server = createServer({ rootDir: '/path/to/drawings' });

// Or add them to a server you already have:
const mine = new McpServer({ name: 'my-tools', version: '1.0.0' });
registerTools(mine, { rootDir: '/path/to/drawings' });
```

Importing the package starts no transport; connect the server to one yourself. Install `@modelcontextprotocol/sdk` in your own project when you import it. The tools themselves are `visioTools` in `ooxml-core/automation`, registered with `registerAutomationTools` from `ooxml-core/automation/node`; the combined [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp) server registers the same definitions. See the [MCP guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md).

## Supported environments

- Node.js 22.3 or newer.
- Any MCP client that starts stdio servers.

## Privacy and network use

The server makes no network request. It reads and writes only files under its root, sends nothing anywhere and collects no telemetry; what your MCP client does with tool results is up to that client. It does not sign in to Microsoft or sync OneDrive: an application that stores drawings in OneDrive can fetch them with `ooxml-core/embedding/onedrive` and hand the bytes to its own tool workflow ([OneDrive guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md)). See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

| Package                                                                                                                                                                       | What it is                                                   |
| ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------ |
| [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp)                                                                                                                        | One server with the Word, Excel, PowerPoint and Visio tools. |
| [`visio-core`](https://www.npmjs.com/package/visio-core)                                                                                                                      | The Visio document API, for your own Node code.              |
| [`visio-react-viewer`](https://www.npmjs.com/package/visio-react-viewer) and the other [Visio viewers](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) | Show and edit the same drawings in a browser.                |
| [`ooxml-core`](https://www.npmjs.com/package/ooxml-core)                                                                                                                      | The Office logic behind the tools.                           |

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE).
