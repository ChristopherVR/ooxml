# visio-viewer-mcp

**An MCP server that lets an AI agent inspect `.vsdx` drawings and apply validated edits, confined to one folder on your machine.**

[![npm version](https://img.shields.io/npm/v/visio-viewer-mcp.svg)](https://www.npmjs.com/package/visio-viewer-mcp)
[![license](https://img.shields.io/npm/l/visio-viewer-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE)
[![types](https://img.shields.io/npm/types/visio-viewer-mcp.svg)](https://www.npmjs.com/package/visio-viewer-mcp)

[MCP guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/mcp/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

## Install

```bash
npm install visio-viewer-mcp
```

`ooxml-core` and `@modelcontextprotocol/sdk` are installed with it. The package ships plain JavaScript and needs no build step. Node.js 22.3 or newer.

## Client configuration

The package's command is `visio-tools`. It serves over stdio and takes the folder it may use as `--root <dir>` (or `--root=<dir>`, or a bare first argument). Without one it uses the current directory and says so on stderr.

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

To use the tools from your own Node code, `createServer({ rootDir })` returns a new server with them, and `registerTools(server, { rootDir })` adds them to a server you already have. Importing the package starts no transport; connect one yourself.

```js
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createServer, registerTools } from 'visio-viewer-mcp';

const server = createServer({ rootDir: '/path/to/drawings' });

const mine = new McpServer({ name: 'my-tools', version: '1.0.0' });
registerTools(mine, { rootDir: '/path/to/drawings' });
```

Import `@modelcontextprotocol/sdk` only if your project declares it as its own dependency.

## Tools

| Tool            | What it does                                                                                                                                                                                      |
| --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `visio_inspect` | Reads a `.vsdx` file: pages, shape ids, text, geometry and the core's diagnostics. Never writes.                                                                                                  |
| `visio_edit`    | Applies 1 to 1,000 edits as one atomic batch: `replace-plain-text`, `move-shape`, `resize-shape`, `delete-shape` and `create-rectangle`. Coordinates are inches from the bottom left of the page. |

Each tool takes a `filePath` under the root. `visio_edit` replaces the source atomically unless `outputPath` names a new file with the same extension; it never overwrites an existing file there. Results and errors come back as JSON text, with diagnostics that describe any fidelity limits. There is no tool that creates a drawing from nothing.

The tools are `visioTools` in `ooxml-core/automation`, registered with `registerAutomationTools` from `ooxml-core/automation/node`; the combined [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp) server and the Office suite's assistant run the same definitions. See the [MCP guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md).

## Safety and limits

- **Confined to the root.** Paths are checked as text before any filesystem call, then again after links are resolved. UNC and device paths are refused, and on Windows so are names the system would reinterpret (`:`, `<`, `>`, `"`, `|`, `?`, `*`, a trailing dot or space, device names such as `NUL`).
- **Bounded.** Files are limited to 64 MiB in and out, and every read or edit has one memory budget sized from the heap Node reports; a drawing that does not fit is refused with `LIMIT_XML_MEMORY`. A page of 10,000 shapes needs a heap limit of about 1,000 MiB, so start Node with `--max-old-space-size=820` or more for drawings that large.
- **Validated and atomic.** The core checks every edit in a batch before anything is written; unsupported edits are refused. Output goes to a temporary file and is moved into place only on success.
- **Beta.** Tool names and schemas change only with a changelog note. Edits are experimental: a saved drawing is not yet guaranteed to keep everything the file contains ([capability ledger](https://christophervr.github.io/ooxml/visio/parity)).

## Privacy and network use

The server makes no network request of its own: it reads and writes only files under its root, answers over stdio and collects no telemetry. On Windows, a path naming a network share can make the system contact that host while the path is checked, before it is refused. What your MCP client does with tool results is up to that client. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#mcp-servers-mcp-and-the-viewers-mcp-packages).

## Documentation

- [MCP guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) for every server in the repository.
- [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp): one server with the Word, Excel, PowerPoint and Visio tools.
- [`visio-core`](https://www.npmjs.com/package/visio-core): the Visio document API for your own Node code.
- The [Visio viewers](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) show and edit the same drawings in a browser.

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/LICENSE).
