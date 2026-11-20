# docx-viewer-mcp

**An MCP server that lets AI assistants inspect, create, edit and fill Word documents on disk, with every change made by the same core as the Word editors.**

[![npm version](https://img.shields.io/npm/v/docx-viewer-mcp.svg)](https://www.npmjs.com/package/docx-viewer-mcp)
[![license](https://img.shields.io/npm/l/docx-viewer-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE)

[Documentation](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/mcp/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

## Install

Add the server to your MCP client (Claude Desktop, an IDE, or any client that starts stdio servers):

```json
{
	"mcpServers": {
		"docx": {
			"command": "npx",
			"args": ["-y", "--package=docx-viewer-mcp", "docx-tools", "--root", "/path/to/documents"]
		}
	}
}
```

Or install it into a project with `npm install docx-viewer-mcp` and run `docx-tools --root <dir>`. Without `--root` (a bare first argument also works) the server uses the current directory and says so on stderr.

## Tools

| Tool                    | What it does                                                                                                                                     |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `docx_inspect`          | Lists the document's blocks with stable paragraph ids, runs, properties and warnings.                                                            |
| `docx_list_form_fields` | Lists content controls and legacy form fields with their ids, kinds, values and choices, and any editing restriction.                            |
| `docx_create`           | Creates a document with one paragraph per string.                                                                                                |
| `docx_set_run_text`     | Sets the text of one ordinary run and keeps its formatting. Fields, pictures, equations and revisions are refused.                               |
| `docx_suggest_run_text` | Proposes new text for runs as tracked changes, which the user accepts or rejects in Word.                                                        |
| `docx_fill_form_fields` | Fills form fields by id: text, a checkbox, a choice or a `YYYY-MM-DD` date. If any value is refused nothing is written, unless `partial` is set. |

The tools, their schemas and their bounds are `docxTools` in [`ooxml-core/automation`](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md), shared with the Office suite's assistant and any host backend, so the server and the editors make the same edits.

## How files are handled

- **Confined to the root.** Every path must resolve inside the root, through links and real paths included. Paths are checked on their text before any filesystem call: UNC and device paths and paths starting with a backslash are refused everywhere, and on Windows so are names Windows would reinterpret (reserved characters, trailing dots or spaces, device names such as `NUL`, and drive-relative forms such as `C:x`).
- **Never overwritten by accident.** `docx_create` never replaces an existing file. Edit tools update the source unless `outputPath` names another `.docx`, and an edit that changes nothing writes nothing.
- **Written atomically.** Output goes to a temporary file that replaces the target only once it is complete. Files over 64 MiB are refused.
- **Checked by the core.** Documents are opened through the core's guarded zip reader and size limits, edits that the core cannot save safely are refused, and unsupported features come back as warnings.

## Use it from code

```ts
import { createServer, registerTools } from 'docx-viewer-mcp';

// A server of its own:
const server = createServer({ rootDir: '/path/to/documents' });

// Or add the Word tools to a server you already have:
registerTools(existingServer, { rootDir: '/path/to/documents' });
```

Importing the package starts no transport; connect the server to the transport you use. The combined [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp) server calls this same `registerTools`. The package ships plain JavaScript with type declarations and needs no build step.

## Known limits

The tools cover inspection, run text, tracked-change suggestions and form filling; they do not lay out pages or edit tables, styles or drawings. A run that holds a field, a picture, an equation or a revision is refused rather than rewritten. Saving keeps the parts the model does not cover, but it is not a lossless round trip for every feature Word can write; see the [preservation guarantee](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/docs/preservation.md).

## Supported environments

- Node.js 22.3 or newer, on Windows, macOS and Linux.
- Any MCP client that starts stdio servers.

## Privacy and network use

The server reads and writes files under its root only and makes no network request of its own; it talks to the MCP client over stdio. Nothing collects analytics or telemetry. For OneDrive, an application you authorise can fetch or store the bytes with `ooxml-core/embedding/onedrive` and hand them to its own document workflow; this package acquires no Microsoft credentials and syncs nothing (see the [OneDrive integration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/onedrive-integration.md)). See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md).

## Related packages

- [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp): one server with the Word, Excel, PowerPoint and Visio tools.
- [`docx-core`](https://www.npmjs.com/package/docx-core): the Word model, for code that edits documents directly.
- The Word editors: [`docx-react-viewer`](https://www.npmjs.com/package/docx-react-viewer), [`docx-vue-viewer`](https://www.npmjs.com/package/docx-vue-viewer), [`docx-angular-viewer`](https://www.npmjs.com/package/docx-angular-viewer), [`docx-svelte-viewer`](https://www.npmjs.com/package/docx-svelte-viewer), [`docx-solid-viewer`](https://www.npmjs.com/package/docx-solid-viewer) and [`docx-vanilla-viewer`](https://www.npmjs.com/package/docx-vanilla-viewer).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE).
