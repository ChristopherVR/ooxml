# docx-viewer-mcp

**An MCP server that lets AI assistants inspect, create, edit and fill Word documents on disk, with every change made by the same core as the Word editors.**

[![npm version](https://img.shields.io/npm/v/docx-viewer-mcp.svg)](https://www.npmjs.com/package/docx-viewer-mcp)
[![license](https://img.shields.io/npm/l/docx-viewer-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE)

[Documentation](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/mcp/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

## Install

```bash
npm install docx-viewer-mcp
```

Or run it without installing through `npx` (below). It needs Node.js 22.3 or newer, on Windows, macOS or Linux, and ships plain JavaScript with type declarations.

## Client configuration

The CLI is `docx-tools`. Add it to any MCP client that starts stdio servers:

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

`--root <dir>` (or `--root=<dir>`, or a bare first argument) is the only flag: the folder the tools may read and write. Without it the server uses the current directory and says so on stderr.

To add the Word tools to a server of your own, use the library entry; importing it starts no transport:

```ts
import { createServer, registerTools } from 'docx-viewer-mcp';

const server = createServer({ rootDir: '/path/to/documents' }); // a server of its own
registerTools(existingServer, { rootDir: '/path/to/documents' }); // or on a server you already have
```

The combined [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp) server calls this same `registerTools`.

## Tools

| Group  | Tool                    | What it does                                                                                                                                     |
| ------ | ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| Read   | `docx_inspect`          | Lists the document's blocks with stable paragraph ids, runs, properties and warnings.                                                            |
| Read   | `docx_list_form_fields` | Lists content controls and legacy form fields with their ids, kinds, values and choices, and any editing restriction.                            |
| Create | `docx_create`           | Creates a document with one paragraph per string.                                                                                                |
| Edit   | `docx_set_run_text`     | Sets the text of one ordinary run and keeps its formatting.                                                                                      |
| Edit   | `docx_suggest_run_text` | Proposes new text for runs as tracked changes (with an optional `author`), which the user accepts or rejects in Word.                            |
| Edit   | `docx_fill_form_fields` | Fills form fields by id: text, a checkbox, a choice or a `YYYY-MM-DD` date. If any value is refused nothing is written, unless `partial` is set. |

Every tool takes a `filePath` under the root; edit tools also take an optional `outputPath`. The tools, their schemas and their bounds are `docxTools` in `ooxml-core/automation`, shared with the Office suite's assistant and any host backend, and their schemas are versioned (see [MCP tools](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md) and the [stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md#mcp-schema-versioning)).

## Safety and limits

- **Confined to the root.** Every path must resolve inside the root, links and real paths included, and must end in `.docx`. UNC and device paths and paths starting with a backslash are refused everywhere; on Windows, so are names Windows would reinterpret (reserved characters, trailing dots or spaces, device names such as `NUL`, and drive-relative forms such as `C:x`).
- **Never overwritten by accident.** `docx_create` never replaces an existing file. Edit tools update the source unless `outputPath` names another `.docx`, and an edit that changes nothing writes nothing.
- **Written atomically.** Output goes to a temporary file that replaces the target only once it is complete. Files over 64 MiB are refused.
- **Checked by the core.** Documents open through the core's guarded zip reader and size limits, edits the core cannot save safely are refused, and unsupported features come back as warnings.
- **Narrow on purpose.** The tools cover inspection, run text, tracked-change suggestions and form filling; they do not lay out pages or edit tables, styles or drawings. A run that holds a field, a picture, an equation or a revision is refused rather than rewritten. Saving keeps the parts the model does not cover, but it is not a lossless round trip for every feature Word can write; see the [preservation guarantee](https://christophervr.github.io/ooxml/docx/preservation).

## Privacy and network use

The server makes no network request of its own: it reads and writes files under its root only and talks to the MCP client over stdio. On Windows, a path naming a network share can make the system contact that host while the path is checked, before it is refused. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#mcp-servers-mcp-and-the-viewers-mcp-packages).

## Documentation

- [MCP tools](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md): every server, tool and bound.
- [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp): one server with the Word, Excel, PowerPoint and Visio tools.
- [`docx-core`](https://www.npmjs.com/package/docx-core): the Word model, for code that edits documents directly.
- The Word editors: [`docx-react-viewer`](https://www.npmjs.com/package/docx-react-viewer), [`docx-vue-viewer`](https://www.npmjs.com/package/docx-vue-viewer), [`docx-angular-viewer`](https://www.npmjs.com/package/docx-angular-viewer), [`docx-svelte-viewer`](https://www.npmjs.com/package/docx-svelte-viewer), [`docx-solid-viewer`](https://www.npmjs.com/package/docx-solid-viewer) and [`docx-vanilla-viewer`](https://www.npmjs.com/package/docx-vanilla-viewer), in the [ooxml repository](https://github.com/ChristopherVR/ooxml#readme).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/LICENSE). See [NOTICE](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/NOTICE) for attributions.
