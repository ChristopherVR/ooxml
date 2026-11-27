# pptx-viewer-mcp

**An [MCP](https://modelcontextprotocol.io) server with 74 PowerPoint tools that let an AI agent read, edit, validate and export `.pptx` files under a folder you choose, plus the same tools as plain TypeScript functions.**

[![npm version](https://img.shields.io/npm/v/pptx-viewer-mcp.svg)](https://www.npmjs.com/package/pptx-viewer-mcp)
[![license](https://img.shields.io/npm/l/pptx-viewer-mcp.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE)
[![types](https://img.shields.io/npm/types/pptx-viewer-mcp.svg)](https://www.npmjs.com/package/pptx-viewer-mcp)

[Documentation](https://christophervr.github.io/ooxml/pptx/packages/mcp) · [Changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/tools/CHANGELOG.md) · [Report a bug](https://github.com/ChristopherVR/ooxml/issues)

The tools are `pptxTools` from `ooxml-core/pptx/automation`, registered with the same `registerAutomationTools` call as the Word, Excel and Visio servers, so the PowerPoint viewers' built-in assistant runs the same code.

## Install

Nothing to install for an MCP client: it runs the server with `npx`. A global install (`npm install -g pptx-viewer-mcp`) provides the same server as the `pptx-tools` command. To use the tools from your own code:

```bash
npm install pptx-viewer-mcp
npm install yjs # only for the collaboration codec
```

## Client configuration

Point your MCP client (Claude Desktop, Claude Code, Cursor or any client that launches stdio servers) at the server:

```json
{
	"mcpServers": {
		"pptx": {
			"command": "npx",
			"args": ["pptx-viewer-mcp", "--root", "/path/to/decks"]
		}
	}
}
```

The only flag is `--root <dir>` (or `--root=<dir>`). Without it the server uses the `PPTX_TOOLS_ROOT` environment variable, else the working directory, with a warning on stderr. To host the tools in your own server or transport, use `createServer({ rootDir })` or `registerTools(server, { rootDir })` from `pptx-viewer-mcp/mcp`.

## Tools

Every tool takes a `filePath` relative to the root.

| Group            | Tools                                                                                                                                                                                                                                                                                                                                 |
| ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Slides           | `get_slide` `add_slide` `delete_slides` `reorder_slides` `duplicate_slide` `update_slide_properties` `set_slide_transition` `set_canvas_size`                                                                                                                                                                                         |
| Elements         | `add_element` `update_element` `rename_element` `delete_elements` `arrange_elements` `clone_element` `set_element_animation` `group_elements` `ungroup_elements` `batch_update_elements` `set_element_lock` `replace_geometry` `update_element_style`                                                                                 |
| Tables           | `update_table_cells` `manage_table_structure` `set_table_style_section` `create_table_style` `delete_table_style` `assign_table_style`                                                                                                                                                                                                |
| Text and review  | `find_text` `replace_text` `manage_comments` `manage_hyperlinks` `run_accessibility_check`                                                                                                                                                                                                                                            |
| Charts           | `create_chart` `update_chart` `add_chart_series` `remove_chart_series` `update_chart_series_data` `format_chart_data_point` `format_chart_data_label` `format_chart_series` `set_chart_helper_line` `set_chart_color_map_override` `chart_user_shape_list` `chart_user_shape_add` `chart_user_shape_update` `chart_user_shape_remove` |
| Themes, layouts  | `get_theme_info` `apply_theme_preset` `update_theme_colors` `update_theme_fonts` `get_layouts` `apply_layout` `find_placeholders` `apply_template`                                                                                                                                                                                    |
| Deck             | `get_metadata` `update_metadata` `get_presentation_properties` `update_presentation_properties` `manage_sections` `manage_custom_shows` `manage_smart_art` `validate_presentation` `repair_presentation`                                                                                                                              |
| Embedded objects | `ole_get_content` `ole_set_sheet_cell` `ole_set_document_paragraph` `ole_set_deck_slide_title` `ole_replace_file` `ole_set_object_name`                                                                                                                                                                                               |
| Export           | `convert_to_markdown` `export_to_svg` `export_slide_svg` `export_to_json` `import_from_json`                                                                                                                                                                                                                                          |

Each tool is also a camelCase function in the package root (`replaceText`, `addSlide`, `getOleContent`; some carry a `T` suffix, such as `addChartSeriesT`, `setElementLockT` and `findPlaceholdersT`): deck data in, changed data out, no file access of its own. `PptxHandler` is re-exported, so one install loads and saves too:

```ts
import { readFile, writeFile } from 'node:fs/promises';

import { PptxHandler, replaceText } from 'pptx-viewer-mcp';

const file = await readFile('deck.pptx');
const handler = new PptxHandler();
const pptxData = await handler.load(
	file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength),
);

const {
	pptxData: updated,
	dirty,
	result,
} = replaceText({ pptxData }, { query: 'Draft', replacement: 'Final' });
if (dirty) await writeFile('deck.pptx', await handler.save(updated.slides));
console.log(result.replacementCount);
```

`executeToolWithContext(filePath, context, tool)` runs the load, tool and save cycle through a `FileSystemProvider` you supply, or writes into a live Y.Doc through a `collaboration` provider. `pptx-viewer-mcp/schemas` has every tool's Zod input schema and `pptx-viewer-mcp/codec` the `PptxCodec` between a Y.Doc and `.pptx` bytes. The [tools reference](https://christophervr.github.io/ooxml/pptx/packages/mcp) gives each tool's parameters.

## Safety and limits

- Every path is confined to the root and judged on its text before any filesystem call: `..` that leaves the root, UNC and device paths and, on Windows, drive-relative paths, alternate data streams, device names such as `NUL` and segments ending in a dot or space are refused with `Path is outside the configured root`. Links are resolved and checked against the root too.
- Reads accept `.pptx` and `.ppt`; edits write `.pptx` only. An edit that changed nothing writes nothing.
- Files over 64 MiB are refused, on read and on write.
- `convert_to_markdown` takes a relative `outputDir` only and writes nothing.

## Privacy and network use

The server talks to its client over stdio and makes no network request of its own. It reads and writes files under its root, and does not fetch a deck's linked pictures. On Windows a path naming a network share can make the system contact that host while the path is checked, before it is refused. Nothing collects analytics or telemetry. See [PRIVACY.md](https://github.com/ChristopherVR/ooxml/blob/main/PRIVACY.md#mcp-servers-mcp-and-the-viewers-mcp-packages).

## Documentation

- The [tools reference](https://christophervr.github.io/ooxml/pptx/packages/mcp) and [how the MCP servers share their tools](https://github.com/ChristopherVR/ooxml/blob/main/docs/mcp.md). Node.js 22.3 or newer.
- [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp) serves the Word, Excel, PowerPoint and Visio tools from one server; [`pptx-viewer-core`](https://www.npmjs.com/package/pptx-viewer-core) is the engine on its own.
- The [PowerPoint viewers](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme), whose assistant panel drives these tools; part of [ooxml](https://github.com/ChristopherVR/ooxml#readme).

## License

[Apache-2.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/LICENSE).
