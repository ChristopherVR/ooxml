# MCP ownership and release order

The standalone packages belong to their viewer repositories:

| Repository   | Package          | CLI         |
| ------------ | ---------------- | ----------- |
| docx-viewer  | docx-viewer-mcp  | docx-tools  |
| xlsx-viewer  | xlsx-viewer-mcp  | xlsx-tools  |
| visio-viewer | visio-viewer-mcp | visio-tools |
| pptx-viewer  | pptx-viewer-mcp  | pptx-tools  |
| ooxml        | ooxml-mcp        | ooxml-tools |

The new format packages own tool names, input schemas, annotations, registration
and the transport. They call `ooxml-core/automation` for headless document
operations, and `ooxml-core/automation/node` for scoped local file execution.
No editor, framework, XML parser or document mutation is copied into those tools.
Core imports neither MCP SDK nor viewer packages at runtime.

The OOXML package composes the standalone packages' `registerTools` exports.
It supports all four formats by default, or an explicit subset. A selected package
that is missing or incompatible fails startup with an installation message.
Importing a package never starts a stdio transport; each CLI does so explicitly.

Initial coverage: DOCX block/run inspection, document creation and ordinary run
text editing; XLSX sheet inspection, workbook creation, bounded range reads and
batch cell inputs with core recalculation; VSDX scene inspection and core-supported
atomic text and conservative geometry edits. Visio inspection omits binary image
payloads. Unsupported edits reject and core warnings/diagnostics are returned.
These APIs do not imply Word, Excel or Visio parity or lossless export.

The existing PowerPoint MCP keeps its established tool set and names. Its server
exports registration and accepts a root directory for composition. Its model
operations, execution pipeline and collaboration codec live in
`ooxml-core/pptx/automation`; viewer-side compatibility modules delegate to core.
Schemas, registration and transport remain in the PowerPoint repository.

Release core first (the automation entry requires 0.11.0 or later), then the new
standalone packages and the PowerPoint version with registration (2.6.0+), then
the combined package. These are independent conventional-commit releases.
The new packages ship JavaScript source and declarations, so no build is required.
After that core release, refresh the PowerPoint workspace lockfile with `bun install`
before landing its dependency bumps. Its current registry lock cannot resolve the
new automation entry before that release. Temporary local core links are for
verification only and must not be committed.
Tests are excluded from npm tarballs. Install a standalone package's dependencies
and run `npm test` inside its `mcp/` directory to exercise protocol contracts.
Core round-trip, edit and filesystem tests live in `src/automation/`.

See [combined server usage](../mcp/README.md) for installation and selection.
