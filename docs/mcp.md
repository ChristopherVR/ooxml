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

XLSX inspection reports, per sheet, its state (`visible`, `hidden`, `veryHidden`), its used
range (the cells it stores, formatting included) with row and column counts, and its tables
(name, range, header and totals rows); it also reports the active sheet and the defined names
(hidden names are left out; a sheet-scoped name carries `sheetIndex`). Tables per sheet and
names are capped at 500 each, and `tablesOmitted` and `definedNamesOmitted` count the rest.
A range read returns, for each cell, its value and formula, the text as the grid displays it
(the cell's number format applied by the core formatter), the number format code (absent for
General) and, for date and time formats, the value in ISO 8601 without a time zone (`2024-04-03`,
`10:30:00` or `2024-04-03T10:30:00`; durations such as `[h]:mm` have none).
`findXlsxCells` searches the way Excel's Find does (formulas as typed or values as displayed,
match case, whole cell, `*` `?` `~` wildcards, optionally one sheet or range) and returns at most
1,000 matches, each with its sheet, address and the same cell view, plus the total count.
`readXlsxTable` reads a table by name (ignoring case, as Excel does): its sheet, range, header
and totals rows, column names and cells (at most 10,000, as a range read).

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
Core round-trip, edit and filesystem tests live in `src/core/automation/`.

See [combined server usage](../mcp/README.md) for installation and selection.
