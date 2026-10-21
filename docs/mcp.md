# MCP ownership and release order

The standalone packages belong to their viewer repositories:

| Repository   | Package          | CLI         |
| ------------ | ---------------- | ----------- |
| docx-viewer  | docx-viewer-mcp  | docx-tools  |
| xlsx-viewer  | xlsx-viewer-mcp  | xlsx-tools  |
| visio-viewer | visio-viewer-mcp | visio-tools |
| pptx-viewer  | pptx-viewer-mcp  | pptx-tools  |
| ooxml        | ooxml-mcp        | ooxml-tools |

## Tools are headless functions in core

Word, Excel and Visio tools are defined once, in `ooxml-core/automation`
(`src/core/automation/tools/`): `docxTools`, `xlsxTools` and `visioTools`, also grouped as
`automationTools` and looked up with `findAutomationTool(name)`. Each definition carries its
name, description, kind (`read`, `create` or `edit`), accepted extensions, its parameters (a Zod
raw shape as `input`, and JSON Schema from `jsonSchema()`) and `run`, a function over the file's
bytes that validates its parameters itself. Reads never return bytes; edits return the new bytes
with a report (warnings, results). No parameter is a file path: where the bytes come from is the
caller's business. So the MCP servers, the suite's assistant and a host's own backend call the
same code and get the same validation and bounds.

`ooxml-core/automation/node` adds the local filesystem: `runAutomationFileTool(files, tool,
params)` runs a tool on a file confined to a root (reads never write, creates never overwrite,
edits replace the source atomically unless `outputPath` is given), and
`registerAutomationTools(server, tools, { rootDir })` registers a list on an MCP SDK server with
`filePath` (and `outputPath` for edits), the read-only and destructive hints, JSON text results
and tool errors. Core types that server structurally and never imports the MCP SDK at runtime.

The standalone packages are therefore thin: each `registerTools` is one call to
`registerAutomationTools` with its format's list, plus its stdio CLI. A new tool, a schema change
or a bound belongs in core, with its tests in `src/core/automation/`, never in a package.

The OOXML package composes the standalone packages' `registerTools` exports.
It supports all four formats by default, or an explicit subset. A selected package
that is missing or incompatible fails startup with an installation message.
Importing a package never starts a stdio transport; each CLI does so explicitly.

## Coverage

- Word: `docx_inspect` (blocks, stable paragraph ids, runs), `docx_list_form_fields`,
  `docx_create`, `docx_set_run_text` (ordinary runs only), `docx_suggest_run_text` (the same
  edit as tracked changes) and `docx_fill_form_fields`.
- Excel, reads first and every read bounded: `xlsx_inspect` (sheets and their sizes; start here),
  `xlsx_read_range` (at most 10,000 cells), `xlsx_read_filled_cells` (one sheet or all, at most
  10,000 cells, default 2,000, with per-sheet totals and omitted counts), `xlsx_find` (at most
  1,000 matches plus the total) and `xlsx_read_table` (at most 10,000 cells); then
  `xlsx_create` and `xlsx_set_cells` (at most 10,000 inputs, recalculated by the core).
- Visio: `visio_inspect` (binary image payloads omitted) and `visio_edit`, an atomic batch of
  the supported text and conservative geometry edits.

Unsupported edits reject and core warnings and diagnostics are returned. These APIs do not imply
Word, Excel or Visio parity or lossless export.

XLSX inspection reports, per sheet, its state (`visible`, `hidden`, `veryHidden`), its used
range (the cells it stores, formatting included) with row and column counts, and its tables
(name, range, header and totals rows); it also reports the active sheet and the defined names
(hidden names are left out; a sheet-scoped name carries `sheetIndex`). Tables per sheet and
names are capped at 500 each, and `tablesOmitted` and `definedNamesOmitted` count the rest.
A range read returns, for each cell, its value and formula, the text as the grid displays it
(the cell's number format applied by the core formatter), the number format code (absent for
General) and, for date and time formats, the value in ISO 8601 without a time zone (`2024-04-03`,
`10:30:00` or `2024-04-03T10:30:00`; durations such as `[h]:mm` have none).
`xlsx_find` (`findXlsxCells`) searches the way Excel's Find does (formulas as typed or values as displayed,
match case, whole cell, `*` `?` `~` wildcards, optionally one sheet or range) and returns at most
1,000 matches, each with its sheet, address and the same cell view, plus the total count.
`xlsx_read_table` (`readXlsxTable`) reads a table by name (ignoring case, as Excel does): its sheet, range, header
and totals rows, column names and cells (at most 10,000, as a range read).

The existing PowerPoint MCP keeps its established tool set and names. Its server
exports registration and accepts a root directory for composition. Its model
operations, execution pipeline and collaboration codec live in
`ooxml-core/pptx/automation`; viewer-side compatibility modules delegate to core.
Its schemas are in core (`ooxml-core/pptx/automation/schemas`); its registration still lists
each tool by hand in `viewers/pptx/packages/tools/src/mcp/server.ts` and has not moved onto
`registerAutomationTools` yet.

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
