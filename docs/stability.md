# Stability policy

This page says what the published packages promise, what counts as a breaking change, how things
are deprecated, and what the Alpha, Beta and Stable labels on a product mean. It describes the
rules the 1.0 releases of `ooxml-core`, `docx-core`, `@christophervr/xlsx-core`,
`docx-viewer-mcp`, `xlsx-viewer-mcp` and `ooxml-mcp` are held to. Releases themselves are
described in [releasing.md](releasing.md), tool definitions in [mcp.md](mcp.md).

Stable is a statement about API and data safety. It is **not** a claim of Office parity or of
lossless export of everything Word, Excel or PowerPoint can write. Unsupported features are
reported, not silently dropped (see "Preservation").

## What is public API

Public API is what a package exposes through the entry points in the `exports` field of its
`package.json`, and nothing else. Deep imports into `dist/` or `src/`, files not listed in
`exports`, and anything marked `@internal` are not public, even if they resolve today.

| Package                                                                                  | Public entry points                                                                                                 |
| ---------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| `ooxml-core`                                                                             | The subpaths in `src/core/package.json` `exports` (`ooxml-core/xml`, `/opc`, `/docx`, `/xlsx`, `/automation`, ...). |
| `ooxml-ui`                                                                               | The subpaths in `src/ui/package.json` `exports`; the `office-ui-*` element tags and their properties and events.    |
| `docx-core`                                                                              | `.` and `./embedded`, both re-exports of `ooxml-core/docx` and `ooxml-core/docx/embedded`.                          |
| `@christophervr/xlsx-core`                                                               | `.` and `./load`, re-exports of `ooxml-core/xlsx` and `ooxml-core/xlsx/load`.                                       |
| `docx-viewer-mcp`, `xlsx-viewer-mcp`, `visio-viewer-mcp`, `pptx-viewer-mcp`, `ooxml-mcp` | The `registerTools` export, the CLI name and the MCP tools they register.                                           |

The MCP tools are public API: their names, kinds, input schemas and the shape of their results.

### The supported surface of `docx-core` and `@christophervr/xlsx-core`

Both packages re-export a whole core area, so the root entry has several hundred runtime names.
The exact list is committed as a snapshot and a test fails on any difference:

- `src/core/docx/__snapshots__/public-surface.json` and `public-surface-embedded.json`
- `src/core/xlsx/__snapshots__/public-surface.json` and `public-surface-load.json`

(The tests are `public-surface.test.ts` next to them. Types are erased and are not listed, but
follow the same rules as the names they accompany.) Changing the list means updating the
snapshot in the same commit, which makes the change visible in review and a reason to pick the
right semver bump.

The list is the **reachable** surface. The **supported** surface, covered by the semver rules
below, is the following; every other exported name is **internal**: it stays exported because
in-repo packages import it, it may change in any minor release and it is not in the changelog.

Word (`docx-core`):

- Loading and saving: `loadDocx`, `saveDocx`, the document model types and the embedded-document
  entry (`docx-core/embedded`).
- The model and its small value helpers: the exports of `model`, `numbering-model`,
  `page-setup-model`, `shape-model` and the units, plus `parseTheme`, the style catalogue parsers
  (`parseRunStyleCatalog`, `parseTableStyleCatalog`, `parseNumberingCatalog`) and the
  style and numbering resolvers.
- Editing commands: the `*-edit` modules (sections, content controls, list start overrides,
  charts) and `fillFormFields`, `refreshStyleRefFields`.
- Review, document statistics, tables of contents, fields (`fieldName`, date fields) and notes
  numbering.

Not supported (internal): the generated schema simple types (`generated/wml-simple-types`),
diagram and chart painting helpers (`diagram*`, `docxChart*`, `shapeSvg` and related paint
functions), table visual resolution used only by the editor, parse-diagnostic plumbing
(`withParseWarnings`, `reportParseWarning`) and layout helpers the editor paints with.

Excel (`@christophervr/xlsx-core`):

- Loading and saving: `loadXlsx`, `saveXlsx`, and from `./load` `loadWorkbook`, `saveWorkbook`,
  `detectWorkbookFormat` with their option and error types (`DataIntegrityError`).
- The workbook model and cell access: `model`, `address`, `cells`, `workbook`, `styles`,
  `createWorkbook`, `createWorksheet`, `getCell`, `putCell`, `parseAddress`, `formatAddress`.
- The formula engine, number formats (`numfmt`), the editing commands (`edit/`), pivots,
  slicers, and `compareWorkbooks` with its report.

Not supported (internal): everything in `layout/` (the DOM-free grid metrics, colours, cell views
and chart SVG the Excel UI paints with; those are consumed through `ooxml-ui`), and the read and
write helpers under `read/` and `write/` beyond the load and save entry points.

The legacy `.xls` reader is never part of the main entry; it needs `./load`.

If you rely on a name in the internal group, open an issue; promoting it is a documented minor
change.

## Semver

Each package is versioned independently with Conventional Commits (see
[releasing.md](releasing.md)). From 1.0, for the public API above:

**Breaking (major release)**

- Removing or renaming a supported export, entry point, element tag, property or event.
- Changing a type incompatibly: a narrower parameter, a wider or different result, a new required
  property, a removed union member.
- Changing an MCP tool name, kind, a required input, or an input's type; removing an input;
  narrowing an input's bounds; changing the shape of a result a client can read.
- Changing serialized output so that content the previous version preserved is dropped or
  corrupted, or so that a file the previous version produced no longer opens in the product.
- Raising the minimum supported Node.js or runtime version.

**Not breaking (minor or patch)**

- Adding exports, optional parameters, optional tool inputs, new tools, new result fields.
- Fixing a bug, including output that was wrong before, and changing error message text.
- Reading more of a format, writing more of a format, adding warnings.
- Any change to an internal name.

The `ooxml-core` root version and the standalone packages move separately: a package that only
re-exports (`docx-core`, `@christophervr/xlsx-core`) takes a major release when the area it
re-exports takes a breaking change to a supported name.

## Deprecation

A supported name is deprecated for at least one minor release before removal:

1. Mark it with a JSDoc `@deprecated` tag that names the replacement.
2. Add a "Deprecated" note to the changelog in the same release.
3. Keep it working, with the same behaviour, until a major release removes it.

An MCP tool is deprecated the same way: its description starts with `Deprecated: use <tool>.`,
and it stays registered until the next major release. Security fixes and data-loss bugs are the
only exceptions and are described in the changelog.

## MCP schema versioning

Every tool carries an integer `version`, starting at 1 (`AutomationTool.version`). It is raised
by one whenever a change to the tool is breaking by the rules above (renamed tool, new required
input, changed input type, narrowed bounds, changed result shape). The MCP servers send it to
clients as `_meta.schemaVersion` on the tool definition, so a client can detect that a tool it
cached has changed. Non-breaking additions (an optional input, a new result field) do not change
it.

The tool names, kinds, accepted extensions, versions and JSON input schemas of the Word, Excel,
Visio and PowerPoint tools are committed as snapshots in
`src/core/automation/__snapshots__/tool-schemas.<format>.json` (PowerPoint: `src/core/pptx/automation/__tests__/tool-schemas.pptx.json`) and checked by
`src/core/automation/tool-schemas.test.ts` and `src/core/pptx/automation/__tests__/tool-schemas.test.ts`. Any schema change is a visible diff; refresh with
`vitest -u` in `src/core` only after deciding whether the version must be raised.

## Alpha, Beta and Stable

Each product (Word, Excel, PowerPoint, Visio, Teams) carries a label in its README.

- **Alpha**: the API can change in any release and saved output may drop content the product does
  not model. For trying it out.
- **Beta**: the API changes only with a changelog note and, where practical, a deprecation. The
  product has a written list of supported features. Round-trip tests exist, but the preservation
  guarantee below is not yet promised in full.
- **Stable** requires all of these:
  1. A **frozen public API under semver**, as defined above, for the packages the product ships.
  2. A **written list of supported features** and of known unsupported ones, kept in the
     product's docs and updated with the code.
  3. A **preservation guarantee**: content the product does not model (unknown elements,
     extension lists, parts, relationships, customXml and the like) survives load and save of a
     file the product edited, or the product reports that it cannot keep it. Silent loss is a
     bug that is fixed in a patch release.
  4. **Native Office save and reopen acceptance** (the real application opens what we saved
     without repair) run on a schedule, with the last passing run recorded in the docs. For Excel
     that is `src/core/scripts/xlsx-excel-acceptance.ps1`.

Until the self-hosted Office runner is registered, a passing local run of the same
workflow (`node scripts/office-acceptance.mjs <product>`), recorded with its date, Office build and
commit in [office-acceptance.md](office-acceptance.md), satisfies point 4.

Stable is not a claim that the product matches Microsoft Office: rendering, layout and feature
coverage are described honestly in each product's feature list, and anything not on it is
unsupported.

## Why the Excel packages keep the `@christophervr` scope

The unscoped npm names `xlsx-core` and `xlsx-viewer` belong to other people, so the Excel
packages stay `@christophervr/xlsx-core` and the other `@christophervr/xlsx-*` names at 1.0 and
afterwards. Renaming them would break every consumer for no gain, and the scope is part of the
public API like any other name. The Word packages are unscoped (`docx-core`) and stay so.

## Current labels

| Product    | Label  | Notes                                                                                                                                                    |
| ---------- | ------ | -------------------------------------------------------------------------------------------------------------------------------------------------------- |
| PowerPoint | (none) | Carries no label on the launcher; not yet assessed against this bar.                                                                                     |
| Word       | Stable | Since 2026-10-29: API snapshot, [preservation guarantee](../viewers/docx/docs/preservation.md), Word corpus and native acceptance.                       |
| Excel      | Stable | Since 2026-10-29: API snapshot, [preservation guarantee](../viewers/xlsx/docs/preservation.md), function list and native acceptance.                     |
| Visio      | Beta   | Plain-text editing and the shape clipboard are still experimental ([API](../viewers/visio/docs/api.md)); there is no written preservation guarantee yet. |
| OpenTeams  | Beta   | 1.0 packages on sync protocol 1 and schema 1; the Docker recipe has not been built and no production deployment has been reported.                       |
