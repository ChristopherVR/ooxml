# XLSX parity status

Updated 8 October 2026. The latest XLSX increment renders imported manual
legend rectangles and overlay reservation, following title placement and
layout preservation at `f12ee7ea6`. Native references, core and browser
regressions, and selected Excel reopen checks bound the implemented scope.

**Full 1:1 Microsoft Excel parity has not been achieved.** The product is a
working spreadsheet editor with substantial supported behavior and an expanding
native comparison corpus. Compatibility is verified for specific cases, not
for every Excel feature, workbook or UI interaction. We have not established
lossless saving or whole-workbook pixel equivalence.

## Current assessment

We have a working spreadsheet editor, with partial compatibility and native
evidence for specific workflows. We do not yet have Excel-equivalent behavior
or UI. A supported feature means its listed scope works; it does not mean all
Excel options, interactions or visual details in that area are implemented.

| Parity dimension               | Assessment   | Evidence or remaining work                                                                                                                                        |
| ------------------------------ | ------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Workbook import and saving     | Partial      | Supported content has round-trip regressions and selected Excel acceptance checks. Unsupported features and structural edits need wider preservation coverage.    |
| Editing and calculation        | Partial      | Values, formulas, formatting, clipboard, history and structural editing work within supported scope. Missing functions and semantic edge cases remain.            |
| Charts                         | Partial      | Eight SVG families and selected formatting/edit/export workflows are tested. Advanced families, complete controls and exact layout remain open.                   |
| SmartArt                       | Display-only | Shared cached drawings render; insertion, editing and reflow are absent.                                                                                          |
| Excel UI and visual appearance | Partial      | Shared ribbon/gallery/fill controls and six bindings are exercised. Full ribbon/dialog behavior, text/layout and raster equivalence are unverified or incomplete. |
| Whole-product parity           | Not achieved | Missing pivots, advanced data features, native page views, collaboration and macro execution prevent this claim.                                                  |

The latest diagram/chart/XLSX and PowerPoint layout-parser sweep passed
**8,591 ordinary tests**, with **17 existing expected raster failures**.
The updated seven-reference title-layout tests and **16 PowerPoint UI layout
checks** also passed. Focused chart browser checks passed **302 unique cases**
across six bindings, including 48 new legend cases. Core/UI typechecks, builds
and package checks passed. The earlier
9,938-test sweep used a broader shared-text filter; these counts have different
scopes.
These counts describe regression coverage, not a completion percentage or a
complete end-to-end Excel comparison.

## Latest completed increments

| Revision    | Implemented behavior                                                                         | Scope of evidence                                                                                                 |
| ----------- | -------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| `4f3f87ea6` | Import and preserve mixed-format chart-title runs and paragraphs through chart regeneration. | Native character formatting and actual UI type-change/save/reopen comparisons.                                    |
| `e74cf9e85` | Use the grid's cached browser font measurer for chart-title and legend widths.               | Rich-run advances, centering and edit regressions. Headless estimates remain a fallback.                          |
| `8499a584e` | Use measured font ascent/descent for title heights and mixed-line baselines.                 | Three native title-box references and a bounded PNG comparison.                                                   |
| `4f0d6198d` | Read, render and preserve title line/before/after spacing in points or percentages.          | Twelve native references, 72 additional browser cases and Excel reopen/resave checks. No spacing-authoring UI.    |
| `14a33ed2e` | Wrap chart titles with shared styled-run text flow and retain fractional chart dimensions.   | Ten native references, six-binding wrapping/resize/history checks and an MCP-reviewed export reopened in Excel.   |
| `f12ee7ea6` | Render imported manual title positions and overlays; preserve title/plot/legend layout XML.  | Seven native references, 42 additional browser cases, shared PowerPoint checks and two Excel-reopened UI exports. |

## Chart-title layout: implemented scope and remaining limits

Imported manual title positions and overlay flags now render and survive
chart type regeneration. Title/plot/legend layout XML and its extensions are
preserved through that path. Seven native references and two MCP-reviewed UI
exports reopened/resaved in Excel verify the selected cases. Title position
parsing and rectangle geometry reuse the PowerPoint implementation through
shared core chart helpers. Manual title dimensions, plot placement rendering,
position authoring controls and full chart geometry remain incomplete.

Automatic chart-title wrapping now uses one shared text-flow helper in XLSX
and the PowerPoint static SVG converter. Words can cross formatting runs;
oversized words split at grapheme boundaries in XLSX. Soft wraps do not alter
the source title or its rich-body XML. Paragraph before/after spacing applies
at the wrapped paragraph edges. The grid retains fractional drawing dimensions
to avoid moving a break when width is rounded to a whole pixel.

Ten independently reopened Excel references cover chart widths from 160 to
480 points, mixed font sizes, an oversized word and paragraph spacing. With
captured Chromium Arial metrics, title widths differ from these native boxes
by less than 2 CSS pixels and heights by less than 4. These bounds are specific
to this corpus. Playwright MCP verified the native 200-point word reference
before/after a type change. Excel reopens, resaves and reopens that export with
the full title and 152.07 by 65.1-point title box unchanged.

Manual title dimensions, autofit, other scripts/fonts and exact title/plot placement
remain unverified or incomplete. The native and browser screenshots still
show different plot geometry and gridline appearance. This is automatic
wrapping support for the tested scope, not full chart or text parity.

This page is the current summary. The [implementation review](xlsx-parity-review.md)
records each increment and its evidence; older sections describe the state at
their respective revisions. The [feature page](../viewers/xlsx/docs/features.md)
lists the broader product surface. [PROVENANCE](../PROVENANCE.md) records shared
logic reuse and extraction.

## Implemented and verified work

| Area                          | Current implementation                                                                                                                                                                                                                                                                                                                                                                                                                          | Important limits                                                                                                                                                                                   |
| ----------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Editing and clipboard         | Finite copies retain trailing blanks; compatible selections tile copied blocks. Directional fills copy content/formatting with translated formulas. Paste Special supports content modes, transpose, Skip Blanks, arithmetic, All Except Borders and column widths. Notes, validation, hyperlinks and supported conditional formats participate in the clipboard path. History, save/reload and structural reference behavior have regressions. | This is not complete Excel clipboard metadata or every cut/paste scenario. Multi-area paste and large-selection semantics have restrictions.                                                       |
| Conditional formatting        | Existing rules plus improvements to x14 data-bar preservation, copied rule identities, colors/borders, direction, axes, lengths, thresholds and context. Native comparisons exercise supported cases.                                                                                                                                                                                                                                           | Advanced rule authoring controls and exact data-bar raster appearance remain incomplete.                                                                                                           |
| Calculation                   | Existing formula engine, dependency recalculation and recorded Excel comparisons; follow-up work includes AMORLINC and DB/DDB compatibility cases.                                                                                                                                                                                                                                                                                              | Missing functions, coercion/array/date edge cases and external workbook resolution remain. A function catalogue is not proof of complete compatibility.                                            |
| Chart selection and structure | Pictures/charts can be selected, moved, resized and deleted. Charts support insertion, family/grouping changes, title and legend changes, history and save/reload.                                                                                                                                                                                                                                                                              | The supported SVG families are column, bar, line, area, pie, doughnut, scatter and radar. Other families and many native chart options remain unsupported.                                         |
| Chart formatting UI           | Shared Change Colors gallery, individual series fills, bar gap/overlap controls and docked fill panes. No fill, solid, transparency and supported linear/rectangular/circular/shape gradients have authoring paths, stop controls, previews and native presets. Chart area, plot area, title and legend reuse the same fill controls.                                                                                                           | Borders/effects, full axes/data-label editing, style/layout galleries and exact native ribbon/pane behavior remain incomplete. Gradient support varies by chart element/type.                      |
| Chart rendering and export    | Imported styles, supported gradients/shadows, axis visibility, built-in text defaults, chart-wide font inheritance and mixed title runs/lines/paragraphs. Supported fills and text formatting survive tested chart edits and regeneration.                                                                                                                                                                                                      | Text measurement, wrapping, spacing, automatic/manual layout and full chart raster fidelity remain approximate. Advanced text styling and mixed-format font authoring UI remain open.              |
| SmartArt                      | Imported cached drawings render through `office-ui-smartart`; shared theme colors, transforms and font choices improve native appearance. A native Basic Block List reference verifies three labels and measured text properties.                                                                                                                                                                                                               | Display-only. No insertion, text editing, layout reflow or complete layout/color/style galleries. Without a cached drawing, the viewer lists text in a placeholder. No whole-diagram pixel parity. |
| Shared architecture           | One XLSX core and one editor component, with six thin framework bindings. Shared XML/OPC, DrawingML/diagram, chart, geometry, color, units and Office UI controls are reused.                                                                                                                                                                                                                                                                   | Shared infrastructure does not by itself implement spreadsheet-specific behavior or establish parity.                                                                                              |

## Reuse from PowerPoint and other packages

- The chart picker uses `office-ui-gallery`, also used by PowerPoint's ribbon
  gallery. It shares visual tiles, selection, keyboard behavior and theme contracts.
- SmartArt uses the shared `office-ui-smartart` component and format-neutral
  diagram model/color/font resolution. XLSX supplies its workbook theme.
- Chart models, colors, gradients, fills, XML preservation and appearance logic
  live in shared core areas. Mixed titles reuse the DrawingML text reader and
  existing chart/theme resolver, rather than a viewer-specific text parser.
- Series and background formatting panes share fill bindings and controls;
  edits use the existing XLSX transaction, history and save pipelines.
- Chart-title wrapping and PowerPoint static SVG text share `wrapStyledRuns`
  in the core text area. Font resolution and measurement stay with their
  existing product/host implementations.
- Manual chart-layout values and rectangle geometry share one core helper,
  extracted from the PowerPoint parser and renderer. XLSX uses it for imported
  title positions and preserves source layout extensions through regeneration.
- Modern OOXML logic stays in `ooxml-core`; legacy XLS/CFB codecs stay in `ole2`.
  Framework adapters do not contain duplicate spreadsheet engines.

## What the recorded validation proves

The latest completed increment has the regression results listed above:
8,591 ordinary core tests, 17 expected raster failures, 302 focused browser
cases across vanilla, React, Vue, Angular, Svelte and Solid, and passing
core/UI build, type and package checks. These are scoped implementation-run
results, rather than a full-product equivalence suite. Historical
results and their different scopes are retained in the implementation review.

Core tests cover supported parsing, editing, preservation and layout behavior.
Browser tests cover selected painted geometry and actual edit/history/save
flows. Excel COM and native PNG comparisons independently verify selected
references. Passing those layers establishes their tested scope, not every
option of a chart family or every Excel workflow. The 17 expected raster
failures remain unresolved visual differences.

Native evidence uses owned hidden Excel instances, principally Excel 16.0
build 20430. The current chart text corpus includes 96 built-in style
references, 24 chart-wide inheritance/direct-format references and three
mixed-title references. These are independently reopened measurements, with
chart XML retained in test-only fixtures.

Playwright MCP reviewed actual imports, controls and export flows. Recent UI
exports were reopened and resaved through Excel COM, then reopened again. The
mixed-title export retained the font size, name, bold, italic, underline and
color of all 23 measured title characters after a chart type change.

These results establish the tested behavior on that Excel build. They do not
prove every chart type, every workbook, every browser, Excel for Mac/web,
all Microsoft 365 versions or complete UI equivalence. Test counts are not
a parity percentage, and no defensible overall completion percentage exists yet.

## Imported manual legends

Eight native Excel references have been collected for automatic, moved,
overlay, wide and tall legends. They record source XML and native geometry.
The XLSX painter now uses the shared chart rectangle resolver for imported
legend positions and dimensions. The tested wide legends use horizontal
entry flow; tall legends use vertically spaced entries with aligned keys.
Moving a non-overlay legend retains the automatic side reservation, while an
overlay releases that band. Manual/overlay legends paint after the plot.

Core checks retain layout and overlay source XML through type regeneration.
MCP reviewed actual tall/wide workbook imports and a type-change export.
Excel reopens, resaves and reopens the tall Line export with Left=40pt,
Top=100pt, Width=90pt, Height=150pt, Arial 12pt and IncludeInLayout=false.
Native PNG comparisons expose remaining label-spacing and paint differences.
Automatic geometry, multi-row/column packing, entry-specific formatting,
line/marker swatches and position-authoring controls remain incomplete.

## Major work still required

1. Native chart text measurement, wrapping, spacing, title/legend placement and
   overall chart geometry, followed by wider screenshot/raster comparisons.
2. Complete chart formatting/authoring UI, advanced families, axes, labels,
   effects and style/layout behavior, with native edit/export comparisons.
3. SmartArt insertion, text editing, reflow and shared layout/color/style UI.
4. Broader workbook save fidelity across structural edits, unsupported parts,
   drawings, tables, external links and metadata. No blanket lossless claim.
5. Grid/font/number-format/layout edge cases and native paginated printing,
   Page Layout and page-break preview.
6. Pivot refresh/rendering, slicers, timelines, sparklines, data connections
   and spreadsheet collaboration integration.
7. Wider formula and clipboard differential coverage. VBA execution, add-ins,
   Power Query and unsupported file formats also prevent whole-product parity.

## Completion criteria

Parity remains an active objective. Each feature needs evidence for import,
rendering, editing, undo/redo, save/reopen in Excel and the actual UI workflow,
using representative native fixtures and explicit unsupported cases. Visual
parity needs measured geometry and raster comparisons, not just matching XML
or SVG attributes. Full parity cannot be declared while any required feature
is missing, unverified or covered only by a narrower comparison.
