# XLSX parity status

Last reviewed 2026-10-30: the gap list under "Major work still required" was re-checked against the code; the figures below were not re-measured. This page is the live status tracker for the Excel product, and the [Office parity checklist](parity-checklist.md) carries its open gaps beside Word, PowerPoint and Visio; the
earlier separate Excel review, lookup evidence and implementation-log pages were folded into it
(the per-increment history lives in git). The chart-title and legend increments below were recorded
at `f12ee7ea6` on 8 October 2026, and any test counts quoted further down are snapshots from that revision (8 October 2026), not current figures.
Native references, core and browser regressions, and selected Excel reopen checks bound the
implemented scope.

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

| Parity dimension               | Assessment                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       | Evidence or remaining work                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| ------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Workbook import and saving     | Partial                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Supported content has round-trip regressions and selected Excel acceptance checks. Unsupported features and structural edits need wider preservation coverage.                                                                                                                                                                                                                                                                                                                                                                                                               |
| Editing and calculation        | Partial                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Values, formulas, formatting, clipboard, history and structural editing work within supported scope. 20 Excel functions are not implemented (see item 9 below), and semantic edge cases remain.                                                                                                                                                                                                                                                                                                                                                                              |
| Charts                         | Partial                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Common families plus combination, waterfall, histogram, bubble, stock, treemap and sunburst are drawn. Charts are edited over their `ChartSpace` (insert, type, title, legend, axes, series and the shared Format pane), and selected workflows are tested. Advanced families, complete controls and exact layout remain open.                                                                                                                                                                                                                                               |
| SmartArt                       | Partial                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Insert/node/text/layout editing, shared Styles and Change Colors, interpreted previews and admitted generated 2D caches are implemented with history and Word/Excel native semantic reopen checks. Wider native styles, rich-text authoring, complex 3D and pixel fidelity remain limited; see the shared delivery ledger.                                                                                                                                                                                                                                                   |
| Shapes and form controls       | Shapes, text boxes, connectors and groups are drawn by `layout/shape-view.ts` from the shared preset geometry tables, custom geometry, fills, outlines, `xdr:style` colours and text bodies. Legacy form controls (`formControlPr` plus the VML caption) are read into `FormControlObject` and drawn with a Windows appearance; check boxes, option button groups, drop-downs, list boxes, spin buttons and scroll bars write their linked cell, and the cached `checked`, `sel` and `val` follow the cell on save. Core tests use an Excel-authored fixture (`excel-shapes.xlsx`, `generate-excel-shapes.ps1`). | Text wraps by an estimate (no measurer in the core), with one style per paragraph; no shape effects, 3D, pattern or picture fills, and gradients are linear only. Form controls are not moved or resized in the editor, buttons do not run macros, the drop-down list is a simple popup, and rendering was not pixel-compared with Excel.                                                                                                                                                                                                                                    |
| Custom Views                   | View > Custom Views saves the zoom, grid lines, headings, panes and selection of every sheet under a name, with optional hidden rows and columns, filter settings and print settings (page setup, print options, print area). `customWorkbookViews`, each sheet's `customSheetViews` and the hidden `Z_<GUID>_.wvu.Rows/Cols/FilterData/PrintArea` names are read, written and edited (Show, Add with replace-by-name, Delete, one undo step each); tests use an Excel-authored fixture (`excel-custom-views.xlsx`, `generate-excel-custom-views.ps1`).                                                          | Excel's refusal in workbooks with tables is mirrored. Hidden-row names do not follow later row inserts and deletes, print titles and page breaks of a view are kept as read but not captured or applied, the window position is not applied, and only one workbook (zoom, grid lines, hidden row and column, filter, orientation, header) was checked in Excel 16 through COM: it opened a file with views written here and its native Show restored them. The file needs `windowWidth` and `windowHeight` on each `customWorkbookView`, which the writer adds when missing. |
| Excel UI and visual appearance | Partial                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Shared ribbon/gallery/fill controls and six bindings are exercised. Full ribbon/dialog behavior, text/layout and raster equivalence are unverified or incomplete.                                                                                                                                                                                                                                                                                                                                                                                                            |
| Whole-product parity           | Not achieved                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | Bounded pivots, partial collaboration and Page Layout / Page Break Preview exist; shared pivot caches, label and value filters, data connections, macro execution and full print fidelity are missing, which prevents this claim.                                                                                                                                                                                                                                                                                                                                            |

The latest diagram/chart/XLSX and PowerPoint layout-parser sweep passed
**8,591 ordinary tests**, with **16 accepted native raster gaps** (one named entry, `translucent-three-stop-colour-weighting` in `src/core/chart/gradient-raster.test.ts`).
The updated seven-reference title-layout tests and **16 PowerPoint UI layout
checks** also passed. Focused chart browser checks passed **302 unique cases**
across six bindings, including 48 new legend cases. Core/UI typechecks, builds
and package checks passed. The earlier
9,938-test sweep used a broader shared-text filter; these counts have different
scopes.
These counts describe regression coverage, not a completion percentage or a
complete end-to-end Excel comparison.

## Latest completed increments

2026-10-31: FILTERXML now evaluates XPath 1.0 through the shared XML area, using the existing
bounded XML parser and the `xpath` dependency already used by signature tooling. It returns
vertical spill arrays and follows the native cases for element/attribute values, namespaces,
whitespace, direct text, empty selections and errors. Numeric dates use year/month/day order.
The 35-case Excel 16 corpus and its COM generator are in
`xlsx/formula/__fixtures__/filterxml-native.*`; all six bindings pass the recalculation, undo and
save workflow in `e2e/xlsx/filterxml.spec.ts`. Excel reopened 35-sheet generated outputs in both
the 1900 and 1904 date systems without repair, with no value mismatches after recalculation.
DTD/entity declarations are refused, and XML and XPath inputs are limited to 32,767 characters.
This is local evaluation; web, OLAP, market-data and AI formula services remain unsupported.

2026-10-31: The existing comment dialog now edits imported threaded replies as well as the root.
Edits keep authors, timestamps, message identities, person accounts, imported mentions and unknown
markup. The dialog also resolves and reopens threads; one submission is one undo step. Mention
offsets follow edits around unchanged labels, and editing a label removes that mention. Copies
receive fresh message and mention identities. Independent reply edits from two collaborators
merge, per-person undo retains the other reply, and an open dialog detects conflicting edits.
Thread metadata also reaches peers without the source package. The native fixture
`excel-threaded-comments.xlsx` and its COM generator accompany the round-trip regressions.
All twelve cases across six bindings pass `e2e/xlsx/comment-replies.spec.ts`. Excel 16 reopened
and resaved edited, reopened, duplicated and pasted threads without repair, retaining their
identities, mention ranges and unknown extensions. Core coverage includes
`xlsx/write/threaded-comments.test.ts` and collaboration regressions. Creating new @mentions and
host people-directory integration remain open and are reflected in Feature Status.

2026-10-31: Excel 2010 (`x14`) data validations now use the same model, list drop-down,
restriction checks and Data Validation dialog as ordinary rules. Cross-sheet references are
written in the x14 form, with current formulas and ranges after sheet renames, row inserts,
copying, partial clearing and undo. Source rule metadata and collection settings are kept;
copied rules do not duplicate revision identities. The Excel-authored fixture
`src/core/xlsx/__fixtures__/excel-x14-validation.xlsx` and its COM generator accompany the
round-trip regressions. All six bindings loaded, edited, saved and cleared the native rule in
`e2e/xlsx/x14-validation.spec.ts`, including the grid's list values. Excel 16 reopened three
outputs with the expected edited, cleared and newly authored cross-sheet list formulas.

2026-10-31: Navigation now saves each sheet's active cell and ranges, the active workbook tab,
and each split pane's selection, without adding content undo steps or marking the workbook dirty.
Disjoint selections write `activeCellId` for the range containing the active cell. Frozen panes
share the sheet selection, matching Excel. Excel 16 COM reopened three generated workbooks with
the expected second-sheet D6 selection, first-sheet B2:C3/F8:G9 selection with F8 active, frozen
B2 selection and top-left split B2 selection. Native resaving also exposed omitted default
`pane/@state` and `pane/@activePane` attributes; the reader now recognises those split panes.
Core regressions are `xlsx/ui/selection-persistence.test.ts` and `xlsx/write/split-pane.test.ts`;
`e2e/xlsx/selection-save.spec.ts` checks navigation and pointer selection in all six bindings.
Fill-handle dragging was already implemented: `e2e/xlsx/fill-handle.spec.ts` now confirms numeric
series extension, one-step undo and redo in every binding.

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

This page is the current summary ([xlsx-parity-review.md](xlsx-parity-review.md) is now only a
pointer to the removed increment log). The [feature page](../viewers/xlsx/docs/features.md)
lists the broader product surface. [PROVENANCE](../PROVENANCE.md) records shared
logic reuse and extraction.

## Implemented and verified work

| Area                          | Current implementation                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         | Important limits                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Editing and clipboard         | Finite copies retain trailing blanks; compatible selections tile copied blocks (no size cap; a million-cell selection is tested). A multi-area copy works when the areas share rows or columns, as in Excel 16 (checked through COM): areas are placed next to each other in sheet order with gaps dropped, and other shapes are refused with Excel's message. Directional fills copy content/formatting with translated formulas. Paste Special supports content modes, transpose, Skip Blanks, arithmetic, All Except Borders and column widths. Notes, validation, hyperlinks and supported conditional formats participate in the clipboard path. History, save/reload and structural reference behavior have regressions. | This is not complete Excel clipboard metadata or every cut/paste scenario. Multi-area copy carries cells, merges, notes, links, validations and conditional formats; it shows its marquee on the first area. Cutting several areas is refused with Excel's message, and pasting into a multi-area selection pastes into each area (one undo step), refusing when an area does not hold whole copies.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| Conditional formatting        | Existing rules plus improvements to x14 data-bar preservation, copied rule identities, colors/borders, direction, axes, lengths, thresholds and context. Native comparisons exercise supported cases.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Advanced rule authoring controls and exact data-bar raster appearance remain incomplete.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| Calculation                   | Existing formula engine, dependency recalculation and recorded Excel comparisons; follow-up work includes AMORLINC and DB/DDB compatibility cases.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             | Missing functions, coercion/array/date edge cases and external workbook resolution remain. A function catalogue is not proof of complete compatibility.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| What-If Analysis              | Goal Seek (secant search with bisection fallback over the recalculation engine, workbook iteration settings, status dialog, one undo step), one- and two-variable data tables with `TABLE()` evaluation, `<f t="dataTable">` round trip and protection against partial edits, and Scenario Manager (add, edit, delete, show; `<scenarios>` round trip). Core tests cover solver cases, tables and a generated openpyxl fixture; browser spec covers Goal Seek and a data table.                                                                                                                                                                                                                                                | Forecast Sheet builds a new sheet (timeline, values, FORECAST.ETS and confidence-bound formulas, optional statistics, line or column chart) from a dialog with Excel's options; its numbers carry the FORECAST.ETS differences from Excel, and it has no forecast start offset, no date axis on the chart and no Excel styling of the new sheet. Scenario Summary (standard report) and Merge (same workbook) follow what Excel 16 wrote for `CreateSummary` and `Merge` (`scenario-summary.ts`, `scenario-merge.ts`); there is no Scenario PivotTable report, no merge from other workbooks, no Goal Seek Step or Pause, results were not compared against real Excel output (hand-worked regression tests only). Data tables shift with row and column inserts and deletes; a deleted input cell shows `TABLE(#REF!,...)` and stops recalculating. Scenarios shift the same way: changing cells and result cells follow inserts and deletes, and a deleted changing cell is marked `deleted` and skipped when the scenario is shown. |
| Chart selection and structure | Pictures/charts can be selected, moved, resized and deleted. Charts support insertion, family/grouping changes, title and legend changes, history and save/reload. Every chart is edited over its parsed `ChartSpace` (the shared chart model): an edit is a `chart/edit` operation, an edited part is saved with the shared writer and an unedited part byte for byte, and a type change keeps the chart's formatting instead of regenerating the part.                                                                                                                                                                                                                                                                       | The base SVG families are column, bar, line, area, pie, doughnut, scatter and radar; combination, waterfall, histogram, bubble, stock, treemap and sunburst are drawn as well (see the assessment table). Many native chart options remain unsupported.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| Chart formatting UI           | Shared Chart Design galleries provide curated colour/style presets and Quick Layouts. The shared Format pane writes fills/lines/transparency, joins/ends, effects, shape 3D, text, pattern fills, admitted secondary axes, analysis, markers and axis titles over ChartSpace. ChartEx exposes its supported common-format, family/axis/label controls with safe local-workbook edits.                                                                                                                                                                                                                                                                                                                                          | Picture allocation, multiple-trendline/custom-error source editors, wider combo admission, native style-part authoring and exact chart text/effect/3D fidelity remain incomplete. Unsupported nested content or unreadable values keeps a chart read-only and source-backed.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| Chart rendering and export    | Imported styles, supported gradients/shadows, axis visibility, built-in text defaults, chart-wide font inheritance and mixed title runs/lines/paragraphs. Supported fills and text formatting survive tested chart edits and regeneration.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | Text measurement, wrapping, spacing, automatic/manual layout and full chart raster fidelity remain approximate. Advanced text styling and mixed-format font authoring UI remain open.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| SmartArt                      | Cached or interpreted drawings use the shared painter and workbook theme. The shared picker, node/text/layout commands, text pane and Styles/Change Colors galleries retain history, node overrides, rich text and unknown XML. Admitted interpreted 2D layouts get generated caches. Native Word/Excel checks verify selected text/style/colour ids and nonempty semantic-node shapes.                                                                                                                                                                                                                                                                                                                                        | Unsupported/partial layouts retain an explicit approximate fallback; folding, spatial 3D and unsupported text decline cache generation. Wider native styles, dedicated per-node formatting, rich-text authoring in the text pane and native pixel fidelity remain open.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| Shared architecture           | One XLSX core and one editor component, with six thin framework bindings. Shared XML/OPC, DrawingML/diagram, chart, geometry, color, units and Office UI controls are reused.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  | Shared infrastructure does not by itself implement spreadsheet-specific behavior or establish parity.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |

## Reuse from PowerPoint and other packages

- The chart picker uses `office-ui-gallery`, also used by PowerPoint's ribbon
  gallery. It shares visual tiles, selection, keyboard behavior and theme contracts.
- SmartArt uses the shared `office-ui-smartart` component and format-neutral
  diagram model/color/font resolution. XLSX supplies its workbook theme. Editing
  shares Word's code: the `diagram/edit` node commands and SmartArt Design list,
  the data part writers, the approximate layout, the SmartArt UI strings, and the
  `office-ui-smartart-text-pane` and `office-ui-smartart-picker` elements.
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
8,591 ordinary core tests, 16 accepted native raster gaps, 302 focused browser
cases across vanilla, React, Vue, Angular, Svelte and Solid, and passing
core/UI build, type and package checks. These are scoped implementation-run
results, rather than a full-product equivalence suite. Earlier
results had different scopes and are not carried here.

Core tests cover supported parsing, editing, preservation and layout behavior.
Browser tests cover selected painted geometry and actual edit/history/save
flows. Excel COM and native PNG comparisons independently verify selected
references. Passing those layers establishes their tested scope, not every
option of a chart family or every Excel workflow. The 16 accepted native raster
gaps (translucent three-stop gradient colour weighting, 3.2 to 3.7 levels against a limit of 3) remain unresolved; the earlier count of 17 had no 17th test in the repository.

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

## Pivots, slicers and sparklines

- Pivots: bounded authoring and refresh in `src/core/xlsx/pivot/` (aggregation, layouts, report
  filter, native cache XML) with `model-pivots.ts`, and dialogs in `src/ui/src/xlsx`. Up to 16
  measures, any number of row and column fields, tabular/compact/outline layouts, per-field custom
  subtotals, Values on rows, number and date grouping (seconds to years), and date groups together
  with report filters, hidden items, slicers and timelines round trip. Date group names written by
  Excel in another language are read by position and written back unchanged
  (`PivotDateGroup.localized`), checked against fixtures relabelled as a German Excel names them
  because only the English Excel is installed. Other captions Excel localizes (Row Labels, Grand
  Total, Values, subtotal names) are English-only, so a fully localized workbook stays preserved.
  Not supported and kept read-only with warnings (`pivot/admission.ts`): shared caches, label and
  value filters, filters on date group items, mixed or custom layouts. Native acceptance:
  `scripts/xlsx-pivot-acceptance.ps1` and the `xlsx-pivot-*-oracle.ps1` scripts.
- Slicers and timelines: `src/core/xlsx/slicer/` views, `edit/slicers.ts`,
  `scripts/xlsx-slicer-acceptance.ps1`.
- Sparklines: line, column and win/loss groups are drawn and authored (create, format, clear, with
  an Insert ribbon group; `edit/sparklines.ts`, `scripts/xlsx-sparkline-acceptance.ps1`). Excel
  ignores the extension if its URI uses uppercase `4FD2`, so the writer uses lowercase `4fd2`;
  edits to revision-tagged groups need a distinct revision UID, which the commands supply. Group and
  Ungroup (`edit/sparkline-grouping.ts`) merge the touched groups into the first one or split each
  into one group per sparkline. The style gallery (no ground truth for Excel's 36 colour sets could
  be read through COM), date-axis authoring and individual source-range editing remain; existing
  date-axis metadata is preserved.

## Calculation: lookups and regular expressions

- Lookups: exact and approximate XMATCH/XLOOKUP keep the full reference extent, compare across
  value types (numbers, text, logicals, blanks) and follow native duplicate positions in binary
  search; MATCH/VLOOKUP/HLOOKUP reject scalar tables; omitted-argument defaults for XMATCH,
  ADDRESS, INDIRECT and OFFSET follow one shared call policy; INDIRECT supports R1C1 with whole
  rows/columns and wrapping (`address-r1c1.ts`). Evidence: 216 recorded comparisons
  (`formula/__fixtures__/lookup-native.json`, `lookup-native.test.ts`, recorded by
  `scripts/record-xlsx-lookups.ps1` on Excel 16.0 build 20430). Binary search requires the declared
  sort order; arbitrary unsorted input is not established.
- Regular expressions: `REGEXEXTRACT` implements return modes 0 (first match), 1 (all matches,
  vertical spill) and 2 (capture groups, horizontal spill); an invalid mode or pattern is
  `#VALUE!`, no match is `#N/A`. Patterns are JavaScript regular expressions, so PCRE2-only syntax
  is unsupported, and the installed Excel returned `#NAME?`, so it could not validate them.
- Functions: odd-period bond pricing (`odd-securities.ts`) and Bessel functions (`bessel.ts`)
  exist; `FORECAST.ETS`, `.CONFINT`, `.SEASONALITY` and `.STAT` use the exponential smoothing filter, starting trend, error-statistic window and interval formula identified against real Excel on 34 series (21 series reproduce to the last digit; typical forecasts are within 0.5%, the worst is 14%); what differs is Excel's own search for alpha, beta and gamma. Without a season its objective is the training-part one-step squared error plus the squared forecast error of the hold-out (`etsHoldout` trailing points forecast ahead from the end of the training part without updating; `ets-validate.ts`), from the least-squares trend, invariant to scale, offset and sign, minimised over a fixed lattice of 25 (alpha, beta) pairs (`ets-search-map.ts`): beta 0.001 with 0.002, 0.1, 0.126, 0.25, 0.5, 0.75 and 0.9, then alpha - 0.001 and a few larger betas, among them (1/6, 1/12), (1/3, 1/6), (2/3, 1/6), (5/6, 1/12) and alpha 0.998 with 0.099, 0.249, 0.499, 0.749 and 0.899 (alpha 0.998 with the lowest beta is not a candidate). The pair that reproduces the error statistics STAT 4-7 of a series to the last digit (found by solving for it) is the lowest-cost pair of that lattice on every one of about 2100 non-seasonal probes: 491 search, 484 morphing, 136 spikes and steps, 392 rescaled or reversed, 209 recorded for large-beta pairs (series drawn so that the cost is far lowest there), 231 blending a drifting and a pair series, and 120 of 8 to 10 points (no hold-out, plain one-step error). A series of 4 to 6 points is extended to 7 with the points its starting trend uses before the one-step error is taken (the pair is then Excel's on all 40 recorded series of 4 to 7 points), and 3 points always report 0.1 and 0.099. What FORECAST.ETS.STAT reports is that pair's alpha always, and its beta unless Excel forecasts the last value plus the least-squares slope (alpha 1, beta 0): then STAT 2 says 0.001 while the error statistics still use the pair. The drift forecast itself is exact (alpha 1, beta 0 from the least-squares slope, equal to Excel to 1e-15 on all 440 drifting probes of the 742 non-seasonal probes whose lowest-cost pair has a large beta, `forecast-ets-drift.test.ts`; it is not a fit with a small beta, which is 3e-4 off), and a depth-3 tree on the hold-out shift, the Durbin-Watson of the linear fit and the pair's trend over the slope predicts the choice for 85% cross-validated against 83% for the rule below, which is kept. That choice is the only open point: pairs (0.1, 0.099), (0.998, 0.749), (0.998, 0.899), (1/6, 1/12) and mostly (0.25, 0.249), (0.9, 0.899), (2/3, 1/2), (0.75, 0.749) always drift, (1/3, 1/6) never, and for the others a series drifts when its hold-out mean is within 1.5 sd of its training mean and its slope is under 9.3 standard errors (83% of 643 mixed probes, 72% by the commonest outcome of each pair); R squared (74%), the cost ratio of the drift model to the pair, the pair's final trend against the slope, n, the hold-out size, the margin to the next pair, autocorrelations and Mann-Kendall (60 to 79%) do not separate them, and blending a drifting and a pair series in 20 steps flips the outcome where those statistics move smoothly. The alpha of Excel is therefore found on every probe, the error statistics on all of them, and 95% of the forecasts (the rest are the drift decision, worst 13%). Ruled out as objectives (best lattice agreement on about 1100 probes): one-step error of the whole series (44%) or of the last k points, k-step errors of the hold-out with training-only starting trend, absolute, relative and symmetric errors, rolling-origin multi-step errors, concentrated initial states, and a weighted mix of the one-step and hold-out costs; as search procedures (replayed on the recorded cost surfaces): Nelder-Mead from many starts, steps and tolerances (16% of alphas), hill climbing along the lattice from either end (51%), and an incumbent-with-tolerance walk fitted by random search (51%). What is left unexplained is listed above; with a season whose starting states are Excel's own (4 whole seasons, 5 and more, 4 seasons plus a partial one, 3 seasons plus a partial one and exactly 2 seasons, for seasons of 4 and more) the same hold-out cost is used over the 20 (alpha, beta, gamma) triples Excel is seen to report (its triple on 2307 of the 2319 recorded series; the same triples with 0.002 or 0.9 at the lowest beta and gamma added did worse on the first 700), on the series padded the way the starting trend pads it (2 seasons are extended to 3, and series up to 3 seasons plus ceil(m/2) points to that length) with the hold-out of the padded length: the unpadded series found it on 1790 of 2319, elsewhere the plain one-step error is kept because the starting states are not Excel's (the hold-out cost then chooses worse: season3-11 forecast error 8% against 1%). Tried and not adopted: snapping the continuous minimum to the lattice where the cost has several valleys (165 of 1416 probes: 56% against 39% for the lowest lattice value, but 29 fewer alphas on the 491 search probes), and choosing the forecast `last value + slope` by the cost of that drift model: 56 of 491 series forecast that way, 53 of them with a drift cost within 1.3 of the chosen cost, but 199 series that forecast normally are as close, so no threshold on that ratio predicts it (best 441 of 491 against 435 for never). Also different: Excel's starting states below 2 seasons, and seasons of 2 or 3 (the start trend without a season is exact for 3 to 16 points, 4 to 6 being the line through the data extended to 7 points by copies of early points moved along the slope, `shortSeriesTrend`; for seasonal series Excel's starting state is a fixed linear function of the data, recovered from its forecasts in `ets-init-excel.test.ts` and `ets-init-maps.test.ts`, and for 4 whole seasons or 5 and more it is known exactly, `ets-init-chain.ts`: the position states follow each other along m consecutive times, `s[t] = s[t-1] + P(t) - P(t-1) - B(t - floor(m/2))`, with P the mean of the first C-h seasons of a position (reading later seasons for t >= m) and B its half-split slope with h = floor(C/2), only the first 5 seasons being read, and the start trend is the least-squares slope of the series with those states removed; it reproduces Excel's forecasts at its own parameters to rounding error on about 1200 probe series of seasons 2 to 12, `ets-init-forced.test.ts`; 3 whole seasons plus up to ceil(m/2) + 1 points and exactly 2 seasons (seasons of 4 and more) are also exact, `ets-init-three.ts`: a blend of a forward and a backward chain from the reference position at time m with a slope-sum drift and a corner term, a trend fitted on the series padded to 3m + ceil(m/2) points, and 2 seasons extended to 3 by y[t-2m] + 2m times the dummy-regression slope (to rounding error on about 1100 more probe series of seasons 4 to 12); 4 seasons plus a partial one (`ets-init-four.ts`: the 4-season chain of the first 4m values plus a shifted template on the states of the first r positions) 3 seasons plus more than ceil(m/2) + 1 points up to 4 seasons (`ets-init-threeplus.ts`: the 3-season rule minus a shifted template on the first positions) and 2 seasons plus a partial one (the series extended to 3 seasons by y[t-2m] + 2m times the mean of the position slopes) are exact too for seasons of 4 and more (seasons of 2 and 3 from 6 points), and the lengths below 2 seasons remain approximate, up to 28%; backcasting, fixed-parameter passes, geometric-weight regressions, seasonal-dummy regression and classical decomposition were tried and ruled out), and seasonality detection on noise (autocorrelation first peak at 0.3 agrees on 113 of 126 recorded series, all noiseless ones; Excel also reports periods on a third of short white-noise series by an unknown rule: its period equals the lag of the largest autocorrelation on 6 of 10, but that statistic is as large on the series where it reports none, and the lowest hold-out cost of this engine's seasonal fit and the differenced autocorrelation do not separate them either; on 352 more designed series (`forecast-ets-detect.test.ts`) Excel's answer is identical for the negated, reversed, rescaled and shifted series (768 of 768), so it is not a model fit, and the period is n / k, rounded with halves down, for the Fourier frequency k with the largest periodogram ordinate of the detrended series on 66 of 77 noise series, but whether any period is reported is no threshold on that ordinate's share of the total, and a periodogram detector with a fitted threshold agrees on 384 of 463 series against 377 for the autocorrelation one, which is kept), see `forecast-ets-excel.test.ts`. Native fixtures cover
  AMORLINC, DB and DDB (`scripts/record-xlsx-amorlinc.ps1`, `record-xlsx-depreciation.ps1`).

## Import, print and collaboration

- Import: `load/detect.ts` identifies XLSB so it can be refused with a message; ODS is not
  supported; legacy `.xls` and CSV load through `xlsx/load`. External connections use cached
  values; VBA is carried but never executed; add-ins and Power Query are not implemented.
- Print: paginated preview and print fields support saved page setup and header/footer fields
  with font family/size, RGB and theme colour (`&KTT+SSS`, resolved against the workbook theme),
  bold, italic, underline, double underline, strike, super/subscript, page-number offsets
  (`&P+n`, negative results print as Excel does) and pictures (`&G`, from `legacyDrawingHF`, for
  the default, first and even headers and footers; `excel-header-images.xlsx`). Native PDF
  checks: `scripts/xlsx-print-acceptance.ps1`, `xlsx-print-format-acceptance.ps1`; the
  offsets and theme colours were measured from an Excel PDF export. Tinted theme colours can
  differ from Excel by a couple of levels per channel; picture crop and recolour settings are
  not applied.
- Collaboration: `ooxml-core/xlsx/collab` maps the workbook to Yjs (cell values and formulas,
  formats, row heights, column widths, merges, sheet add/rename/move/delete/state/tab colour,
  defined names, selections; row and column inserts travel as the resulting cell moves), and
  `<xlsx-editor>` is wired (share dialog, presence). Also shared, as keyed entries per item:
  notes, threaded replies (one entry per reply), hyperlinks, conditional formats, validations,
  tables and charts (last writer wins per chart). Not shared yet: images and other drawings,
  filters, pivots, freeze panes and views, page setup, protection and document properties. See [collab-area.md](collab-area.md).
- Clipboard: finite copies keep trailing blanks and tile into compatible selections; directional
  fills translate formulas; Paste Special supports content modes, transpose, Skip Blanks,
  arithmetic, All Except Borders and column widths; notes, validation, hyperlinks and supported
  conditional formats travel (`edit/clipboard.ts`; native recordings `record-xlsx-paste-*.ps1`).
  Clipboard payloads are dense arrays and whole-row/column copies are bounded to used content.

## Major work still required

1. Native chart text measurement, wrapping, spacing, title/legend placement and
   overall chart geometry, followed by wider screenshot/raster comparisons.
2. Complete chart formatting/authoring UI, advanced families, axes, labels,
   effects and style/layout behavior, with native edit/export comparisons.
3. Expand the implemented SmartArt Styles/Change Colors catalogue, add rich-text authoring
   in the text pane, and improve unsupported/partial layouts and native typography/3D fidelity.
   Shared interpreted previews and admitted generated caches already cover uncached diagrams.
4. Broader workbook save fidelity across structural edits, unsupported parts,
   drawings, tables, external links and metadata. No blanket lossless claim.
5. Grid/font/number-format/layout edge cases and native paginated printing.
   Page Layout view and Page Break Preview are implemented (page plan, rulers,
   header/footer editing, break and print-area dragging, saved in
   `sheetView/@view`); still missing there are copy and paste
   inside Page Layout view,
   separate per-view zoom attributes, and native-printer metrics.
6. Pivots beyond the bounded profiles (label and value filters, shared caches, filters on date
   group items, report captions in other languages), sparkline style galleries and
   individual source-range editing, data connections, and the sheet parts the collaboration
   mapping does not share yet (see "Import, print and collaboration").
7. Wider formula and clipboard differential coverage. VBA execution, add-ins,
   Power Query and unsupported file formats also prevent whole-product parity.
8. Filter and sort. The sheet AutoFilter header cells have a drop-down (`grid/filter-buttons.ts`,
   `grid/filter-menu.ts`): Sort, Clear Filter From, Filter by Color (cell fill and font),
   Text / Number / Date Filters with the Custom AutoFilter and Top 10 dialogs, Above and Below
   Average, the date periods, a search box and the value checklist with (Select All) and
   (Blanks). Value lists with date groups, custom (with wildcards), top 10, dynamic and colour
   criteria are modelled (`AutoFilterColumn`), read, written and evaluated to hide rows
   (`filter-criteria.ts`); icon criteria are evaluated when the sheet's conditional formats
   supply the icon. Not done: drop-downs on table headers (a table keeps value lists only),
   the date checklist as a year / month / day tree (dates list as their displayed text),
   and colours that come only from conditional formats in the menu's colour list. Data > Sort
   has Sort On cell values, cell colour, font colour and conditional-format icon (On Top / On
   Bottom, or On Left / On Right), Order A to Z / Z to A / Custom List (the four built-in
   weekday and month lists or a list typed in the dialog), and Options for case sensitivity
   (lower case first) and left-to-right sorting (`edit/sort.ts`, `edit/sort-keys.ts`,
   `edit/sort-choices.ts`, `sort-lists.ts`). The levels are saved as `sortState` (sheet or
   auto-filter: `sortBy`, `dxfId`, `iconSet`, `iconId`, `customList`, `caseSensitive`,
   `columnSort`) and round-trip, a table's sort in its own part (`table.sortState`, patched into
   `xl/tables/tableN.xml`). Data > Reapply re-runs the saved sort (table, filter or sheet) and
   the filter criteria (`edit/sort-reapply.ts`). Typed custom lists persist in browser
   localStorage (`dialogs/custom-lists.ts`). Not done: an Options > Edit Custom Lists manager
   (rename, delete, import from cells) and a roaming store; hidden columns are skipped but not
   checked against Excel when sorting left to right.
   The criteria and sort modes were checked against the ECMA-376 semantics and unit tests, not against real
   Excel row results.
9. Functions. 502 functions are registered (`viewers/xlsx/docs/functions.md`); these 20
   return `#NAME?`: INFO, ASC, DBCS, JIS, PHONETIC and EUROCONVERT (no external service
   needed), and CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY, CUBERANKEDMEMBER, CUBESET,
   CUBESETCOUNT, CUBEVALUE, WEBSERVICE, STOCKHISTORY, RTD, IMAGE, COPILOT,
   TRANSLATE and DETECTLANGUAGE (need an OLAP, web or AI service). These integrations are named in Feature Status and the generated function list. Cached unsupported results can survive initial loading, but cannot be refreshed. FILTERXML is implemented locally.
10. Data tools and auditing (View > Split is implemented: `layout/split.ts`, `SheetView.split`):
    Flash Fill, Subtotal, Consolidate and Advanced filter have no commands. Formula auditing (Trace
    Precedents and Dependents, Remove Arrows, Evaluate Formula with Step In and Step Out, Error
    Checking and the Watch Window) exists. Error Checking covers seven rules (not the two-digit-year
    and list data validation ones) and has no green-triangle indicators; OFFSET and INDIRECT targets
    are not traced. Text to Columns splits at
    delimiters or typed fixed-width break positions (no per-column data format step), and
    Page Layout > Themes > Colors only reports that changing the
    theme is not supported.
11. File format. `.xlsb` is detected and refused (`load/detect.ts`) and `.ods` is not read or
    written. A save keeps the loaded `calcPr` and writes `fullCalcOnLoad="1"` only when a
    cached formula result may be stale (no cached value, or an edit not recalculated), so Excel
    recalculates those workbooks on open. x14 data validation is modelled and editable, including
    lists over another sheet. Rich values and in-cell pictures (`xl/richData`, IMAGE() results, linked
    data types) have no model, phonetic guides (`rPh`) are read and written with their string
    but not displayed, edited or copied,
    connections and query tables are only carried verbatim, and named sheet views and
    ActiveX controls are not modelled.
12. Ribbon. There are no Draw, Picture Format, Shape Format, Chart Format (a pane instead),
    PivotTable Analyze or PivotTable Design tabs; Insert has no Shapes, Icons, Screenshot or
    Text group (Text Box, Header & Footer, WordArt, Object), Page Layout has no Scale to
    Fit, Print Titles, Background or Arrange group, and Review has no Spelling,
    Accessibility checker or Notes commands.

## Completion criteria

Parity remains an active objective. Each feature needs evidence for import,
rendering, editing, undo/redo, save/reopen in Excel and the actual UI workflow,
using representative native fixtures and explicit unsupported cases. Visual
parity needs measured geometry and raster comparisons, not just matching XML
or SVG attributes. Full parity cannot be declared while any required feature
is missing, unverified or covered only by a narrower comparison.
