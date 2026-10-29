# XLSX parity status

Last reviewed 2026-10-22 (text only; the figures below were not re-measured then). This page is the single live status tracker for the Excel product; the
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
| Editing and calculation        | Partial                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Values, formulas, formatting, clipboard, history and structural editing work within supported scope. Missing functions and semantic edge cases remain.                                                                                                                                                                                                                                                                                                                                                                                                                       |
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

| Area                          | Current implementation                                                                                                                                                                                                                                                                                                                                                                                                                                                          | Important limits                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Editing and clipboard         | Finite copies retain trailing blanks; compatible selections tile copied blocks (no size cap; a million-cell selection is tested). A multi-area copy works when the areas share rows or columns, as in Excel 16 (checked through COM): areas are placed next to each other in sheet order with gaps dropped, and other shapes are refused with Excel's message. Directional fills copy content/formatting with translated formulas. Paste Special supports content modes, transpose, Skip Blanks, arithmetic, All Except Borders and column widths. Notes, validation, hyperlinks and supported conditional formats participate in the clipboard path. History, save/reload and structural reference behavior have regressions.                                 | This is not complete Excel clipboard metadata or every cut/paste scenario. Multi-area copy carries cells, merges, notes and links but not validations or conditional formats of the areas, cannot be cut, and shows its marquee on the first area; pasting into a multi-area selection is refused.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| Conditional formatting        | Existing rules plus improvements to x14 data-bar preservation, copied rule identities, colors/borders, direction, axes, lengths, thresholds and context. Native comparisons exercise supported cases.                                                                                                                                                                                                                                                                           | Advanced rule authoring controls and exact data-bar raster appearance remain incomplete.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| Calculation                   | Existing formula engine, dependency recalculation and recorded Excel comparisons; follow-up work includes AMORLINC and DB/DDB compatibility cases.                                                                                                                                                                                                                                                                                                                              | Missing functions, coercion/array/date edge cases and external workbook resolution remain. A function catalogue is not proof of complete compatibility.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| What-If Analysis              | Goal Seek (secant search with bisection fallback over the recalculation engine, workbook iteration settings, status dialog, one undo step), one- and two-variable data tables with `TABLE()` evaluation, `<f t="dataTable">` round trip and protection against partial edits, and Scenario Manager (add, edit, delete, show; `<scenarios>` round trip). Core tests cover solver cases, tables and a generated openpyxl fixture; browser spec covers Goal Seek and a data table. | Forecast Sheet builds a new sheet (timeline, values, FORECAST.ETS and confidence-bound formulas, optional statistics, line or column chart) from a dialog with Excel's options; its numbers carry the FORECAST.ETS differences from Excel, and it has no forecast start offset, no date axis on the chart and no Excel styling of the new sheet. Scenario Summary (standard report) and Merge (same workbook) follow what Excel 16 wrote for `CreateSummary` and `Merge` (`scenario-summary.ts`, `scenario-merge.ts`); there is no Scenario PivotTable report, no merge from other workbooks, no Goal Seek Step or Pause, results were not compared against real Excel output (hand-worked regression tests only). Data tables shift with row and column inserts and deletes; a deleted input cell shows `TABLE(#REF!,...)` and stops recalculating. Scenarios shift the same way: changing cells and result cells follow inserts and deletes, and a deleted changing cell is marked `deleted` and skipped when the scenario is shown. |
| Chart selection and structure | Pictures/charts can be selected, moved, resized and deleted. Charts support insertion, family/grouping changes, title and legend changes, history and save/reload. Every chart is edited over its parsed `ChartSpace` (the shared chart model): an edit is a `chart/edit` operation, an edited part is saved with the shared writer and an unedited part byte for byte, and a type change keeps the chart's formatting instead of regenerating the part.                        | The base SVG families are column, bar, line, area, pie, doughnut, scatter and radar; combination, waterfall, histogram, bubble, stock, treemap and sunburst are drawn as well (see the assessment table). Many native chart options remain unsupported.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| Chart formatting UI           | Shared Chart Design galleries provide curated colour/style presets and Quick Layouts. The shared Format pane writes fills/lines/transparency, joins/ends, effects, shape 3D, text, pattern fills, admitted secondary axes, analysis, markers and axis titles over ChartSpace. ChartEx exposes its supported common-format, family/axis/label controls with safe local-workbook edits.                                                                                           | Picture allocation, multiple-trendline/custom-error source editors, wider combo admission, native style-part authoring and exact chart text/effect/3D fidelity remain incomplete. Unsupported nested content or unreadable values keeps a chart read-only and source-backed.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| Chart rendering and export    | Imported styles, supported gradients/shadows, axis visibility, built-in text defaults, chart-wide font inheritance and mixed title runs/lines/paragraphs. Supported fills and text formatting survive tested chart edits and regeneration.                                                                                                                                                                                                                                      | Text measurement, wrapping, spacing, automatic/manual layout and full chart raster fidelity remain approximate. Advanced text styling and mixed-format font authoring UI remain open.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| SmartArt                      | Cached or interpreted drawings use the shared painter and workbook theme. The shared picker, node/text/layout commands, text pane and Styles/Change Colors galleries retain history, node overrides, rich text and unknown XML. Admitted interpreted 2D layouts get generated caches. Native Word/Excel checks verify selected text/style/colour ids and nonempty semantic-node shapes.                                                                                         | Unsupported/partial layouts retain an explicit approximate fallback; folding, spatial 3D and unsupported text decline cache generation. Wider native styles, dedicated per-node formatting, rich-text authoring in the text pane and native pixel fidelity remain open.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| Shared architecture           | One XLSX core and one editor component, with six thin framework bindings. Shared XML/OPC, DrawingML/diagram, chart, geometry, color, units and Office UI controls are reused.                                                                                                                                                                                                                                                                                                   | Shared infrastructure does not by itself implement spreadsheet-specific behavior or establish parity.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |

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
  exist; `FORECAST.ETS`, `.CONFINT`, `.SEASONALITY` and `.STAT` use a documented additive exponential smoothing whose optimiser is not Excel's; 17 series were compared with real Excel (forecasts typically within 0.3% to 2%, up to 37% on short or exponential series; seasonality agrees on 14 of 15; alpha, beta, gamma and the error statistics differ), see `forecast-ets-excel.test.ts`. Native fixtures cover
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

## Completion criteria

Parity remains an active objective. Each feature needs evidence for import,
rendering, editing, undo/redo, save/reopen in Excel and the actual UI workflow,
using representative native fixtures and explicit unsupported cases. Visual
parity needs measured geometry and raster comparisons, not just matching XML
or SVG attributes. Full parity cannot be declared while any required feature
is missing, unverified or covered only by a narrower comparison.
