# Office suite parity review

Last reconciled 2026-10-26 for shared charts, SmartArt and DrawingML; other headline areas reviewed 2026-10-22: the headline table below was re-checked against the later rounds in this file and
the code; the per-batch sections are dated by the work they describe (through round 6, 2026-10-16),
and their test counts are snapshots. Per-product trackers: [xlsx-parity-status.md](xlsx-parity-status.md),
[word-parity-review.md](word-parity-review.md), [visio-parity-review.md](visio-parity-review.md).
This is an evidence-based source review and tested implementation batch, not
a claim of Microsoft Office equivalence. The consolidated repository contains
all engines and product editors (Word, PowerPoint, Excel, Visio and Teams), with framework bindings
under `viewers/`.

The [shared Office delivery ledger](office-shared-parity-delivery.md) records current chart,
SmartArt and DrawingML implementation, browser/native evidence and remaining migration limits.
Dated batches below retain their original scope and counts.

## Current state

| Product    | Working foundation                                                                                                                                                                                                                                               | Largest remaining workflow gaps                                                                                                                                                                                                     |
| ---------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Word       | OOXML editing and preservation, section pagination, continuous section bands, column balancing, tables, footnotes, review features, editable Print Layout, shared chart/ChartEx and SmartArt editing, authority-based and Yjs collaboration                      | Dedicated page-story (header/footer and note) editing in Print Layout, vertically justified sections, text-box pagination, broader drawing/layout fidelity and complex-table structural editing                                     |
| PowerPoint | Presentation model and editing, shared rendering and ribbon, slide sorter, selected-slide transport between decks through the system clipboard, multiple framework bindings, import/export and conversion                                                        | Full text layout in raster export, clipboard interoperability with desktop Office beyond slide packages, inherited master-text search, remaining animation target modes and broader native render comparisons                       |
| Excel      | Workbook read/write, recalculation and dynamic arrays, editing/history, grid, charts, formula discovery, bounded PivotTables with slicers, timelines and number/date groups, sparkline authoring, Page Layout and Page Break Preview, shared editor and bindings | PivotTable label and value filters, shared caches, XLSB and ODS imports, FORECAST.ETS, sparkline style galleries and grouping, remaining header formatting and drawing fidelity; see [xlsx-parity-status.md](xlsx-parity-status.md) |
| Visio      | VSDX import and source-backed edits, shapes, grouping, connectors, routing, ShapeSheet subset, page setup, data/metadata, themes, layers, shared ribbon and framework bindings                                                                                   | Full ShapeSheet behavior, complete native geometry/text fidelity, unmeasured routing profiles, parts of data graphics and accessibility/proofing workflows                                                                          |

Evidence entry points:

- Word: `src/core/docx/layout/{page-flow,balance-columns,adapter}.ts`,
  `src/ui/src/docx/{print-layout-view,table-model-adapter,image-media,equation-schema}.ts`,
  `src/core/docx/ui/yjs-collaboration.ts`.
- PowerPoint: `src/core/pptx/converter/ShapeTextRenderer.ts`,
  `src/ui/src/pptx/render/slide-sorter-state.ts` and the shared search/editor modules.
- Excel: [the Excel status page](xlsx-parity-status.md),
  `src/core/xlsx/read/load.ts`, `src/core/xlsx/load/detect.ts`.
- Visio: `src/core/visio/{edit,formula,line-jumps}.ts`,
  `src/ui/src/visio/page-setup-dialog-form.ts` and
  [the capability ledger](../viewers/visio/docs/parity.md).

Existing capability documents sometimes lag implementation. Disabled ribbon
controls and package manifests alone are insufficient evidence of support.
Check the command, renderer, save path and tests together.

## Implemented in this batch

### Word

Centered and bottom-aligned page content now reserves footnote heights and the
separator. Odd/even section padding pages no longer copy the preceding page's
footnotes or floating objects. Core regressions and the actual Print Layout DOM
check overlap and duplicate-note behavior. Real Chromium verified note spacing
and an odd-page blank using current renderer source and CSS.

Long footnotes now continue at complete line boundaries across physical pages,
retaining original note IDs, formatting and source ranges. Continuations reserve
space before later body paragraphs and table rows, finish after body content ends,
and use full-width continuation rules. A shared double-offset error in stacked
paragraph lines is corrected. Browser measurement checked 28 unique note lines
across five pages; save/reopen preserves one original note and custom separator XML.
Footnote reservations now follow the actual reference line across paragraph
fragments, including generated list labels, hidden runs and widow/keep handling.
The focused Word suite passed 1,234 tests and six framework browser edit/history/
save/reopen workflows. Custom footnote and continuation separator stories now
retain mixed text and rule markers, reserve their measured height and render in
Print Layout. Empty separator stories suppress the default rule. Core and UI
regressions and six framework browser workflows pass. Per-column notes and tables
inside separator stories remain incomplete. Owned native Word automation is now available; compatibility-specific note-column comparisons are in progress.

Editable Print Layout now routes typing, selection, clipboard input and composition through the existing editor model, with projected caret and selection geometry on paginated pages. Six framework browser suites cover editing, zoom, later pages and columns, read-only mode, and disconnect/reconnect. Immutable document and rendered-line indexes avoid rescanning the document on each caret movement. PageUp/PageDown, Ctrl+Page navigation and edge drag autoscroll now have browser regressions at 100% and 150% zoom, including release and disposal. Dedicated page-story editing remains incomplete.

### PowerPoint

Raster text export honors explicit zero body insets, inherited body styling,
soft breaks, embedded newlines and empty paragraphs. The slide sorter supports
additive Ctrl/Meta+Shift range selection, preserving the selection anchor.
The common sorter implementation was exercised in React, Vue, Angular,
Svelte and vanilla browser demos.

Find/replace now spans formatting runs and descends into grouped objects, table
cells and SmartArt nodes. Replacement preserves the starting run's formatting;
paragraph and soft-break boundaries remain search boundaries. Save/reopen tests
cover edited nested shapes and rich table text, including hyperlinks and opaque
XML. Slide sorter paste retains copied-content snapshots after source edits or
deletion and inserts in clipboard order after the selected destination in one
history transaction. The original shortcut clipboard remains local to the sorter session. Separate Copy to System Clipboard and Paste from System Clipboard actions now transport selected-slide packages between suite decks in all five bindings. They preserve images, rich notes and independent themes, with undo/redo. Native PowerPoint opens selected-only, imported and core-resaved acceptance decks. This custom browser format does not establish desktop PowerPoint clipboard interoperability. Different slide dimensions, embedded fonts, custom table styles, comments and references to unselected slides are explicitly refused. Chart text scopes are described below; inherited master text remains outside search.

An explicit Include Speaker Notes search scope is now available in all five
PowerPoint bindings. Replacement retains rich note runs, honors subsequent plain
note edits and updates the native notes part on save. Eleven focused core tests,
five note-search browser workflows and the prior five rich-search workflows pass.
Native PowerPoint exposed a new-notes packaging failure: generated notes masters
now receive an independent theme part, including outgoing theme relationships.
Native reopen accepts the generated rich-notes fixture after this repair.

### Excel

`REGEXEXTRACT` supports all matches as a vertical spill and capture groups as a
horizontal spill. Insert Function documents and discovers these modes.
Tests cover recalculation, grid values, blocked/shrinking spills, undo/redo,
XLSX round trips and browser insertion. Source demo aliases now include the
PowerPoint DOM export needed by the shared UI dependency graph.

Sparkline authoring now supports Line, Column and Win/Loss groups, group settings,
selected/group clearing and history through the existing preserved XML model.
The translated Insert controls expose source/location ranges, marker flags,
colors, empty-cell display and axis bounds. Native Excel opens newly authored,
cleared, undone and extended imported groups, with exact extension URI casing,
namespace declarations and revision identifiers verified. Keyboard history after
menu actions restores grid focus. Style galleries, grouping/ungrouping, date-axis
and source-range editing remain incomplete.

Excel now authors, configures, refreshes and deletes range-based PivotTables with
one row field, an optional column field and one Sum, Count, Average, Min or Max
measure. Ribbon dialogs use one undoable edit transaction and save native table,
cache definition and cache record parts. Sixteen pivot regressions cover output,
history, save/reopen and preservation; the wider Excel suite passes 6,628 tests.
Browser creation, settings, refresh, save/reopen, deletion and history pass. Eight
native Excel acceptance cases include source changes followed by native refresh.
Advanced layouts, unknown options/extensions, custom field formats and shared caches
stay preserved with warnings; slicer-connected workbooks are admitted for the
supported profiles (see the slicer section below). Names from preserved
unsupported tables are reserved to prevent authoring collisions.

Excel print preview and printing now share a core page plan for paper size,
margins, orientation, scaling, fit-to-page and page order. Manual row/column
breaks support insert, remove, reset and history. Native Excel and Chromium
four-page outputs pass. Leading repeated row/column titles now reserve page capacity, and Page Setup authors native scoped title names. Header/footer sections support page count/number, sheet/file/date/time fields and first/even variants, with raw-code editing that preserves untouched imported formats. Native Excel and browser save/reopen/PDF workflows validate four pages. Non-leading titles, header formatting/images, page offsets and drawings remain explicit preview limits; locale-dependent date text can differ. Merge rendering now indexes covered cells once instead of scanning all merges for every cell.

### Visio

Connector-specific saved jump styles and horizontal/vertical directions are
read through the existing inheritance pipeline. Page directions can be saved,
and Page Setup has a Layout and Routing tab for jump placement, style and
directions, applied as one undoable transaction. An existing native enum error
is corrected: page `LineJumpStyle` values 0 and 1 both mean Arc, 2 means Gap,
3 means Square and 9 means seven sides. Print snapshots retain the page layout,
connector route and glue metadata needed to draw the same jumps as the canvas.

Connector-specific `ConLineJumpCode` is now read and rendered, including Always
when page jumps are disabled. Conflicting Always/Other policies follow the page
setting. A recorded Visio 16 SVG corpus tests all 125 combinations of five page
policies and both connector policies. Format Shape exposes connector-only jump
policy, style and directions through the existing source-backed undo/save path.
Native Visio reopened an edited dynamic connector with values 2/3/2/1 intact.
The recorder is `scripts/record-visio-line-jump-policy.ps1`.

Jump sizing now uses saved axis factors, routing spacing and drawing scale,
matching 68 recorded Visio 16 cases. Page Setup exposes horizontal/vertical
size percentages. See `scripts/record-visio-line-jump-sizing.ps1`.
An additional 160-case native corpus establishes predominant-axis behavior for
slanted straight connectors and confirms that curved connectors omit jumps.
Exact 45-degree ties remain sensitive to native endpoint rounding.

Native saved drawings cache jump decorations in Geometry rows. Straight routes
now rebuild that decoration for display and refresh supported plain caches on
save. Native reopen/SVG checks verify disabled jumps, gaps, larger arcs and
reversed directions. This also corrects the ArcTo sweep mapping into local y-up
coordinates. Page edits admit independent native connector text/self-trigger
formulas while retaining refusal of affected or unknown dependencies. Unknown
row metadata and geometry references retain their caches with diagnostics.

Native Shape Data dependency indexing now maps XML Property sections to the
Prop formula prefix across local and master scopes. A native two-shape fixture
verifies resize, transitive cache updates and Visio reopen. Native data-graphic
groups still require broader dynamic formula and text-field recalculation;
generated local 2D text, bars and icons now follow owner translations through
native pin formulas, with eight native movement observations and save/history
regressions. They remain separate shapes; native masters, stencil/group owners,
rotation-relative positioning and automatic value refresh remain incomplete.

The existing last-routed approximation uses drawing order. Native routing
chronology and broader geometry/text equivalence remain unverified.

## Validation and acceptance boundaries

The combined core run passed 30,280 tests across 1,761 files, with 16 expected
failures and 430 skipped tests. Core and UI builds and typechecks passed.
Focused UI, framework and browser checks supplement the core run; skipped
coverage is not evidence of support.
PowerPoint's 15 browser scenarios passed across React, Vue, Angular, Svelte and
vanilla (14 in the final matrix run and one corrected locator rerun). They cover
sorter selection, destination paste/history, rich find/replace/history and saved
XML preservation. The run also exposed and fixed vanilla sorter focus and
undo/redo refresh behavior.

Regression tests exercise the owning core and UI layers. Browser checks verify
real rendering and interactions, but do not replace native Office comparisons.
Word's Chromium check exercises the print renderer, not every editor binding.
Installed Excel does not support REGEXEXTRACT, so it cannot establish native
calculation acceptance. The regex engine still uses JavaScript syntax rather
than all PCRE2 features.

For each subsequent feature, require a real authored file where possible,
import assertions, edit/undo/redo behavior, preservation and save/reopen tests,
UI reachability, keyboard behavior and appropriate framework coverage. For
visual claims, record native Office version, fonts and reference renders.
Keep unimplemented controls explicit and avoid aggregate parity percentages
without a defined, weighted and reproducible test set.

Visio now reads and edits native shape accessibility titles and descriptions,
including local stencil overrides, through Review > Alt Text with normal history.
The canvas and static SVG expose those descriptions. Check Accessibility finds
visible graphics without text alternatives and links them to the editing dialog;
it does not validate reading order or contrast. Native Visio 16 acceptance verifies
setting and clearing both fields, including multiline descriptions. Formula-driven
metadata and dependent caches are refused. Page alternatives are now exposed through
Review > Page Alt Text and validated with native page property getters. The complete
Office accessibility rule set remains outstanding. Reproduce native acceptance with
`scripts/visio-alt-text-acceptance.ps1`; core, UI and Chromium regressions cover the
workflow. The writer now preserves tabs/newlines in XML attributes as character
references, with tests for unrelated attributes, comments and CDATA.

## Next priorities

Separate command coverage from end-to-end readiness: open, render, edit, save and reopen.
Feature lists of other products are not proof of fidelity or a measure of this project's parity.

Recommended next batches:

1. Expand native Office acceptance corpora for the most-used document workflows.
2. Word nested and vertically merged table pagination, page-story editing and continuous sections with notes.
3. PowerPoint native rendering comparisons, complex chart-source editing and broader clipboard preservation.
4. Excel multi-measure pivot layouts, grouping and remaining print formatting/drawing fidelity.
5. Visio geometry/text comparisons and curved connectors, followed by remaining data graphics controls.

Microsoft's [page jump style enum](https://learn.microsoft.com/en-us/office/client-developer/visio/linejumpstyle-cell-page-layout-section),
[connector jump style enum](https://learn.microsoft.com/en-us/office/client-developer/visio/conlinejumpstyle-cell-shape-layout-section)
and [horizontal jump direction](https://learn.microsoft.com/en-us/office/client-developer/visio/conlinejumpdirx-cell-shape-layout-section)
define the Visio mappings used here. The Excel review links the official function contract.

Native Visio inherited string fields display instance Field.Value caches with FIELDPICTURE(0), including empty and numeric-looking strings, instead of the master's placeholder. Pictures 37, 38 and 39 now support normal, lower and upper casing using mappings observed in Visio 16.0.20430.20146. A generated Unicode corpus and synthetic inherited-field fixture cover native differences from JavaScript casing, including sharp s and dotted I. Other internal field pictures retain their stored display pending explicit support.

Word compatibility mode 15 supports independent, balanced footnote columns, including continuation pages and explicit section footnote column counts. Native Word fixtures and PDF measurements cover equal and unequal body columns, default and explicit counts, and long notes. The More Columns dialog edits and saves the count. Modes 12 and 14 attach notes to their reference columns and continue through later columns and pages with independently positioned continuation separators. Full and partial continuation cases also verify body placement. The native corpus contains 31 cases; unverified compatibility modes report a page-wide fallback rather than claiming equivalent layout. Complex separator tables, continuous sections with notes and page-story editing remain gaps.

Excel sparkline date axes now have authoring and settings controls, chronological point ordering, paired hidden-cell filtering and native-measured column widths. The saved date range supports clear and undo. Native Excel acceptance covers these paths; fractional and out-of-calendar serials remain independently unvalidated.

PivotTables support tabular, compact and outline layouts with one or two row fields, one measure and an optional column axis. Compact/outline layouts offer top, bottom or no automatic subtotals; outline/tabular layouts support repeated labels. Native Excel acceptance compares every cached and refreshed output cell, type, dimension and compact indentation across 46 cases. More than two row fields, custom subtotals/indentation, grouping, mixed field layouts and shared caches remain unsupported and preserved. Multiple measures and bounded Values-on-rows placements are covered by the later report-layout batches below. Conflicting global layout flags are rejected for editing rather than normalized on refresh. Typed-cache flags and text labels are also covered.

Excel print headers and footers now render styled runs for font family, font size, RGB color and bold/italic/underline/strike toggles, with formatting controls in Page Setup. Native Excel PDF checks verify font, size and color, and browser checks cover save/reopen, computed preview styles and a four-page PDF. Theme colors, images, page offsets, double underline, superscript/subscript and unknown font styles remain preserved but unsupported in this preview.

PowerPoint chart titles are searchable across all five bindings. Simple rich titles support replacement, history and saving while preserving authored run/paragraph XML and embedded workbook bytes. Linked, ChartEx, shared-part and complex titles are read-only search targets; replacement rejects them atomically, and the chart-title scope can be excluded. Complex field values are not comprehensively indexed. Series/category replacement is covered in the follow-up below. Ten browser cases and five native PowerPoint reopen cases cover this batch. Speaker-note and clipboard controls use the shared UI components and translated labels.

Visio curved walking glue matches the overlap midpoint of facing, axis-aligned sides, including unequal target sizes and reversed directions. Single-corner routes between axis-aligned shapes now use the native five controls and two NURBS spans. Seventeen native captures establish the routing boundary; ten edited drawings reopen and save in Visio with matching endpoints. The six corner cases also match control points and knots, with text placement within 0.0002 inches using bounded curve flattening. Browser coverage verifies a two-span drag preview, saved geometry and undo/redo. Rotated targets, fixed endpoints and obstacle detours retain the general routing approximation. Reproduce with `scripts/record-visio-curved-route.ps1`, the `VISIO_NATIVE_CURVE_DIR` regression and `-VerifyCore` acceptance. The native row follows Microsoft's [NURBSTo field contract](https://learn.microsoft.com/en-us/office/client-developer/visio/cell-element-nurbsto-rowvisio-xml).

Body-table footnote references now reserve note space during row splitting, preserve clipped logical references and avoid duplicating stories for repeated headers. Twenty-three native layout cases cover modern and legacy columns, long notes, late markers, row heights and padding, including partially split padded cells. Split fragments retain cell margins and allow independent cuts for unreferenced neighboring cells. Core round-trip, browser editing and edited native reopen/PDF checks cover these paths. The nested-table follow-up below projects nested text and references; vertically merged cells can still paginate differently. Complex interacting long notes remain unverified.

PowerPoint title-only saves preserve embedded workbook bytes, including rich shared strings. Repeated saves track chart XML and workbook reconciliation separately, so saving data, undoing it and changing a title restores the workbook source cells. Failed or partial workbook writes remain eligible for retry on a later save.

Nested Word table text and footnote references now follow source child order in Print Layout, including continued rows and references in a second cell. The native body-table corpus now contains 27 cases; three nested edited exports reopen in Word with matching page and note counts. Core and browser checks cover local text bounds, source preservation, history and saving. Structural edits in complex source tables are blocked before history when the serializer cannot preserve them; inline text and soft breaks remain editable. Nested borders, shading, vertical merges, inner repeated headers and legacy compatibility layouts remain incomplete. The final Word core suite passed 1,317 tests, with 21 browser cases for the guard and related editing flows.

PowerPoint has an opt-in series/category label scope across all five bindings. Literal labels and private, coherent embedded workbook sources support replacement and history; saving updates chart caches and source cells together. Formula-dependent, external, shared, missing, inconsistent, rich-string, date, multilevel and ChartEx sources are read-only. Newly authored classic category charts support label edits before first save, and loaded charts become read-only for this workflow when their source-grid dimensions change. Twenty browser cases and four native PowerPoint cases cover labels and the existing title workflow, including workbook activation, refresh and reopen. These checks do not establish general chart-data or rendering parity.

### Visio inherited decimal fields (native verified)

- Added FIELDPICTURE 0 through 9 and 15 through 18 display for inherited instance numeric values, retaining explicit inches, millimeters, centimeters, feet, degrees and radians.
- All 64 synthetic native Visio cases match core parsing and unrelated edit/save/reopen. Browser coverage checks each rendered field and export/reopen.
- Reuses the existing core number-picture formatter. No formatter logic is duplicated in UI or bindings. Other internal numeric/date pictures and page-dependent units remain cached and unsupported.

- Boundary verification adds 104 native cases: half-decimal rounding, zero, small-number scientific mantissa ties/carries and large decimals. Another 46 native general-number cases verify codes 0/1. Another 120 fractional cases verify nearest one/two-digit denominators, including profiles that differ from Excel. All 334 fields match core edit, browser export and native resave/reopen. Integer overflow and general magnitudes above 1e15 remain cached; page-dependent units and locale pictures remain unsupported.

### Nested Word table frames and native border occupancy

Nested Print Layout tables retain resolved row/cell geometry and paint borders and shading through the existing table renderer. Fixed-width alignment and center/bottom text alignment use shared core geometry. Decorations do not create another editable story or intercept text hit targets.

Nineteen synthetic native Word profiles cover one/two rows, centered placement, fixed-height vertical alignment, uniform 0.5/1/2/3 point solid borders, asymmetric consistent solid edges, margins, bordered enclosing cells, recursive nesting and split rows. Ink positions match within 0.15 points and frame dimensions within 0.2 points; the split cases match native 16/4 and unequal 16/16 then 4/2 line placement. Continuation height uses the greatest parallel story bottom instead of the last source box. Core checks source XML preservation for every fixture. Validation includes 1,372 Word tests, 183 print UI tests and browser editing guard/history/export/reopen cases.

Border reservation is admitted for the measured complete solid-edge profiles with matching shared width/color and stable row grids. Missing edges, conflicting shared borders and other styles report approximate occupancy. Decorative frame layers preserve external horizontal strokes, source ink origins and vertical row clips. Ordinary enclosing row border-height reservation is covered by the later thirteen-profile ordinary-row corpus below; interacting nested profiles outside the bounded inner-story admission below remain approximate. Inner vertical merges, repeated inner headers, glyph metrics and rasterization remain limits. Native evidence and regeneration instructions are in [the nested-frame corpus](../src/core/docx/__fixtures__/nested-table-frames/README.md).

### Excel multiple-measure report layouts

PivotTables now support up to sixteen Values-on-columns measures in tabular, compact and outline reports, with one/two row fields and an optional column field. The model, per-measure aggregation, headers, source cache and Values-axis XML remain in core; the shared UI configures fields, aggregates, captions and layout. Refresh, undo/redo and save/reopen reuse existing editing transactions.

One hundred twenty-five native Excel cases compare every cached/refreshed cell value, type and dimension, including repeated measures, errors, sparse intersections, subtotals and label repetition. Existing x14 Values-header extensions are normalized on conversion. Output budgets are checked before header allocation. Validation includes 6,826 Excel tests, 379 UI tests and repeated production browser workflows.

Values-on-rows are supported in tabular, compact and outline reports with one real row field, Values last, multiple measures and an optional column field. Native import, refresh, conversion between axes, unchanged settings, history and browser export/reopen are covered. Compact and outline group headings and indentation are native verified. Other row-field placements, visible/custom Values-header variants, grouping, filters, slicers and shared cache profiles remain preserved unsupported cases. The native acceptance script is `scripts/xlsx-pivot-acceptance.ps1`.

### PowerPoint picture effect extraction

Animated picture frames paint above their image pixels while retaining the crop and external shadow. The duplicate placement decision and background paint-layer CSS move from React's interactive/static renderers into `src/ui/src/pptx/render/picture-frame.ts`, with provenance recorded. Ten static/animated browser cases across all five bindings and focused renderer regressions pass. Broader animation targets and native slideshow pixel comparisons remain separate limits.

### PowerPoint background animation extraction

Supported shape/text background targets now select and split their paint through shared UI descriptors used across all five bindings. Background fill/stroke animation leaves foreground text visible; independent whole-element opacity remains on its positioning box. Svelte presentation exit clears animation state. Ten browser cases verify background color, opacity ramps and stop/re-entry across all bindings; native PowerPoint preserves the independent effects and target XML on resave/reopen.

This establishes supported behavior and package preservation, not native slideshow pixel equivalence. Background-only fixed p:set opacity, chart/SmartArt targets and broader animation combinations remain gaps.

### PowerPoint whole-shape fixed opacity

Finite, isolated whole-shape p:set style.opacity now uses shared core admission and shared UI playback/rendering across all five bindings. Plain decimal numeric strings and finite numeric fltVal values are accepted; scientific-notation strings are refused. Missing effect/behavior fill restores the authored opacity after the active interval. Explicit hold or freeze on both retains it; remove on either restores it. Delayed starts, seek, backward end-state projection and presentation exit/reset are covered.

Native PowerPoint 16 video samples from synthetic whole-shape fixtures distinguish transparent fill and gray foreground text under hold from restored opaque fill and black text under remove. Package preservation and native rendered playback are separate checks. Ten browser cases pass across all five bindings; 40 focused parser/resolver/round-trip tests, shared opacity playback/style/timeline tests and per-binding regressions cover the supported slice. Reproduce with the generator and native video script in e2e/pptx/native.

Background-only numeric sets remain refused because measured native composition differs from simple CSS paint/body multiplication. Text/subshape targets, hover, repeated/reversed/additive or mixed behavior trees, unresolved event/indefinite child timing, indefinite effect duration, typed/raw end conditions, end sync, time filters and multiple or unsupported legacy image hints remain preserved but unsupported. This bounded slice does not establish general animation or slideshow pixel parity.

### Visio inherited ISO date fields

Inherited DATE fields regenerate ISO date-only text for the native-verified explicit FIELDPICTURE(204) profile. System short-date picture 20 is admitted only with a valid ISO cached placeholder; other locale, invalid or empty caches remain preserved. Two native corpora cover 22 admitted dates and 24 source edit/save/reopen preservation cases covering unsupported pre-epoch dates. Other locale pictures remain preserved and unsupported without native acceptance in those 24 cases. Core parsing and text-field regressions, 3,831 Visio tests and two browser cases validate this bounded scope. Explicit picture 204 does not broaden picture 20 to other system locales.

### Word ordinary table border occupancy

Ordinary table rows reserve measured border occupancy through the core layout helper and shared print UI border-width interpretation. Thirteen supported native profiles cover ordinary and enclosing auto/exact/clipped borders, early/late splits, margins and repeated headers. The original independent nested-cut diagnostic is covered by the bounded inner-story profiles below; other complex nested cuts remain approximate. Validation includes 1,387 Word tests, 184 print UI tests, 20 final repeated browser cases and 37 related browser cases. Native fixtures and measurement scripts preserve the distinction between supported ordinary rows and unsupported interacting nested border profiles.

### Excel nested tabular Values-on-rows reports

Tabular PivotTables now support two distinct real row fields followed by Values, with two to sixteen measures and an optional column field. Shared core hierarchy indexing, per-measure source aggregation, header dimensions and row-item serialization drive creation, native import and refresh. Bottom or disabled automatic outer subtotals and repeated row labels are supported. Native Excel displays tabular top-subtotal flags below the group as well. Compact/outline nested Values-on-rows reports are covered by the later native report-layout batch below. Other unmodeled Values/header profiles, grouping, filters, slicers and shared cache profiles remain preserved unsupported cases.

Excel 16.0.20430.20146 accepts all 140 cached/refreshed cases, including fifteen new cases for typed output, errors, sparse intersections, repeated measures, imported changes and transitions to columns, one row field or one measure. An eighteen-profile native oracle establishes layout boundaries. Every cached/refreshed cell value, type and report dimension is checked. The output budget rejects oversized reports before header allocation. Validation includes 6,834 Excel core tests, 379 shared Excel UI tests, strict typechecks, six repeated new browser workflows and nine final combined browser cases covering authoring, history, save/reopen and preservation. Reproduce with `scripts/xlsx-pivot-acceptance.ps1` and `scripts/xlsx-pivot-nested-row-oracle.ps1`; frozen native provenance is recorded in the nested-row evidence JSON.

### Word independent nested footnote stories

Fifteen owned native Word probes establish independent continuation for one automatic, top-aligned inner row, including the original single reference and ten two/three-reference cases. The complete enclosing row is admitted with at most three distinct anchored references. An already placed marker stops constraining its story; a deferred marker stops only its own story immediately before its reference line. Native cases cover either referenced inner cell, two notes in one inner story, neighboring enclosing-cell markers and different pending markers on the same physical page. Following direct outer text starts after the longest remaining story, and a trailing-only page does not reserve a completed inner border.

The planner retains the shared authoritative cell-paragraph projection, including blocks-only inputs, and uses existing source-token reflow and shared border geometry. Dynamic expansion is confined to admitted inner stories; other paths retain the previous cell-cut behavior. Source structural editing guards and exact-row partial glyph clipping remain intact. Native regressions compare per-page body/note marker sets, every body and note line baseline, and physical border bands within 0.2 pt. Exact document and footnote XML and saved/reopened models remain unchanged without edits. Validation includes 1,422 Word tests, 102 repeated related browser cases and 20 final-source browser cases covering edit, undo/redo and save/reopen, plus strict core and targeted browser typechecks.

Admission excludes reused marker IDs, multiple unanchored notes, four references in the complete enclosing row, multiple inner rows, deeper nesting, fixed-height or centered/bottom-aligned inner rows, inner repeated headers, leading outer paragraphs, noted trailing paragraphs and explicit paragraph keeping/page-break rules. The ten cases in this batch use three-line note bodies. Seven long interacting note-body continuations are covered by the subsequent acceptance corpus below; broader note formatting and column profiles remain unverified. Vertical merges and font raster fidelity remain unverified. See [the single-reference corpus](../src/core/docx/__fixtures__/independent-inner-notes/README.md) and [the multiple-reference corpus](../src/core/docx/__fixtures__/multiple-inner-notes/README.md), with reproducible synthetic generators and owned native acceptance scripts.

### Visio ISO date authoring and native field caches

The shared field dialog offers an explicit ISO date picture. Creation-date fields now write timezone-free ISO DATE caches and string-typed Format caches, preserving formulas. Before this correction, core displayed dates but native Visio displayed `0` for the numeric cache. Native Visio 16 accepts all five short, ISO, long, time and combined date/time creation-field formats through resave/reopen. A browser workflow verifies ISO insertion, history and export/reopen. Host time strings are Australia/Brisbane evidence, not timezone-independent expected values.

This validates the creation-date writer and these five formats. Other date functions, locale calendars, custom shape formulas and general field-editing parity remain separate evidence scopes. Reproduce with `scripts/visio-date-field-write-acceptance.ps1`; [native writer evidence](../src/core/visio/__fixtures__/date-field-write-native.md) records cache syntax and host-zone limits.

### Excel compact and outline nested row Values

Compact and outline reports now support two distinct real row fields followed by Values, up to sixteen measures and an optional column field. The existing indexed hierarchy engine emits outer/inner headings and transposed measure rows. Compact rows use native indentation levels zero, one and two through cached shared style patchers. Native Excel places automatic subtotals below these row-Values groups for both top and bottom flags; the dialog explains this behavior and disables an ineffective placement control. Repeated labels remain editable in outline/tabular reports.

Excel 16.0.20430.20146 accepts all 179 cached/refreshed cases, including 39 additional layout cases. These compare every cell value/type, dimensions, compact indentation and left alignment, sixteen-measure reports, all six layout conversions, imported source changes and collapse to one row or measure. Validation includes 6,853 core tests, 381 shared Excel UI tests, strict core/UI and targeted browser typechecks, 26 repeated authoring/history/save workflows, six repeated native nested-file unchanged-settings preservation cases, and sixteen final combined browser cases. Output budgets include both heading levels before allocation. Custom Values-header profiles, custom indentation/subtotals, grouping, filters, slicers, shared cache and visible/custom Values-header profiles remain preserved unsupported cases. Native provenance and three owned XLSX fixtures are recorded in `native-pivot-nested-row-layouts-evidence.json`; reproduce with `scripts/xlsx-pivot-acceptance.ps1`.

### Word long interacting inner note continuations

Seven additional owned native Word probes establish the existing layout behavior for forty-paragraph long notes in the bounded automatic, top-aligned inner-row profile. They cover a long first or second inner note, two long notes, consumed markers followed by inner or enclosing-cell references, two notes in one inner story and three notes with long first/last bodies. The reports span four to eight pages, including note-only pages and later body resumption. This acceptance batch changes fixtures, measurement scripts and tests only; it does not add product layout logic.

Every page's body/note marker sets, line baselines and physical border bands match within 0.2 points, with checks for continuation flags, separator rules, body ink positions and no table frames on note-only pages. Exact document/footnote XML and saved/reopened source models are preserved. Validation includes 1,429 Word core tests and fourteen repeated browser cases covering editing, history, export/reopen and continued-note rendering. The measured profile uses one automatic top-aligned inner row, stable complete solid borders, single-column bottom notes, fixed twelve-point line spacing and up to three distinct anchored references. Reused IDs, multiple unanchored notes, four references, complex inner structures, independently balanced long-note columns, altered separator stories, other paragraph-keeping/formatting profiles and glyph raster fidelity remain outside this evidence. See [the long-note corpus](../src/core/docx/__fixtures__/long-inner-notes/README.md) and its synthetic generator and owned native recorder.

### PowerPoint imported whole-shape after-effect clearing

Explicit Don't dim now clears recognized imported RGB/theme dim, hide-after-animation and hide-on-next-click actions. Undefined remains an untouched native state. DOM-free authoring, native/editor precedence and transient imported-control eligibility move into PowerPoint core, with shared UI compatibility exports and provenance. Controls use the same state descriptor across all five bindings; unresolved native theme colors expose clearing without an invented RGB swatch.

Validation includes 865 animation regressions, 111 focused preservation tests, 100 shared UI tests and twenty actual browser workflows across React, Vue, Angular, Svelte and Vanilla for clearing, history, export/reload and playback. Fifteen owned native PowerPoint cases cover original, embedded-sound and following-click fixtures through save/reopen, preserving effect counts, other after-states, media bytes and audio relationships. Native RGB video samples distinguish the source dim from the cleared fill and foreground text. This does not establish general dim rendering or entrance timing parity: imported appearance effects with no modeled duration retain the existing 500 ms viewer fade fallback. Supported imported controls require one whole-shape entrance/emphasis effect matched by native node ID and target; ambiguous effects, unknown after-state, background/text/subshape targets and missing clear targets are refused or retained without an imported control.

### Word final inner-note tails in independent columns

Eleven owned native Word probes extend the interacting inner-note corpus to forty/sixty-line bodies in two or three independent note columns. The bounded correction reserves the shared geometry minimum needed to finish a single remaining unreferenced inner-story line beside one pending enclosing-cell marker. Longer stories retain their existing reference-band constraint. Source structural/reference admission gates, single-column behavior and the global note-column preview are unchanged.

Validation includes 1,438 Word core tests and 46 browser cases, with native body/note sets, baselines, column positions and physical borders compared within 0.2 points across all eleven profiles. Exact document/footnote XML and reopened source models are preserved. The measured profile has one automatic top-aligned inner row, complete stable solid borders, one body column, fixed twelve-point line spacing and up to three distinct anchored references. Other note-column counts, spacing/keeping and section profiles, reused IDs, multiple unanchored references, complex inner structures and raster fidelity remain outside this evidence. See [the independent-column corpus](../src/core/docx/__fixtures__/interacting-note-columns/README.md).

### Excel Values between nested row fields

PivotTables now support Values between two distinct real row fields in tabular, compact and outline layouts, with two to sixteen measures and an optional column axis. The existing indexed hierarchy, aggregation, header descriptors, output budgets and style helpers drive the measure-major rows. Shared field settings expose Last and Between placement. Native automatic outer subtotals render below each group; repeated real row labels do not repeat the Values caption.

Excel 16.0.20430.20146 accepts all 243 cached/refreshed cases, including 64 new placement cases. Checks compare every cell value/type, dimensions, indentation and alignment, errors, sparse sources, sixteen measures, layout conversions and placement/axis/collapse transitions. Validation includes 6,899 Excel core tests, 384 shared UI tests and 52 repeated browser authoring, source-refresh, history, export/reopen and unchanged-native-settings cases. Exact axis/data-position admission preserves malformed or unsupported sources. Values first is covered by the later native placement batch below; grouping, filters, slicers, shared caches and unmodeled header profiles remain unsupported. See `native-pivot-between-values-evidence.json` and `scripts/xlsx-pivot-acceptance.ps1` for owned native provenance.

### Visio inherited system-date cache profiles

Inherited master date fields now refresh from instance DATE caches for measured system picture 200 with a valid stored ISO date, and pictures 211/212 with the exact stored ISO date and English twelve-hour AM/PM minute/second suffix. The core display delegates to the existing cache parser and token formatter. Culture alone does not select a system picture; mismatched, invalid and pre-epoch source caches retain their original text. Explicit ISO picture 204 retains its existing profile.

Thirty owned Visio 16.0 cases compare native text, value formulas and picture formulas through native save/reopen, including midnight/noon, leap dates, epoch and fractional-second boundaries. Validation includes 3,848 Visio core tests and browser coverage for all thirty fields, rename history and export/reopen, with exact page/master XML preservation. This corpus does not establish long-date pictures; the following independently measured slice admits their strict English cache profile. Other system locales and picture codes remain unmeasured. See [the native corpus](../src/core/visio/__fixtures__/field-system-date-native.md) and its acceptance/reopen scripts.

### Visio inherited English long-date cache profiles

System pictures 21/201 now refresh inherited DATE fields only when the complete stored English `dddd, MMMM d, yyyy` calendar date and weekday validate through the existing token formatter. Machine culture does not choose a profile. Invalid dates or weekdays, abbreviated/reordered English, German/French and other locale caches remain preserved. Exact page/master XML stays unchanged through an unrelated rename, export and reopen.

Forty-four owned native Visio probes cover all twelve months and seven weekdays, leap/non-leap centuries, epoch and fractional times. Native core-export open/save/reopen matches forty-two admitted texts and all forty-four value/picture formulas. Two pre-epoch source cases fall outside admission: native year-30177 long-date text changes after reopening even the untouched native file. Their original and reopened observations are recorded without inferring a rendering rule. Validation includes 3,881 Visio core tests and eight repeated browser runs covering long/system/ISO dates, rename undo/redo and save/reopen. See [the native long-date corpus](../src/core/visio/__fixtures__/field-system-long-date-native.md). Other languages, English ordering/punctuation and internal picture codes remain outside this evidence.

### Excel Values before nested row fields

PivotTables now support Values before two distinct real row fields in tabular, compact and outline layouts, with two to sixteen measures and an optional column axis. The existing indexed hierarchy and single-measure aggregation emit measure-major groups; shared header budgets and style helpers govern dimensions and indentation. Native compact/outline first placement honors automatic top subtotals on each real-field heading, while tabular subtotals remain below. Bottom and disabled subtotals and repeated real-row labels are supported; native Values captions remain unrepeated.

Excel 16.0.20430.20146 accepts all 310 cached/refreshed cases, including 67 new first-placement authoring, source-refresh and conversion cases. Thirty-six owned native profiles establish top/bottom/disabled subtotals with and without repeated labels across all layouts and optional columns. Checks compare every cell value/type, dimensions, indentation/alignment, sparse/error sources and sixteen measures. Regression coverage includes native row-item prefix serialization, exact untouched source parts, malformed/inconsistent axis metadata refusal, history and before-allocation output budgets. Validation includes 6,938 Excel core tests, 387 shared UI tests, 72 repeated browser authoring/history/export/reopen and unchanged-native-settings cases, strict core/UI and targeted browser types, lint and formatting. Native source provenance is recorded in `native-pivot-first-values-evidence.json`; reproduce with `scripts/xlsx-pivot-acceptance.ps1` and the nested-row oracle. Grouping, filters, slicers, shared caches, mixed/custom layouts and visible/custom Values-header variants remain preserved unsupported profiles.

### PowerPoint bounded native visibility timing

DOM-free effect timing, fill and repeat decisions now live in PowerPoint core; the shared UI retains a compatibility export and provenance records the extraction. A native preset-zero whole-shape entrance with one held visibility set now uses its finite child delay and one-millisecond interval instead of an invented fade. All five bindings consume the shared discrete visibility timeline and CSS property ownership.

Three owned PowerPoint 16 save/reopen and video cases establish durations of 1, 500 and 1,500 ms. At 20 fps, a concurrent zero-delay marker distinguishes movie startup from target appearance: the measured target intervals are 0, 500 and 1,500 ms, with no intermediate fade pixels. Validation includes 918 core tests, 109 shared UI tests, twenty-five actual browser workflows covering timing and imported after-effect clearing, plus five final timing reruns across all bindings. Export/reload preserves native timing; seek completes the held state and stop/re-entry restores hiding.

Admission requires the measured whole-shape target, preset/subtype, explicit hold, one visibility behavior and ordinary concurrent seek sequence. Unsupported values, repeated/reversed/event/reference clocks, positive enclosing delays, unknown or mixed behavior trees, alternate targets/fills and intervals beyond the signed timer range remain preserved and reported unsupported. Native composite opacity/visibility timing on one shape remains unverified; these cases do not establish general animation parity. Reproduce with the visibility generator and native render script under `e2e/pptx/native`.

### Visio existing local text-field editing

The shared Field dialog can edit a selected existing local field formula and output format. The atomic command preserves its Field row IX, unique text marker, literal neighbors, other fields, text formatting, unknown source XML and untouched package parts. Source text, field cache and row proof guard against stale edits; dependency and protection admission remains atomic. Omitting format in the update API means General, as for insertion.

Core-projected edit eligibility excludes known inherited, grouped, nested, ambiguous, protected, nonstandard calendar and oversized source profiles from the chooser. Three owned native Visio numeric, creation-date and document-string profiles pass export, native open/save/reopen and formula/cache checks. The original feature passed 3,913 Visio core tests; the final eligibility correction passes 3,922 tests with 200 existing skips, plus eight focused shared UI tests and strict core typechecking. Eight repeated reachable browser cases cover numeric/date/string editing, UTC/Brisbane date behavior, undo/redo and export/reopen; four related insertion/date browser runs also pass. Literal DATE authoring and retaining an old literal DATE formula for a format-only edit remain unsupported. Other internal pictures and complex/inherited/grouped field editing are unmeasured. See [the native field-edit corpus](../src/core/visio/__fixtures__/field-edit-native.md).

### Excel bounded PivotTable report filters

PivotTables now support one report filter with All or one exact cached literal selection, one distinct real row field, one to sixteen Values-on-columns measures and an optional column axis in tabular, compact and outline layouts. Existing core aggregation projects selected records while preserving the complete source cache. Shared field settings expose typed cached choices; refresh, undo/redo and export/reopen use existing transactions. Report captions and the spacer row participate in output budgets, overlap checks and editing locks.

Excel 16.0.20430.20146 accepts all 395 cached/refreshed cases, including 85 report-filter cases. Thirteen owned native profiles establish All, selected text, numeric, boolean, blank and error captions; seven sanitized native XLSX files cover import and unchanged source preservation. Checks compare every output cell value/type, dimensions and field counts. Validation includes 7,005 Excel core tests, 393 shared UI tests and 98 repeated combined browser cases covering selection, source refresh, history, export/reopen and untouched native settings. Strict admission checks full-cache item/index/count metadata, typed literals, axes and report prefixes, and rejects missing selections atomically. Reproduce with `scripts/xlsx-pivot-acceptance.ps1` and `scripts/xlsx-pivot-page-oracle.ps1`; provenance is recorded in `native-pivot-page-filter-evidence.json`.

Multiple or dynamic report filters, nested real row fields with a filter, Values-on-rows filters, grouping, slicers, shared caches, mixed/custom layouts and ambiguous or unmeasured caption/cache forms remain preserved unsupported profiles. Native English captions are measured; this evidence does not establish other native locales or general filtering parity. Earlier unsupported-filter statements above apply to profiles outside this bounded report-filter slice.

### PowerPoint font editing extraction

Font-family face/token updates and paragraph insertion-style updates now live in PowerPoint core, with shared UI compatibility exports. One immutable core whole-text patch replaces repeated body/run merging in the shared selection-aware updater, inspector and Vanilla binding. Selection restoration, paragraph overrides and framework history stay in UI. Provenance records the original modules and commit; both moved implementations match their source except type imports.

Validation includes fourteen core rule, immutability, insertion-style and mixed-rich-run save/reopen tests, thirty-one shared UI selection/paragraph/inspector tests, eighty-four existing regressions across all five bindings and ten existing five-demo font-picker browser cases. Core/UI/binding typechecks, shipping builds and ESM/CJS imports pass. These checks establish extraction compatibility, not new native font semantics or rendering parity. The existing uniform-run writer policy is unchanged; native font-family behavior beyond prior evidence remains unmeasured.

### Word native two-row nested table fragments

Print Layout repeats the source nested table outer top and bottom edges at physical page cuts, instead of retaining an interior row edge. Continuations reuse shared row/cell layout, source tokens, frame clipping and border reservations. Admission is bounded to one two-row/two-cell automatic top-aligned nested table with direct note-free paragraph stories, resolved width and complete consistent solid borders, followed by direct enclosing text. Keeping, fixed heights, notes, merged/inconsistent edges, inner headers and deeper/multiple/extra rows retain existing behavior and explicit approximation diagnostics.

Five owned native Word probes match per-page body line sets and page counts, including three pages. Corrected frame heights differ by less than 0.2 PDF points, widths by less than 0.1 point, and relative projected ink baselines by less than 0.15 point for the measured Times New Roman profile. Ten core exports pass twenty native open/save/reopen checks. Validation includes 1,462 Word core tests, 36 shared DOM tests and sixteen repeated browser workflows. Browser checks use deterministic paragraph-start edits, exact print waits for undo/redo and old editor/page detachment before reopening. No-op exports preserve every source package entry/model; direct text edits preserve nested XML and other parts. These bounds do not establish glyph/raster parity or note-bearing multi-row continuation. See [the native two-row corpus](../src/core/docx/__fixtures__/nested-multi-row/README.md).

### PowerPoint paragraph and run style scope extraction

Body/run style partitioning, line-spacing mode exclusivity, authored paragraph/indent overrides and inherited decoration reconciliation now live in PowerPoint core. Shared UI retains compatibility exports; only the content-run predicate and decoration flag definitions move out of selection state. Selection ranges, pressed-state computation, DOM restoration and framework history stay in UI. Provenance records the original modules and commit, and the scope implementation matches its source except imports.

Validation includes nineteen core rules, marker/break immutability and paragraph save/reopen tests, forty-eight shared UI selection/inspector checks, ninety-four existing regressions across all five bindings, and fifteen existing five-demo list-history/fixed-spacing browser cases. Core/UI/binding typechecks, shipping builds and both new ESM/CJS imports pass. This establishes extraction compatibility and preserves prior editing behavior; it does not expand native paragraph semantics or establish native rendering parity.

### Word structural automatic nested row fragments

Print Layout now applies the existing nested physical-fragment reflow to structurally admitted tables with two or more automatic top-aligned two-cell rows. The guard retains resolved positive width, direct note-free paragraph stories, one leading nested table, complete consistent solid edges and direct enclosing text. Source-edge, row/token continuation, shared frame geometry and border reservations are reused; there is no additional layout engine or row-count whitelist. Earlier two-row evidence above remains historical.

Six three-row native probes cover thirteen pages, and twenty-two four/eight-row probes add forty-eight pages, including first/middle/last cuts, a ten-page all-long case and minimal/empty rows. Native physical fragments repeat the original top and bottom borders independently of source row count. Per-page body sets and page counts match; frame heights differ by less than 0.2 PDF points and widths by less than 0.1 point. The original first-ink bound remains below 0.15 points. A separate all-line relative ink bound is below 0.2 points, with measured maximum 0.176 also present in the three-row reference corpus. These assertions use the measured Times New Roman profile and do not establish glyph/raster parity.

The two new corpora produce fifty-six core exports and 112 owned native open/save/reopen checks. No-op exports preserve source package entries and models; direct enclosing-text edits preserve nested XML and all other parts. Validation includes 1,545 Word core tests, 64 shared DOM tests and 23 browser workflows, including eight repeated new cases plus the full fifteen-case shared family. Browser tests retain deterministic paragraph-start edits, exact rendered undo/redo waits, structural editing refusal and old editor/page detachment before fresh reopening. Strict core/UI and focused browser-spec types, shipping/source demo builds, lint and formatting pass. See the [three-row corpus](../src/core/docx/__fixtures__/nested-three-row/README.md) and [structural row corpus](../src/core/docx/__fixtures__/nested-structural-rows/README.md).

Native coverage is representative of 2/3/4/8-row structures, not an exhaustive proof for every table. Notes, merges, incomplete/conflicting borders, heights, repeated headers, paragraph keeping/breaks, non-top alignment, unresolved widths and deeper/multiple nested tables remain outside this admission. Multi-row note continuation and external note interactions remain separate evidence scopes. Existing two-row fixture bytes and repaired note browser specs are preserved.

### Excel cached report filters with nested row fields

The existing single cached report filter now combines with two distinct real row fields, Values on columns and one to sixteen measures in uniform tabular, compact and outline layouts. One optional column axis, top/bottom/disabled automatic outer subtotals and repeated real-row labels are measured. Selection is All or one exact cached text, number, boolean, blank or error literal. The existing full-cache projection, indexed hierarchy, row-item mapping, aggregation, header budgets and styles are reused. This extends the earlier flat-row report-filter scope without introducing another aggregation engine.

Excel 16.0.20430.20146 accepts 114 new authored/imported cached and refreshed cases, comparing every cell value/type, dimensions, row-item metadata and hierarchy formatting. The full acceptance corpus reached 509 cases before the repeated-column guard correction; all 114 new positive cases were rerun after that correction and on the publication dependencies. The committed oracle records 108 native layout/filter profiles and nine native workbooks, including one preserved unsupported column-repetition source. Validation includes 7,184 Excel core tests, 396 shared UI tests and 48 repeated browser authoring, literal/All selection, history, export and fresh reopen workflows, plus strict core/UI types and locale coverage.

Admission rejects column-caption repetition across every matching extension entry, accepting trimmed valid boolean forms in the local guard and refusing malformed explicit values. Public admission retains its stricter existing boolean profile; unrelated extensions are preserved. Multiple/dynamic report filters, grouping, slicers, shared caches, custom/mixed layouts and ambiguous or unmeasured cache/caption forms remain preserved unsupported. Values on rows now works under a report filter (round 6), checked against real Excel files. Native caption evidence is English. See [the native nested-filter corpus](../src/core/xlsx/__fixtures__/native-pivot-nested-page-evidence.json) and the acceptance/oracle scripts. Public feature help now states these measured bounds in all five locales.

### Visio page context menu setup routing

The Page Setup entry on empty-page context menus now routes to the existing shared dialog used by Design and Backstage. It no longer reports that core setup editing is missing. This is a UI entry-point correction: the existing planner, transaction, source, loading and protection guards remain unchanged.

Validation includes thirty focused shared UI tests, forty-nine related page setup/size/native-fixture core tests with four existing cases explicitly skipped, and four repeated browser workflows across the new context-menu path and existing Design path. Context coverage checks cancellation without edits, no-document/no-source refusal, a protected PageWidth refusal without an undo entry, rename undo/redo, export and old SVG detachment before fresh reopened settings. Shipping builds, full UI and viewer/browser-spec types, lint and formatting pass. This slice adds no printer tiling behavior or new native rendering claim; snapshots still disclose that saved printer paper, margins, orientation, tiling and scale are unsupported.

### PowerPoint explicit inherited line spacing and paragraph geometry ownership

Whole-text line-spacing commands now override conflicting parsed paragraph inheritance even when the chosen value equals the body or first-paragraph baseline. The existing context-free API and levels without parsed line-spacing metadata retain their previous behavior. Source inheritance snapshots, unrelated paragraph before/after spacing and indentation remain unchanged. Only the shared line-spacing ribbon opts into reapplying its displayed select value; ordinary selects remain change-only, and pointer/keyboard commits emit one input/change pair with disabled controls refused.

PowerPoint property and saved-XML measurements cover one proportional whole-text profile: paragraph 1 directly sets 1.2 and paragraph 2 inherits 1.5. Applying 1.2 to the whole text range makes both 1.2, adds paragraph 2's direct percentage and retains the list style. The core export opens with both paragraphs at 1.2. Source XML, observed outputs and reproducible generation/measurement instructions are committed in [the native spacing evidence](../e2e/pptx/native-line-spacing/README.md). This is not a pixel/video measurement or a new native claim for exact spacing, before/after spacing, indentation, inline selections or unparsed inheritance.

The pure five-field paragraph-own/body geometry cascade moves verbatim into PowerPoint core with exact provenance, shipping entries and a UI compatibility re-export. DOM assembly, CSS, selection and history stay in shared UI and the bindings; the existing fallback policy is unchanged. Twenty-eight focused core tests and 106 shared UI tests cover marker and empty-paragraph carriers, soft breaks, multiple runs, absent inheritance, immutable snapshots, unknown paragraph extension XML, save/reopen and sequential same-handler save after restoration. Ten five-demo browser checks cover fixed-line rendering and actual ribbon edit/export/document Undo/Redo with original and edited paragraph XML. Keyboard history behavior is unchanged; the ribbon workflow uses the visible document Undo control because Save can return focus to a select. Core/UI/all-five types, shipping builds and both ESM/CJS imports pass.

### Word ordinary table-cell paragraph keeping

Print Layout now honors Keep lines together in the measured ordinary automatic, top-aligned two-cell row with direct note-free stories, one body column, exact positive line spacing, zero paragraph spacing, disabled widow control and complete consistent solid borders. Either cell's first kept paragraph can defer the whole starting row. Later unstarted kept paragraphs defer their own cell and following content while the other cell advances. An oversized paragraph defers once when appropriate, then splits at a fresh origin; consumed suffixes continue without repeated keeping. Existing PageCursor startup, paragraph boxes/source tokens, paragraph cuts and row reflow are reused, without another planner or font/line-count rule.

Eighteen paired native Word compatibility-mode 15 cases cover forty-five pages, mirrored cells, occupied-page startup, oversized paragraphs and consumed prefixes. Exact page/body sets match, with physical frames within 0.2 PDF points, the retained first-ink bound below 0.15 points and a separate all-line ink bound below 0.2 points. Twenty-four additional native border/margin controls measure minimum minus 1 point, equality and minimum plus 1 point, confirming the existing startup reservations without changing keeping logic. Both native records reproduce exactly on the publication dependencies; thirty-six core exports pass seventy-two owned native open/save/reopen checks. See the [keeping corpus](../src/core/docx/__fixtures__/table-keep-lines/README.md) and [compact fit-boundary record](../src/core/docx/__fixtures__/table-keep-boundaries/README.md). These are source layout measurements, not glyph or raster parity.

Source merge provenance now prevents horizontally or vertically merged inner/enclosing cells from entering the existing measured nested-fragment profile. Explicit zero row heights are retained by the parser and refused by both nested guards and the keeping admission. Real XML regressions preserve parsed models, every no-op package entry, row-property XML and complete nested XML during unrelated direct text edits. Keeping-corpus edited paragraph comparison separately accounts for existing XML declaration newline normalization, a redundant namespace, stale lastRenderedPageBreak removal and the resulting derived structural diagnostic. It does not claim byte-exact edited XML preservation. All published two/three/four/eight-row fixture bytes and the three repaired note browser specs remain unchanged.

Validation includes 1,649 private Word core tests, 185 focused publication core checks, 79 shared DOM checks, forty-one related browser workflows and eight repeated long-table workflows after the separate inherited-spec synchronization repair. That sole inherited browser-spec exception, nested-multi-row.spec.ts, now measures page markers, decorative frame, ink and borders in one connected visible DOM snapshot while retaining exact page/body/source/history/reopen assertions and every numeric tolerance. It fixes reproduced detachment between awaited measurements during scheduled print replacement; product scheduling is unchanged. Strict core/UI and affected browser-spec types, shipping core/UI and source demo builds, lint and formatting pass.

Admission remains bounded to the ordinary single-row profile above. Nested keeping, notes or external note reservations, multiple body columns, merges, explicit heights, headers, keep-next, page breaks, floats, paragraph frames and unmeasured spacing/widow combinations retain existing behavior and explicit approximation diagnostics. No broader table keeping or native rendering claim is made.

### PowerPoint collaboration geometry oracle and raster diagnostics

Host-owned collaboration snapshot checks now read the full-precision synced shape model through a read-only protocol peer, assert the saved position and unchanged size EMUs, then compare reopened geometry with the parser's exact pixel contract. They retain strict live-peer, read-only and source/save/reopen assertions. A dedicated CSS half-pixel boundary regression proves why rounding serialized computed CSS can disagree with the model, and confirms that a real additional one-pixel save error still fails. This repairs the test oracle, not the model or writer.

Raster-fidelity tests now surface capture page errors with their original cause and remove the owned diagnostic listener on success or failure. The fallback page and its resources are closed in finally. Existing PNG/pixel assertions, timeouts, renderer and encoding behavior are unchanged. The measured synchronous PNG encoding delay remains a separate unresolved performance issue; this test-only batch does not claim to repair it or expand native rendering parity.

Fresh publication validation includes fifteen geometry/read-only workflows across React, Vue, Angular, Vanilla and Svelte, twenty regular Vue collaboration/raster cases and eleven Vue headless host cases. One existing custom-shell-only case remains skipped in the regular Vue project and passes in the headless project. Affected strict browser-spec types, framework neutrality, lint and formatting pass, with fresh core/UI and required demo prerequisites built from the preserved source. All inherited Word/native fixtures, release metadata, PowerPoint product files and prior ledger bytes are preserved.

### Visio bounded non-fit printer tile counts

The new core count result resolves only unique canonical No Style0 static print defaults with measured page/document style references0, plus explicit page overrides. Direct DocumentSheet print cells are retained but do not supply effective page defaults. Physical extents reuse the existing drawing-scale projection, and paper orientation and usable-area calculations reuse existing helpers. Inherited count metadata is isolated from legacy display settings, approximate page-break origins and editing permissions. No writer, UI print enablement, clip or snapshot tiling behavior changes.

Owned Visio 16.0.20430.20146 source evidence covers 66 admitted explicit non-fit equal-scale Letter/A4 orientation profiles, five near-machine-precision cases conservatively refused without rounding, and eight other excluded fit, centering or unequal-scale cases. Seven independent precedence controls retain ineffective DocumentSheet overrides and an effective page override. Native PrintTileCount observations and raw VSDX hashes are retained; PDF exports each cover one full physical drawing page and do not establish printer tiling. Native reopen acceptance is zero after an automation hang. Counts do not establish tile positions, raster output or other printer-driver behavior.

PagesX/Y apply only in fit mode and remain optional inactive metadata for non-fit arithmetic. Explicit non-positive, fractional, malformed or non-finite targets are conservatively refused without fallback or source repair, rather than claiming native coercion. Custom/chained styles, duplicates, dynamic or invalid local cells, ambiguous sources and count-budget overflow are also refused. The native corpus regressions verify exact no-op packages, unchanged approximate page-break arrays and unrelated rename changes limited to existing page names and app.xml TitlesOfParts, with all other entries and print-default XML preserved. See the [count corpus](../src/core/visio/__fixtures__/print-count-native.README.md) and [independent precedence controls](../src/core/visio/__fixtures__/print-count-inheritance-native.README.md).

Fresh preparation on the preserved published dependencies passes 4,061 Visio core tests with 200 existing intentional skips, including all 86 native source/preservation cases and the explicit metadata refusal regressions. Eighty-three affected page-setup, viewer-command and print-snapshot DOM checks pass. Strict types, shipping builds, lint, formatting and exact inherited-tree/frozen-source checks pass. This is count-only native source evidence, not native reopen or printer-render parity.

### Excel slicers, timelines and number groups (round three)

Table slicers (`x15:tableSlicerCache`) and PivotTable slicers (tabular `slicerCacheDefinition`) are read, drawn, edited and saved. The reader models slicers, caches, timelines and timeline caches in `workbook.slicers`; a slicer part, cache or frame the model cannot edit (OLAP or hierarchy sources, several connected PivotTables, a table with custom filter criteria, a PivotTable the admission refuses) stays verbatim and its frame is drawn as a labelled placeholder. Untouched slicer parts save byte for byte; an edit rewrites only the slicer parts of the workbook (slicer part, cache, sheet and workbook extension lists, `Slicer_*` defined names), patching the cache XML so unmodelled markup survives.

A table slicer sets a value-list `TableColumn.filter` (written into the table's `autoFilter`) and hides the rows, keeping manual hiding. A PivotTable slicer sets hidden items (`item h="1"`, `PivotTable.itemFilters`) and refreshes the report. Pure core logic covers item states (selected, unselected, no data, with sort order, hide-no-data and cross-filter modes), Excel's click selection (plain, Ctrl, Shift, empty means all), button geometry and the 14 built-in styles resolved from the workbook theme. Timelines read date-only cache fields (date items, number format 14) on no axis, accept the `dateBetween` filter a timeline writes, compute periods for years, quarters, months and days between the stored bounds and write the range into the timeline cache, the PivotTable `filters` element and the report. While a date range or item filter applies, the PivotTable keeps every item in its axis field lists, as Excel writes.

Numeric row and column fields can be grouped by integer ranges (starting at, ending at, by): the cache stores the range and the group items including the outer `<start` and `>end` items, the report shows the buckets. Date grouping (extra cache fields), discrete groups, more than two row fields and custom subtotals were not attempted and remain preserved unsupported.

The UI draws slicers and timelines over the grid, adds Insert > Slicer and Insert > Timeline field pickers, contextual Slicer and Timeline tabs (settings, 14 style gallery, columns, button height, size, time level, delete) and Group Field and Ungroup in the PivotTable menu; every string is translated in all five locales. Native Excel 16 accepts inserted, filtered, edited and grouped files with the same slicer item states and PivotTable text (seven cases in `scripts/xlsx-slicer-acceptance.ps1`; Excel-authored references come from `generate-excel-slicers.ps1` and `generate-excel-pivot-groups.ps1`). Two Excel requirements found this way are encoded: the slicer `rowHeight` attribute is mandatory, and the PivotTable slicer cache list must precede the table list in the workbook extension list. Slicer visuals follow the look of Excel's styles and are not measured against native renders; button and header metrics are approximations. Validation: 7,297 core Excel tests, 423 shared Excel UI tests, four slicer and timeline browser workflows plus a grouping workflow, and the existing pivot, ribbon and dialog browser specs.

PivotTable captions follow the language of the file (round 6). Excel takes the row and column header, grand total, Values, subtotal, (All) and (blank) words from its UI language, so a workbook saved by a localized Excel holds them in that language. The reader takes `rowHeaderCaption`, `colHeaderCaption`, `grandTotalCaption` and `dataCaption` from the definition, `item/@n` display names, the data field names, and reads the words only the cells record (`East Ergebnis`, `Gesamt Summe von X`, `(leer)`, custom subtotal function names) by building the report with a marker in place of each caption and reading the cached cell the marker lands on. `PivotTable.captions` and `PivotTable.itemNames` carry them through refresh and save; a new report uses the editor locale (de, es, fr, zh-CN tables in `pivot/captions-locales.ts`). Only the English Excel was at hand and it has no language pack, so the real-Excel fixtures set what COM can set (headers, grand total, Values, data field and item captions) and the words only a UI language supplies were relabelled in the shared strings; the localized words themselves follow the wording of the localized Excel releases and are not measured. An English Excel refreshes those words in English, so the relabelled fixtures are not in the Excel acceptance run.
