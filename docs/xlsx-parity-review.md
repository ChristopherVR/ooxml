# Excel parity review and implementation order

For the current implementation, evidence and remaining gaps, see
[XLSX parity status](xlsx-parity-status.md). **Full Excel parity is not achieved.**
The sections below retain the history of individual implementation increments;
their test counts and limitations apply to the revisions described.

Reviewed 7 October 2026 against the consolidated source at `b9af841c0`, before the changes
described below. This is a source and regression-test review, not an assertion of Excel parity.
The viewer's [features and limitations](../viewers/xlsx/docs/features.md) describe the existing
product surface. A function count or ribbon button count does not establish compatible behavior.

## Current architecture

- `src/core/xlsx` owns workbook data, formulas, editing, undo, layout, reading and writing.
- `src/ui/src/xlsx` owns one `<xlsx-editor>` element, its grid, ribbon, dialogs and input handling.
- `viewers/xlsx/packages` contains six thin framework adapters over that element and a core re-export.
- XML and OPC packaging, document properties, package encryption, signature detection and SmartArt
  already use shared core areas. Shared Office controls belong in `ooxml-ui`.
- `src/core/xlsx/formula/excel-parity.test.ts` compares calculations with recorded real-Excel results.
  The existing Windows Excel acceptance script checks saved packages for repair prompts.

## Findings and priorities

| Priority       | Area                    | Evidence and gap                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      | Next acceptance target                                                                                                                                                                 |
| -------------- | ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| P0             | Everyday editing        | `edit/clipboard.ts` clipped every copied rectangle to used cells; `grid/grid-commands.ts` routed Fill Down/Right through series inference. Fixed in this slice.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       | Trailing blanks overwrite the full finite destination; date and weekday fills repeat content, retain formatting and translate formulas.                                                |
| P0             | Save fidelity           | The reader/writer retain unsupported XML and parts, and patch some carried references. This does not prove preservation across every edit. Default clipboard payloads also do not cover all Excel metadata such as conditional formats and drawings.                                                                                                                                                                                                                                                                                                                                                                                                  | Excel-produced fixtures for metadata, structural edits, tables, drawings, external links and unsupported parts; compare package contents and reopen in Excel without repair.           |
| P1             | Clipboard and selection | Compatible selection paste now repeats copied blocks, with size checks and one undo step. Whole-row/column copies are intentionally bounded to used content. Clipboard payloads are dense arrays.                                                                                                                                                                                                                                                                                                                                                                                                                                                     | Skip blanks, full metadata semantics and bounded large-selection performance.                                                                                                          |
| P1             | Formula compatibility   | A substantial function catalogue and recorded Excel corpus exist; missing functions and cached external references remain documented.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | Expand real-Excel differential cases for coercion, errors, arrays, names, tables, dates and structural edits before treating any function as complete.                                 |
| P1             | Grid and print fidelity | The core supplies grid metrics and cell views; printing uses a browser-rendered range. Page Layout and page-break preview are absent.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | Reference workbooks and screenshots for widths, fonts, wrapping, merges, panes and number formats; paginated print with repeating titles and scaling.                                  |
| P2             | Charts and drawings     | Live chart rendering and limited chart edits exist; many drawing types remain placeholders. SmartArt primarily uses cached drawings.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  | Shared chart/DrawingML support for axes, labels, series formatting and geometry, with read/edit/save fixtures.                                                                         |
| P2             | Pivot/data features     | Pivot caches/tables are retained without pivot refresh/rendering; slicers and timelines are not drawn. Sparklines are drawn (line, column, win/loss) but cannot be authored. Connections are cached.                                                                                                                                                                                                                                                                                                                                                                                                                                                  | Implement bounded pivot aggregation and views, then filtering controls and refresh, with explicit unsupported-feature reporting.                                                       |
| P2             | Collaboration           | `ooxml-core/xlsx/collab` maps the workbook to Yjs and binds an edit session per edit: cell values, formulas with cached values, formats (content-addressed), row heights, column widths, merges, sheet add/rename/move/delete/state/tab colour, defined names and selections are shared; local-only undo, reconnect and conflict cases are tested. Not shared: charts, drawings, images, comments, hyperlinks, conditional formats, validations, tables, filters, pivots, freeze panes and views, page setup, protection, document properties. Row/column insert and delete travel as the resulting cell moves. The `<xlsx-editor>` is not wired yet. | Wire the editor and bindings (Ctrl+Z through the binding, remote selections on the grid); share the remaining sheet parts, starting with comments, hyperlinks and conditional formats. |
| Separate scope | Excel platform features | VBA is carried but never executed; add-ins and Power Query are not implemented. XLSB/ODS are unsupported.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             | Define platform and file-format requirements explicitly before claiming whole-product 1:1 parity.                                                                                      |

## Reuse decisions

| Concern                | Reuse and ownership                                                                                                                                                                     |
| ---------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Reading/writing        | Shared `xml` and `opc`; SpreadsheetML semantics stay in `xlsx`. Do not introduce another parser/model in a binding.                                                                     |
| Formatting and drawing | Shared `units`, `color`, `geometry`, `diagram`, `chart` and the planned DrawingML extraction. Move format-neutral logic out of PPTX with provenance and strict types before reusing it. |
| Editing                | Existing XLSX edit session, formula translation, history and change notifications. Fix behaviors once so all adapters inherit them.                                                     |
| Ribbon/dialog controls | Existing `office-ui-*` controls, dialog shell, icons and theme contracts. Product-specific commands and labels stay with XLSX.                                                          |
| Collaboration          | `collab` provider/session infrastructure; implement only the spreadsheet mapping in XLSX. Coordinate shared lifecycle work with the other product sessions.                             |
| Legacy/encryption      | Keep binary XLS/CFB codecs in `ole2` and modern OOXML encryption in shared `crypto`.                                                                                                    |

## Implemented slices

Finite copies now retain their entire rectangle, including trailing blank rows and columns, in
internal payloads, TSV and HTML. Whole-row and whole-column copies retain the existing used-area
optimization only on the full-sheet axis. This intentionally does not implement complete whole-row
or whole-column Excel clipboard semantics.

`EditSession.fill` accepts an optional `FillMode`: `auto` retains series inference and `copy`
repeats source cells with translated formulas. Fill Down and Fill Right request `copy`; the fill
handle continues to request the existing default. Both modes reuse the same history and calculation
pipeline. This is an additive API change with no new package or framework-specific logic.

The viewer's opt-in local-source aliases now resolve the directory entry of `ooxml-ui/xlsx`
and nested core entries such as `ooxml-core/xlsx/ui`. The old file mapping prevented the demo
from building against the consolidated editor source, so browser validation could not exercise
unreleased shared changes.

Regression coverage includes clipboard geometry and representations, empty selections, paste
modes, transpose, cut references to blank cells, undo/redo, save/reopen, date/text fills, mixed
formula references, read-only command handling and browser keyboard behavior.

Microsoft's [paste options](https://support.microsoft.com/en-au/excel/paste-options) document
blank-cell replacement unless Skip blanks is selected. Its
[keyboard shortcuts](https://support.microsoft.com/en-gb/accessibility/excel/keyboard-shortcuts-in-excel)
define Fill Down/Right as copying contents and formatting. These references specify the intended
behavior; the new tests are not a substitute for new recorded real-Excel differential fixtures.

### Selection paste and cut completion

`EditSession.paste` now accepts either a top-left address or a destination range. A copied block
repeats across compatible finite selections, with translated formulas and tiled merges. Values-only
paste retains destination formatting. The operation captures one destination snapshot, recalculates
once and makes one undo step. Incompatible sizes fail before edits. Repeated selections above
250,000 cells are explicitly rejected; full-sheet axes retain the existing bounded behavior.

The shared editor passes the selection to this core operation and reports failures. It now captures
cut state before the core consumes the move flag, so a successful cut/paste clears the marquee.
Multi-area paste is explicitly rejected until its semantics are supported. Tests cover formulas,
blanks, transposition, merged cells, formatting, undo/redo, UI errors and native browser copy/paste.
Complete clipboard metadata support remains outstanding.

### Combined Paste Special options

The core paste engine accepts independent content mode, transpose and Skip Blanks options,
while retaining the existing string-mode API. Ribbon paste variants and the grid use the same
clipboard path; the dialog gathers options and its requesting command applies them once.
Ctrl+Alt+V now opens Paste Special through the shared shortcut map.

An intentional fixture recorded with Microsoft 365 Excel 16.0 build 20430 covers all 16
combinations of All/Values/Formulas/Formats, transpose and Skip Blanks. A true empty cell,
including one with formatting, leaves destination content and formatting untouched when
skipped. A formula returning an empty string, zero and FALSE still paste. Cut with Skip Blanks
is rejected before edits, matching a separate native check. Empty string constants produced
by Paste Values now survive cell pruning and save/reload. Tests cover tiled selections and
one-step undo/redo. Column widths are implemented in the later slice below; full
clipboard metadata remains incomplete.

Reproduce the fixture with `scripts/record-xlsx-paste-options.ps1 -OutputFile <temporary-json-path>`
on Windows with Excel installed. The recorder creates its own hidden Excel instance and
closes its workbook without attaching to an existing user session.

Validation before integration with concurrent main changes: all 3,801 core XLSX tests,
301 shared XLSX UI tests, 47 viewer/binding tests and 52 browser tests passed. Core/UI/viewer
typechecks and core/UI/viewer package builds passed. The browser test exercises native copy,
Ctrl+Alt+V, combined options and one-step keyboard undo; the command returns focus to the grid.
Published-import guards and clean-consumer package smoke checks passed for all seven XLSX
packages, including `.xlsx` and legacy `.xls` loading, declarations and framework bindings.

### Ribbon directional fill

Ribbon Fill Down/Right/Up/Left now request the shared core copy mode, matching the keyboard
commands instead of inferring a weekday or numeric series. Four-direction regressions verify
weekday repetition, relative and absolute formula references, source formatting and undo/redo.
AutoFill and the explicit Series command retain their separate series behavior.

### Arithmetic Paste Special

Add, Subtract, Multiply and Divide are available through the existing shared dialog and core
paste options. The arithmetic adapter reuses the formula engine's scalar operators and numeric
text parser. Formulas are combined with translated copied references and existing destination
formulas; Values uses copied results while preserving destination formulas. Numeric text
participates, while nonnumeric constants and boolean constants stay unchanged. Boolean formula
results participate in Values arithmetic. Formatting follows the selected content mode.

An intentional Microsoft 365 Excel 16.0 build 20430 corpus records 512 combinations across four
operations, four paste content modes, constants, numeric text, blank cells, formulas, booleans
and errors, including formatting. Reproduce it with
`scripts/record-xlsx-paste-arithmetic.ps1 -OutputFile <temporary-json-path>`. Every native result
and formula passes a direct comparison with undo/redo. Separate regressions cover tiled and
transposed operations, Skip Blanks, relative references, save/reload and atomic cut rejection.
A live native check also confirmed Excel rejects arithmetic Paste Special for cut cells.

The native clipboard browser regression uses Values with Multiply on an existing formula,
checks the combined formula and result, and verifies one-step keyboard undo. All 898 core
editing tests and the three clipboard browser regressions passed. These checks establish this
tested subset, not complete clipboard or whole-workbook parity.

The full core XLSX suite passed 4,839 tests across 127 files, the shared XLSX UI suite passed
307 tests, and core/UI strict typechecks and builds passed. Initial machine resource pressure
was handled by running UI tests serially; broad core validation ran after memory recovered.
All 54 browser tests and 47 viewer/binding tests passed, along with viewer typechecks,
seven package builds, published-import guards and clean-consumer package smoke checks.

Primary reference: [Microsoft Paste options](https://support.microsoft.com/en-au/excel/paste-options).

### All Except Borders and transpose references

The `noBorders` paste mode copies content and supported formatting while retaining the complete
destination border model, including edge colors and diagonal directions. A truly empty copied
cell clears destination content and other formatting while keeping its borders. Skip Blanks
preserves the entire destination cell. The shared dialog enables All Except Borders; column
width paste and complete clipboard metadata support remain outstanding.

Twenty native Excel 16.0 build 20430 cases compare four destination cells across transpose,
Skip Blanks and all five arithmetic-operation choices. They verify values, formulas, fonts,
fills, number formats and six border directions with undo/redo. Reproduce them with
`scripts/record-xlsx-paste-borders.ps1 -OutputFile <temporary-json-path>` in a fresh hidden
Excel instance. Additional regressions cover truly empty source cells, tiled merges and border
save/reload. The browser test uses native copy, the shared dialog and keyboard undo.

The corpus also exposed incorrect reference translation during transpose. Excel applies the
copied block's top-left offset to every formula while transposing cell positions. The core now
uses that offset for each tile. Live native checks confirmed the same reference behavior for
All, Formulas and All Except Borders. The earlier transposed-arithmetic regression had asserted
the implementation's incorrect references; its expectations now follow this native evidence.

Validation: the full core XLSX suite passed 4,865 tests, the shared XLSX UI suite passed 308,
the viewer/binding suite passed 47 and all 55 browser tests passed. Core/UI/viewer typechecks,
core/UI builds, seven package builds, published-import guards and clean-consumer package smoke
checks passed. A subsequent native
check confirmed All Except Borders rejects cut cells; the atomic rejection regression and
the full 26-test border suite passed after adding that guard.

### Column Widths paste

The shared dialog now enables Column widths. Clipboard snapshots carry visible column metrics and
the original row extent, independently of the used-area optimization for copied cells. Core paste
reuses column-span editing and history, changes only destination dimensions, and preserves cell
contents, formatting, merges and column outline properties. Hidden source columns paste as zero
width; unhiding their destination reveals the receiving sheet's default width. Transpose repeats
the original column-width pattern, while incompatible selections fall back to one block. Arithmetic
and Skip blanks do not change width-only paste behavior.

`scripts/record-xlsx-paste-widths.ps1` reproduces 80 cases from Microsoft 365 Excel 16.0 build 20430:
finite and whole-column copies, visible/hidden/default widths, transpose, compatible/incompatible
selections and sheet limits. Tests compare stored width and visibility, unchanged cell contents,
one-step undo/redo, cross-workbook clipboard snapshots and save/reload. The browser test copies
through the system clipboard, applies the shared dialog, and checks undo/redo. Text-only clipboard
sources lack column metrics and are rejected explicitly; complete external Excel clipboard format
interoperability remains outstanding.

### Notes and validation clipboard semantics

Clipboard snapshots now include notes (authors, text and modeled replies) and clipped validation
rules. All and All except borders paste these annotations; the shared Paste Special dialog also
offers Comments and Validation separately. Values, Formulas and Formats retain destination notes
and rules. Skip blanks applies independently to annotation absence, so a note or validation on an
otherwise empty source cell still copies. Validation formulas move from their individual source
origins, including transposition. Range removal rebases surviving relative formulas when a rule's
original anchor disappears. Cut/paste moves annotations and records both sheets in the same undo
step, using the existing reference-rewrite engine.

`scripts/record-xlsx-paste-annotations.ps1` records 56 Microsoft 365 Excel 16.0 build 20430 cases
across seven paste modes, transpose, Skip blanks and arithmetic. Regression tests compare 224
destination cells' values, note authors/text, validation formulas and settings, then undo/redo.
Additional tests cover multi-area rule clipping, cut across sheets, tiled cross-workbook snapshots
and save/reload. Browser tests exercise native clipboard note paste/undo and validation paste with
subsequent accepted/rejected entry. [Microsoft's paste type enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlpastetype)
documents the corresponding modes. The native corpus covers legacy notes; it does not establish
threaded-comment presentation or complete clipboard metadata parity. Conditional formats,
drawings and external native clipboard interchange remain to be implemented. Hyperlink clipboard
support is implemented in the following slice.

Verification: 5,010 core XLSX tests, 310 other shared UI tests plus the corrected translation test,
47 viewer binding tests and all 58 browser tests passed. Core/UI/viewer typechecks and builds passed.
A generated library workbook was opened and saved in native Excel, then reloaded: note author/text
and the transposed validation's accepted/rejected results remained correct after Excel combined
equivalent validation ranges. This is a focused native acceptance probe, not a lossless-save claim.

### Hyperlink clipboard semantics

Clipboard snapshots now carry clipped hyperlink ranges, external targets, in-workbook locations
and tooltips. All and All except borders replace destination links; the other supported modes
retain them. Skip blanks retains destination links wherever the source has no link, independently
of whether source cell content is present. Range splitting preserves portions outside the paste
area, tiled paste shares one undo step, and cut uses the existing reference-rewrite engine for
internal targets. HTML copy emits escaped anchors; HTML import reads their targets and tooltips.
Both paths reuse the shared OPC hyperlink policy. Plain TSV cannot represent link metadata.

`scripts/record-xlsx-paste-hyperlinks.ps1` records 64 Microsoft 365 Excel 16.0 build 20430 cases
covering eight modes, transpose, Skip blanks and arithmetic. Tests compare 256 destination cells'
values and link targets/locations/tooltips, plus undo/redo. Additional regressions cover clipped
spans, cross-workbook snapshots, tiling, save/reload, HTML interchange, cut references and shared
link policy. The browser test inserts links through the existing dialog, copies through the system
clipboard, and verifies target replacement and undo/redo. Conditional formats, drawings and complete
native Excel clipboard interchange remain outstanding.

Verification: 5,078 core XLSX tests, all 311 shared XLSX UI tests, 47 viewer binding tests and all
59 browser tests passed. Core/UI/viewer typechecks and package builds passed. Excel opened and
saved a generated workbook containing transposed copied links; library reload retained both target
kinds and the tooltip. This focused probe does not establish complete hyperlink or export parity.

## Conditional formatting clipboard follow-up

Internal copy snapshots now carry modeled conditional rules in the shared XLSX core. All, Formats
and All except borders replace destination rules, while All merging conditional formats retains
them. Skip blanks replaces rules only where source rules apply, independently of blank cell values;
native Excel applies this replacement behavior even when the merging option is selected. Other
paste modes retain existing rules. Copied rules take the highest priorities, relative formulas
follow their actual rule anchors, and clipped or split surviving ranges keep their semantics.
Repeated paste creates one rule spanning the tiles, retaining shared color-scale and ranking
statistics. Cut removes source coverage and uses the existing workbook reference rewrite. Clearing
rules also rebases formulas when the original anchor disappears. Undo records all affected rules.
Color-scale interpolation now truncates each channel delta toward the starting color, matching
native Excel; the shared color helper's default rounding for other rendering stays unchanged.

`scripts/record-xlsx-paste-cf.ps1` reproduces 216 Microsoft 365 Excel 16.0 build 20430 cases across
nine paste modes, transpose, Skip blanks, arithmetic, partial source coverage and empty sources.
Tests compare every sheet rule's ranges, formulas, priorities, fill, bold and Stop If True, plus
1,512 cells' values, displayed fills and applicable priorities. Additional regressions cover tiling,
cross-sheet cut, clearing, cross-workbook snapshots, formula-valued visual thresholds and save/reload.
The shared dialog exposes the merging option and reuses the existing paste command.

This covers modeled base conditional formats. Advanced x14 data-bar attributes and unsupported
extension rules are not copied; new base data-bar rules deliberately do not reuse an original
extension ID. Complete native clipboard interchange, drawings and full Excel parity remain open.
Excel opened and saved a generated workbook with a repeated conditional-rule paste; the farthest
destination cell displayed the expected red fill, and library reload retained the rule and priority.
This is focused evidence, not a complete export-fidelity claim.

Verification: 5,299 core XLSX tests, 312 shared XLSX UI tests, 47 binding tests and all 60 browser
tests passed. Core/UI/viewer typechecks, builds, published-import guards and clean-consumer package
smoke checks passed. The color-scale regression compares eight native colors from a repeated Formats
paste, proving that destination color-scale statistics span the entire pasted selection.

## Advanced data-bar preservation follow-up

Linked x14 data-bar rule XML now belongs to its base rule in the XLSX model, rather than the
worksheet's unrelated preserved XML. Copy snapshots retain gradient/solid settings, direction,
axis configuration, positive and negative borders/colors, automatic limits and unknown rule fields.
Copies receive independent Office GUIDs. Range clipping, splitting, tiling, clearing and cut now
regenerate extension coverage from the modeled rule, avoiding stale or orphan records. The existing
formula traversal also visits extension threshold formulas. Numeric/formula threshold edits update
both representations when saving; untouched automatic limits retain their extension types.

This uses the shared XML model and serializer. PowerPoint's GUID generator and its four regression
tests moved into `crypto/uuid`; PowerPoint retains its existing API and XLSX reuses the helper,
including older-runtime fallbacks. `PROVENANCE.md` records the extraction.

`scripts/record-xlsx-databar-clipboard.ps1` reproduces 25 Excel 16.0 build 20430 cases as before/after
worksheet XML. Tests compare complete linked rule settings across All/Formats, transpose, clipped
sources, three appearance variants and automatic limits, then save/reload and undo/redo. Additional regressions cover
unknown fields, unrelated extensions, clearing, structural formula rewrites and cross-workbook/cut
copies. Excel accepted and saved three library-authored copied-rule workbooks, retaining all
applicable measured properties. It also accepted changed numeric limits in a separate edit probe.

These settings are preserved, but advanced data-bar rendering and their dedicated editor controls
remain incomplete. The opaque rule record is not yet a complete typed appearance model. Other x14
rule types and native clipboard interchange remain open. The extension's available fields are
documented in [Microsoft's x14 data-bar reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.databar?view=openxml-3.0.1).

Verification: the full core XLSX suite passed 5,327 tests; all 29 data-bar regressions passed with
the additional automatic-limit case. The 12 shared GUID/PowerPoint identity tests, 312 XLSX UI tests,
47 binding tests and 12 clipboard/file browser tests passed. Core (strict and PowerPoint), UI and
viewer typechecks, core builds, published-import guards and packed-consumer smoke checks passed.

## Data-bar appearance rendering follow-up

The shared conditional-format evaluator now resolves linked x14 fill settings,
explicit direction, positive/negative colors, enabled borders and their separate
colors, and value visibility. It reuses the shared XML parser, SpreadsheetML color
reader and theme/tint resolver, caching decoded appearance per rule. The DOM painter
consumes those resolved settings: solid bars stay solid, explicit RTL bars anchor
at the right, negative fills no longer always become red, and disabled borders
remain absent. Legacy bars inherit their base fill for negative values.

Three appearance variants from the existing native Excel fixture are checked with
the real calculation engine. A grid regression loads the native worksheet and
checks appearance changes repaint correctly. A browser regression imports the
native sheet XML into a workbook and checks actual CSS, copy/paste and undo/redo.

This establishes the tested appearance settings, not pixel parity. Axis placement,
negative/positive bar lengths, minimum/maximum lengths and automatic thresholds
still use incomplete geometry. Gradient endpoint intensity and outline dimensions
need native visual comparisons. Context direction currently falls back to the
sheet's RTL flag and still needs native comparisons against reading order and
locale. Dedicated appearance editor controls and a complete typed model remain open.

Verification: 5,331 core XLSX tests, 313 shared XLSX UI tests and 47 binding tests
passed. The new data-bar browser regression and the 12 clipboard/file browser
checks passed. Core/UI/viewer typechecks, strict core ESM/CJS/declaration builds,
the UI build and all seven viewer package builds passed.

## Data-bar axis and length follow-up

The core now calculates normalized bar starts, lengths, growth direction and axis
positions. Automatic axes separate positive and negative bars at zero; middle axes
allocate half the cell width to each side. Explicit RTL mirrors that geometry.
Automatic minimum/maximum limits include zero for one-sided ranges. Equal limits
produce half-length bars, while all-zero axis ranges show the axis without a bar.
The grid consumes these coordinates and paints the rule's dashed axis color.

`scripts/record-xlsx-databar-geometry.ps1` uses a fresh hidden Excel application to
record 72 worksheet/PDF cases, spanning three axis modes, both directions,
automatic versus numeric limits, mixed/one-sided values and constant ranges.
`scripts/extract-xlsx-databar-geometry.py` uses PyMuPDF to extract native solid-fill
vectors, normalized to a full-length native reference bar. Committed JSON retains
the source worksheet XML and measured coordinates, not generated workbooks/PDFs.
Regression comparisons allow 0.015 for axis gaps and print quantization. All 72
native comparisons and the legacy-length save/reload regression passed.

Legacy base rules now retain explicit minimum/maximum length percentages and use
their specified interpolation (defaults 10/90). That formula is documented in
[Microsoft's base data-bar reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.databar?view=openxml-3.0.1).
Automatic limits and axis options use the preserved extension and shared XML/color
readers rather than introducing a second OOXML model.

This covers the measured normalized geometry. Advanced non-default length
percentages, thresholds outside the measured cases, contextual reading direction,
gradient endpoint intensity and pixel-level insets/dashes still need native
comparisons. Dedicated editor controls remain incomplete.

Verification: 5,404 core XLSX tests, 313 shared UI tests, 47 binding tests and
13 clipboard/file/data-bar browser checks passed. Core/UI/viewer typechecks,
core ESM/CJS/declaration builds, the UI build and all seven binding package builds
passed. Excel also opened and saved a library-authored legacy bar with explicit
20/80 lengths, reporting those same percentages through its native object model.

## Advanced data-bar percentages and context follow-up

Native measurements now cover seven additional percentage pairs: 20/80, 10/90,
40/40, 0/0, 5/55, 40/80 and 100/100. These 504 cases exposed and fixed minimum
lengths around zero, percentage-dependent automatic axis positions, one-sided
negative axis placement and capped middle-axis lengths. A nonzero minimum gives
zero a half-minimum-length bar. Equal percentage bounds and zero/full lengths
are included. The existing shared geometry helper applies the resulting rules.

Another 108 cases record contextual direction on both sheet directions and all
three cell reading orders. They retain native worksheet and style XML as well as
PDF vectors. Context follows the sheet's RTL setting even when a cell's reading
order differs, confirming the existing fallback. That measured behavior agrees
with [Microsoft's spreadsheet rendering protocol](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-exspxml3/a32e609d-a33b-460a-b2db-b1153730699b).
Together with the earlier 72 cases, 684 native geometry comparisons pass.

Logical length percentages now come from the linked x14 rule on read. One shared
resolver supplies layout and the legacy fallback writer. Typed edits update the
extension when saving; full-width 0/100 bars retain Excel's required legacy 10/90
fallback. Editing, clipboard paste, undo/redo and save/reload tests cover those
representations. Excel accepted four library-authored edits (5/55, 0/100, 40/40
and 0/0), saved them, and the library reloaded the same logical percentages.

Reproduce each length profile with `scripts/record-xlsx-databar-geometry.ps1`
using `-PercentMin`, `-PercentMax` and a separate `-OutputFolder`. Extract each
folder with `scripts/extract-xlsx-databar-geometry.py <folder> <output-json>`;
use `--append` after the first profile to merge disjoint cases from the same Excel
build. Context profiles use `-Context -Kinds mixed,positive,negative` and
`-ReadingOrder -5002`, `-5003` or `-5004`. Native PDFs/workbooks stay in temporary
folders. The committed fixtures retain measurements and the source XML.

Verification: 6,020 core XLSX tests, 313 shared UI tests, 47 binding tests and
13 focused browser checks passed, along with core/UI/viewer typechecks and core
ESM/CJS/declaration builds. The geometry comparison tolerance remains 0.015 for
axis gaps and print quantization. Other threshold configurations, gradient
endpoint intensity, pixel-level insets/dashes and dedicated editor controls
still need work; these checks do not establish whole-workbook Excel parity.

## Shared chart UI and SmartArt review

The XLSX Insert Chart / Change Chart Type picker now uses `office-ui-gallery`,
the same component behind PowerPoint's `pptx-ui-ribbon-gallery`, with labelled
visual tiles and keyboard selection. A new panel mode reuses its tile rendering,
selection contract, theme styles and safe SVG parser. Chart data and SVG generation
remain in the strict XLSX core. No PowerPoint engine or model was copied into XLSX.

Existing column/bar charts can now change between clustered, stacked and 100%
stacked through the dialog. Previously the grouping control was hidden while
editing, and submission only changed the chart family. Undo restores the original
chart and redo/save/reopen retain the edited grouping, title and series.

`scripts/record-xlsx-chart-types.ps1` records six transitions in an isolated hidden
Excel instance. The committed Excel 16.0 build 20430 fixture retains native chart
XML, series formulas, names and values. Six regression tests check reading those
parts and all 36 source-to-target edits through save/reload and undo. Native Excel
also opened and saved six library-authored variants; the library reloaded each with
the expected family, grouping, title and values. Regenerating a changed chart type
can materialize default theme colors and does not establish full style fidelity.

Playwright MCP reviewed the chart dialog and its keyboard selection, plus its fit
inside a 390 px viewport. Browser coverage exercises insertion, grouping changes,
undo/redo and save/reopen. The existing native SmartArt fixture was inspected
through COM and Playwright MCP: Basic Block List, three nodes, labels Plan/Build/Ship,
rendered by `office-ui-smartart`. This checks those labels and cached drawing
display, not pixel equivalence or SmartArt editing.

Next shared UI targets are Chart Design's color/style/layout galleries and
SmartArt's layout/color/style galleries and text pane. Their PowerPoint operations
currently depend on its product model; extract format-neutral behavior into shared
`chart` / `diagram` core areas and shared UI, then add XLSX adapters and native
fixtures. SmartArt insertion/reflow/text editing, effects and typography fidelity,
advanced chart types/axes/labels, recommended-chart behavior and whole-workbook
visual equivalence remain unverified or unsupported. The data-bar rule editor also
needs controls for the advanced settings already preserved by core.

Verification: 314 XLSX UI tests plus ten shared gallery tests, 14 focused PowerPoint
gallery tests, 47 binding tests and six native chart tests passed. Core/UI/viewer
typechecks, core and UI builds, viewer package builds, published-import guards,
script tests and clean-consumer package smoke checks passed. The new browser test
covers keyboard selection, narrow-screen fit, grouping edits and save/reopen;
all 62 browser tests passed, including the six-framework matrix.

## Evidence required for parity

### Native SmartArt appearance follow-up

COM found that the Basic Block List fixture uses white 59 pt Aptos Narrow text;
the shared renderer previously inherited black shell text and the shell font.
`office-ui-smartart` now uses each cached shape's font-reference color and theme
Latin font, with explicit run formatting taking precedence. Its stylesheet no
longer overrides a declared typeface. XLSX supplies its workbook's major/minor
fonts through a format-neutral `schemeFonts` property.

The same renderer now calls `resolveDrawingColor` in the shared `diagram` core
for DrawingML colors and transforms, instead of resolving base colors again in
UI. Alpha becomes an SVG rgba paint. Unsupported transforms are reported in
`unappliedColorTransforms`; gradient geometry and other existing approximations
remain reported separately. No PowerPoint product engine is imported or copied.

`scripts/record-xlsx-smartart-appearance.ps1` reproduces the committed three-node
font/color measurement against the existing native workbook with its own hidden
Excel application. Playwright MCP confirmed white text, Aptos Narrow and a
computed 78.6667 px font size (59 pt at 96 dpi). Browser coverage checks those
colors and the actual computed font family. This does not establish typography
layout, mixed runs, wrapping, effects, editing/reflow or whole-diagram pixel parity.

Verification: 321 combined XLSX/SmartArt UI tests, 47 binding tests and all 62
browser tests passed. Core/UI/viewer typechecks and builds, viewer script/import
guards and package smoke checks passed. The shared UI tarball imported all 93
entry points in a clean SSR consumer and registered 45 element tags in a DOM
consumer. Generated native output remains temporary; only measurements are
committed alongside the existing workbook fixture.

Track reading, display, editing, calculation and writing separately for each feature. A retained
part does not count as rendering or editing support. Each completed feature needs representative
Excel-produced fixtures, behavioral assertions, editing and round-trip checks, and browser coverage
where user interaction is involved. Test every framework's binding contract; exercise shared editor
behavior once, supplemented by the existing framework browser matrix.

Use one pinned Excel version/platform as the first comparison target. Expand that matrix later
because desktop, web and Mac Excel do not have identical behavior. Maintain explicit exclusions
until each passes its acceptance checks.

## Validation of this slice

- Full XLSX suite: 113 files, 3,628 tests passed with one worker. The first parallel run failed
  two existing 500 ms performance limits; both passed in isolation and in the serial full run.
- Focused UI command/grid/keyboard suites: 16 tests passed. Local-source alias tests: 2 passed.
- Production demo build and the browser Fill Down/Right keyboard/undo/redo regression passed.
- Core and UI strict typechecks and `git diff --check` passed; changed code was formatted.
- A focused live Excel COM check on Microsoft 365 Excel 16.0 build 20430 confirmed weekday
  Fill Down repeats `Monday`, tiled formula paste produces `=F6+$C$1`, and trailing copied
  blanks clear destination values. Full native save acceptance and new differential fixtures
  remain outstanding.
- The selection-paste iteration passed 1,675 editing and recorded formula-corpus tests after
  rebasing onto current main, 47 viewer unit/binding tests, all 51 browser tests, core/UI/viewer
  typechecks, package builds, published-import guards and clean-consumer package smoke checks.

## AMORLINC

Added the shared financial function, automatically discoverable by every editor's
existing function catalog. Its direct implementation handles the initial prorated
period, regular depreciation, the final salvage remainder and exhausted assets.
It runs in constant time even for large requested periods. Native Excel 16.0
results establish valid bases 0/1/3/4, fractional argument behavior, equal-date
full periods, actual-day February 29 normalization, purchase-year denominators
for basis 1, and required omitted-slot errors distinct from empty cell references.
The generic function-call boundary now lets a function declare that error policy.

Primary reference: [Microsoft AMORLINC](https://support.microsoft.com/en-us/excel/functions/amorlinc-function).
Reproduce the committed 130-result fixture with
`scripts/record-xlsx-amorlinc.ps1 -OutputFile <temporary-json-path>` on Windows
with Excel installed. The recorder uses its own hidden application/workbook.
It never attaches to an existing workbook. The recorded JSON is intentional test
data; generated native workbooks and distribution bundles are not committed.

Verification: 130 native-result comparisons plus catalog discovery, incremental
recalculation, dynamic arrays and save/reload in both date systems. All 135
function tests passed. The broader XLSX suite passed 3,748 tests before the two
additional date-system round-trip tests; the strict core typecheck also passed.
These checks establish the tested function subset, not whole-workbook native parity.

## DB and DDB follow-up

Native comparisons found existing fractional-argument bugs. DDB rounded periods
up to whole periods; Excel keeps fractional exponents above period 1 and uses
first-period depreciation below period 1. DB incorrectly returned zero below
period 1 and used fractional months instead of truncating them. Both functions
iterated once per period, allowing enormous finite life/period arguments to tie
up calculation. Their declining-balance closed forms now take constant work,
preserve the salvage behavior, and retain the rounded fixed rate for DB.

The [DDB reference](https://support.microsoft.com/en-us/excel/functions/ddb-function)
and [DB reference](https://support.microsoft.com/en-us/excel/functions/db-function)
supply the depreciation formulas. `scripts/record-xlsx-depreciation.ps1` reproduces
184 results from native Excel 16.0, covering fractions, final periods, optional
arguments, coercion, errors, salvage and a trillion-period lifetime. The committed
fixture and dynamic-array/round-trip tests establish this measured subset.

The full XLSX core suite passed 3,958 tests across 116 files before adding the last
round-trip assertion; all 297 shared XLSX UI tests passed. Both core TypeScript
projects passed. A five-formula workbook containing AMORLINC, DB and DDB was saved
by the library, opened/recalculated/saved by Excel 16.0, then reloaded/recalculated
by the library with unchanged values. This was a generated native acceptance
probe, not a general save-fidelity claim. Browser tests were not rerun because
these changes are in shared calculation code and add no new UI controls.

## Shared Change Colors catalog

PowerPoint's 17 Chart Design > Change Colors palettes now live in the shared,
strict `ooxml-core/chart` area. PowerPoint imports compatibility exports and keeps
its UI descriptors and product-specific edits. The catalog reuses the shared
DrawingML color resolver instead of maintaining its own color conversion math.

`scripts/record-xlsx-chart-colors.ps1` records all 17 palettes with 1, 2, 4, 7,
10, 19 and 55 series. Run it again with `-CustomTheme` for the second fixture.
Both probes use a new hidden Excel application and close their own workbook.
Excel 16.0 build 20430 provided 238 cases and 3,332 colors. The Office baseline
up to 10 series matches exactly. Across both fixtures, 224 cases match exactly;
38 individual colors in 14 cases differ by one RGB channel step. That native
rounding gap remains open. A temporary before/after comparison confirmed that
all 3,332 shared-catalog outputs match the previous PowerPoint implementation.

The catalog extraction initially left XLSX without a Change Colors control.
The following iteration adds its authoring command and shared gallery.

Validation: 239 palette regressions, the prior broader 1,525 chart-related core
tests, 14 PowerPoint gallery/style tests, strict/legacy core and UI typechecks,
a full core build and clean package imports. Native JSON fixtures are test data
and absent from the published distribution.

## Change Colors authoring and shared gallery

Chart Design now exposes the shared Office gallery with four Colorful and
13 Monochromatic palettes. XLSX supplies theme colors, translations and its
command adapter; the gallery and SVG swatch helpers are shared with PowerPoint.
Dropdown arrows/Home/End move focus across sections, activation applies the
palette, and Escape returns focus to the trigger. The editing command supports
undo/redo, read-only restrictions, and save/reopen across all six bindings.

Core stores native chart color-style ids and DrawingML series/point choices,
including transforms. Palette edits replace matching automatic choices while
preserving custom RGB fills and marker overrides. Existing chart XML is patched
in place so axes, labels, effects and extension data survive. A changed palette
gets a private color-style part, avoiding changes to other charts sharing the
source style. Unedited chart parts remain byte-identical. Type changes also
serialize the explicit point colors now carried by the model.

`scripts/record-xlsx-chart-palette-edits.ps1` records 32 states from eight native
families: column, bar, line, area, pie, doughnut, scatter and radar, each before
and after automatic/manual palette edits. It uses a fresh hidden Excel instance
and reopens its saved copies to measure the visible marker colors correctly.
The committed fixture came from Excel 16.0 build 20430. Regression tests compare
the tested series/point colors exactly and check editing history, save/reload,
native axes/extensions, legacy color edits and shared-style isolation.

A separate temporary acceptance probe opened all 16 library-authored palette-12
workbooks in Excel, compared their series and point colors to the native
recording, saved native copies and reloaded those copies with unchanged chart
views. Excel's saved manual line/radar examples report COM ChartColor 2 rather
than 12, identically for native and library edits; their color-style metadata
and displayed colors still match. This is evidence for the measured slice.

Playwright MCP reviewed the actual ribbon, palette previews and chart rendering.
The browser regressions check all six bindings, keyboard focus, selection state,
small-viewport popup bounds, undo/redo and save/reopen. Whole-Excel UI parity
remains open: chart styles, Quick Layout, extended chart families and richer
SmartArt authoring are still outstanding, along with the previously measured
palette rounding differences. Arbitrary gradients, scheme colors indistinguishable
from automatic choices, combination-chart plots and all marker styles have not
been established by these fixtures.

Validation for this iteration: 6,046 XLSX core tests, 283 focused palette/style,
DrawingML writer and native editing tests, 326 XLSX/shared-gallery UI tests,
125 PowerPoint gallery regressions, 47 binding tests and 68 browser tests passed.
Core, shared UI and viewer typechecks, production builds and clean-consumer
package checks also passed. Generated native acceptance workbooks were temporary
and are not part of the published package or committed fixture set.

## Native chart-style model

PowerPoint's existing Chart Styles gallery is six recolor presets, so using it
unchanged would not supply native Excel styles. The native style definition
reader and resolved contract now live in the shared strict chart area;
PowerPoint keeps compatibility adapters. XLSX imports the same theme-relative
font, fill, line, reference and opaque-XML model. Existing style-part bytes
survive title edits and save/reload. Style authoring is excluded from ChartPatch
until the native picker and rendering are wired.

The native recorder `scripts/record-xlsx-chart-styles.ps1` creates each chart
independently, sets styles 201 through 216, fixes palette 10, saves and reopens
owned copies through a hidden Excel instance. Recorded style XML and COM font
sizes cover different title sizes and major/minor font references. Style 204
leaves the title size inherited and COM returns a non-positive size; the fixture
records that measurement as unavailable. It is not evidence for its displayed
title size. The 16 chart import/edit/round-trip regressions retain source style
parts exactly, while shared reader and PowerPoint compatibility tests cover
explicit formatting precedence and preserved transforms/effects.

This step supplies the shared model, not a working native style gallery. XLSX
still needs style rendering, explicit chart formatting overrides, style
application, applicable per-chart-family catalogs and gallery previews. The
legacy 1 through 48 COM range is documented by
[Microsoft Chart.ChartStyle](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.chartstyle);
the modern ids used here were established by the local Excel 16.0 build 20430
probe, not inferred from that older reference.

Validation: the full XLSX suite passed 6,062 tests before the final unreadable-style
regression was added; all 48 focused reader, import/preservation and PowerPoint
runtime tests then passed. Strict and PowerPoint typechecks, the core build and
clean-consumer package imports passed. Seven chart-gallery browser tests passed
across all six bindings. A Playwright MCP review of the native style 212 workbook
confirmed chart selection and the contextual ribbon, and exposed the remaining
style-rendering gap described above.

## Imported native chart appearance

XLSX now paints imported title, legend and axis font sizes, theme-relative
major/minor fonts, colors, bold and italic settings. Direct chart-part formatting
overrides external style defaults, including no-fill and zero-width lines.
Deleted axes and hidden tick labels are respected. Chart and plot backgrounds,
axis/gridline colors and widths now follow imported formatting. Linear and
circle-path background gradients reuse the extracted PowerPoint chart gradient
contract and geometry. Each painted chart receives distinct gradient targets.

The recorder now saves COM title, legend and both axis font properties,
value-axis availability, and chart-area colors/positive line weights after
reopening each workbook. All 16 native column styles match the recorded text
properties in SVG, including the three deleted value axes and unavailable
style-204 title size. A shared DrawingML resolver option reproduces Excel's
half-channel HSL rounding for this appearance path; existing callers retain
their rounding default. This does not resolve the separately recorded extended
palette rounding gap.

Playwright MCP reviewed style 209's dark gradient and contrasting text. It also
exposed missing Office fonts falling back to serif; chart and cell text now
share the workbook font-stack helper. Matching font properties does not prove
pixel parity when the authored font is unavailable. Layout still uses estimated
text widths. Native style authoring/gallery previews, mixed-run rich text,
theme style-matrix references, series gradients/effects, rectangular path
gradients, exact focus/radius mapping, native plot placement and complete chart
families remain incomplete. Direct formatting is imported read-only; chart type
or series-count regeneration can still lose unmodelled chart-part details.

An Excel COM PNG export of style 209 supplied a direct visual reference.
That comparison exposed dropped alpha on its gridlines, now rendered at the
saved 10% opacity. The export also shows series gradients/shadows and native
plot positioning still differ; it is evidence of remaining work, not a claim
of pixel parity.

Validation: 6,081 XLSX core tests, 2,270 chart/color/diagram and PowerPoint
regressions, 315 XLSX UI tests, and 11 PowerPoint gradient/style UI tests passed.
The full browser suite passed 76 tests across six bindings; after the final
gridline-alpha correction, all 18 focused appearance tests and eight native-style
browser tests passed. Core and UI typechecks/builds and clean-consumer imports
passed (93 UI entry points and 45 custom-element registrations). Font-stack
extraction also passed the existing cell-paint tests. Native workbooks and the
COM-rendered PNG remain temporary acceptance artifacts.

## Imported series and point gradients

Imported series and point gradients now use the shared PowerPoint gradient
contract and SVG geometry. Cartesian series paints and explicit point paints
receive distinct gradient targets, with stop alpha retained. Chart color edits
replace the automatic theme base while retaining the ordered native shading
transforms; explicit RGB edits replace the gradient. Custom or mixed-base
gradients are retained. Point paint additions, replacements and removals are
patched into existing chart parts. Series gradients survive chart regeneration,
although regeneration still loses other unmodelled chart properties.

The shared DrawingML writer retains source gradient flags (including `scaled`
and `rotWithShape`), opaque stops, tile rectangles and extensions. Palette edits
retain existing series effects and axes. The document-order color resolver keeps
fractional intermediate channels and applies repeated transforms in order.
Native saturation can temporarily exceed the sRGB gamut before shading; this
path reuses the shared HSL conversion with an opt-in saturation setting.

`scripts/record-xlsx-chart-gradients.ps1` creates owned hidden Excel workbooks,
saves and reopens them, then records style 209 under palettes 10, 12 and 14.
The committed `excel-chart-gradients.json` fixture contains native parts and
all 18 COM-measured gradient stops. Rendering, palette edits, save/reload and
undo/redo match those measured stops. Reopening an edited full native workbook
in Excel also returned palette 14 and all six expected stop colors.

Playwright MCP reviewed the production chart and the Change Colors interaction.
Its screenshot confirms visible series gradients; the native PNG comparison
still shows missing rendered shadows, differing plot placement and bar widths,
and font-dependent differences. Source effects are preserved, not yet painted.
Series gradients on vary-colors pie/doughnut charts, complete scatter marker
paint semantics, rectangular path gradients and exact unscaled gradient geometry
remain incomplete. These checks do not establish full chart or Excel parity.

Validation: all 8,233 XLSX/chart/diagram/color core tests and 76 production
browser tests passed, along with core and XLSX binding typechecks, the core
build and clean-consumer imports of every core entry point. The native-style
browser case verifies rendered series stops and a live palette change; the
six-binding typography cases also remain green.

## Imported ordinary chart shadows

Series effect lists are now retained as opaque XML in the chart model, including
effects not yet painted. They survive type/series regeneration, can be replaced
or removed without replacing the series fill, and remain unchanged through
palette edits. Imported title text effects also supply their native shadow.
Cartesian, radial and radar SVG painters apply ordinary series drop shadows;
text uses separate filters. Legends retain their own appearance. Filter IDs are
unique across charts, and chart-space filter bounds keep horizontal lines and
short bars visible. The shared `feDropShadow` primitive comes from PowerPoint's
existing effect gallery; its original caller still uses the same primitive.

Fresh Excel COM measurements record style 209's series and title shadow colors,
blur, offsets, opacity and visibility for all three gradient palettes. SVG
geometry matches those properties after converting points to pixels. COM also
reopened a regenerated horizontal bar chart: both series remained, and its
shadow retained 4.5-point blur, 1.5-point vertical offset and 63% opacity.
Playwright MCP reviewed the visible title and eight bar shadows in production.

This closes the ordinary series/title shadow gap recorded above. Style 204's
scaled/skewed projected shadow still requires an affine effect painter. Point
effect overrides, effect DAGs, inner shadows, glow, reflection, theme effect
references and exact raster falloff remain incomplete. Title effects and rich
text can still be lost when the entire chart is regenerated. Native plot
placement, spacing and font availability continue to prevent pixel parity.

Validation: 8,238 core regressions and all 76 production browser tests passed.
After adding the effect-removal regression, all 14 focused native gradient and
shadow tests passed. Core/UI typechecks and builds, 30 PowerPoint gallery tests,
every core clean-consumer import, 93 UI entry imports and 45 custom-element
registrations passed. The shared PowerPoint caller retains its existing output.

## Native bar gap width and overlap

XLSX imports and paints `c:gapWidth` and `c:overlap` using the same cluster width
and spacing calculation as PowerPoint. Horizontal bars now order adjacent series
in Excel's direction. Clustered, stacked and percentage-stacked charts use the
same geometry. Gap width is measured relative to one bar's width; overlap changes
the total cluster span before the bar is sized. PowerPoint's existing callers
retain their absent-gap behavior, and its horizontal painter shares the helper.

Core chart edits validate the OOXML integer ranges (gap 0..500, overlap -100..100),
patch source plot properties without replacing axes, series or effects, and
support undo/redo. Regenerated bar/column charts retain authored spacing.
These fields are not yet exposed in a Format Data Series pane.

`scripts/record-xlsx-chart-spacing.ps1` records 13 owned native workbook cases with
three series and four categories: vertical/horizontal clustered charts under
five gap/overlap settings, plus stacked and percentage-stacked charts. A copy
of the native stacked file with its overlap declaration removed establishes
Excel's imported default of 0; new authored stacked charts explicitly write 100.
Saved and
reopened COM point rectangles establish bar width, category pitch and adjacent
series step. The renderer matches their normalized measurements and direction,
including 0/100 and 500/-100 extremes. COM also reopened our edited full native
workbook with gap 5, overlap 23, three series and the expected bar-width ratio.

Playwright MCP reviewed style 209's corrected geometry (native gap 100, overlap
-24); its browser regression verifies width relative to category pitch. This
closes the recorded gap/overlap discrepancy, while exact plot placement, font
measurement, reversed-axis settings and complete 3D charts remain incomplete.

The native fixture supplies independent geometry measurements rather than an
expected value computed with the renderer's formula.

Validation: 8,254 core regressions passed after adding the omitted-overlap case.
The full production browser suite passed 76 tests; the final native-style run
passed all eight tests across six bindings. Existing PowerPoint bar tests (16),
core/UI typechecks and builds, every core clean-consumer import, 93 UI imports and
45 custom-element registrations passed. Exact plot dimensions remain a separate
acceptance gap; normalized bar spacing does not establish whole-chart pixel parity.

## Docked chart series spacing controls

Chart Design now opens a docked Format Data Series pane for bar and column charts.
Double-clicking a painted bar opens it too. The pane reuses the editor's labelled
number fields, command registry, localization, theme tokens, protection checks,
core chart edit/history and source-preserving chart writer. Effective spacing
defaults are shared with the SVG painter through `chartBarSpacing` in core.
Five locales include the new labels and validation messages.

Changes repaint the chart, follow undo/redo and survive save/reload. A selection
change disables the fields or shows the appropriate empty state. Read-only and
protected-object settings prevent writes, even while the pane is already open.
Escape and Close return focus to the grid. Drawing selection now updates grips
without rebuilding chart contents, and a click with no movement no longer writes
a redundant anchor edit. Bar hit targets preserve their identity through selection;
pointer capture retains the actual series hit for double-click handling.

Playwright MCP reviewed the docked pane in production with a full native horizontal
chart. Excel COM reopened the workbook edited and downloaded from that browser:
gap width was 5, overlap 23 and all three series remained. Browser checks cover
editing, rendered spacing, save/reload and double-click/close in all six bindings.

This is the spacing section of the native pane. Its slider controls, axis assignment,
individual series/point formatting, fill/effect tabs and broader chart Format tab
remain incomplete; this increment does not establish full pane or UI parity.

Validation: all 82 production browser tests and 318 XLSX UI regressions passed.
After extracting effective spacing and adding the protected-pane regression,
all 14 native-style browser tests, 21 focused UI/locale tests and 15 native core
spacing tests passed. Core and UI typechecks and builds passed.

## Shared chart spacing sliders

The pane pairs native sliders with Gap Width and Series Overlap percentage fields.
Both use the range input template and focus/disabled styles extracted from the
shared Office zoom slider; the zoom control keeps its DOM and event contract.
Keyboard Home/End and arrow handling come from the native input. Input updates
the percentage and accessible value text; release commits one chart edit. Undo,
selection, locale and protection changes keep the paired controls synchronized.

Playwright MCP reviewed and dragged the slider in the production UI. Excel COM
reopened its downloaded full native workbook with gap width 500, overlap -24 and
three series. Six-binding browser checks cover keyboard End, paired values,
undo and disabled state. Continuous chart preview during pointer dragging,
axis assignment and the remaining individual series formatting controls are
still incomplete.

Validation: all 82 production browser checks passed, along with 31 shared range,
zoom, subclassing, chart-pane, chrome and manifest regressions. The UI typecheck,
refreshed core build and UI build passed. A real pointer drag in Playwright MCP
changed overlap from -24 to 85; one undo restored -24.

## Individual chart series fills

Double-clicking a bar now opens the pane for that specific series. A Series
selector uses the core chart view's current names and keeps its choice through
edits and undo. Solid fill and No fill reuse core chart edits and the ribbon's
existing theme/tint/standard color picker. Stale color menus cannot edit another
series or workbook. Read-only and object protection disable these controls.

Excel COM confirms that series-level solid/no-fill edits replace point paint
overrides. The core helper matches this behavior while retaining series effects,
references, untouched series and native axes. Theme tints now serialize as
DrawingML luminance transforms instead of losing their tint. Saved/reopened COM
measurements verify the rendered tint color independently of the implementation.
The recording script and fixture retain the Excel version and build.

Playwright MCP reviewed a full native horizontal chart, selected its second
series, chose Accent 1 lighter 40% and downloaded it. Excel reopened the file
with three series and brightness 0.4. All 88 production browser checks passed,
including series selection, fill editing, saving, undo and protection across
six bindings. All 6,111 XLSX core regressions passed before the additional COM
fixture regression, which also passed. Core/UI typechecks and builds passed.

Imported gradients remain displayed and preserved; gradient, pattern and picture
fill authoring, point formatting, transparency, fill/effect tabs and axis
assignment still need implementation. The pane layout and chart rendering do
not yet establish pixel parity with Excel.

## Solid chart fill transparency

The series pane now pairs a Transparency percentage field with the shared range
control. Core edits retain theme/tint transforms, replace ordered alpha transforms
and support save/reload, undo and read-only/protection guards. Choosing another
color keeps its opacity, matching a saved/reopened native Excel COM probe.
Gradient-to-solid conversion stays in core and preserves the first stop's color
and alpha instead of trying to turn a CSS rgba value back into RGB.

Solid series and point rendering previously discarded the alpha from the shared
DrawingML resolver. Both now reuse the CSS paint helper extracted from chart
appearance, including explicit solid DiagramML fills. Native imported 0%, 37%
and 100% fills render with their saved opacity. Gradient stops continue using
their existing shared alpha handling; the UI enables transparency editing only
for solid fills.

Playwright MCP opened the full native 37% workbook, confirmed its rgba paint,
changed the field to 23% and downloaded it. Excel COM reopened that file with
transparency 0.23000002, a visible red fill and the original series. Core native
regressions cover alpha rendering, round-trip and color-change preservation;
six-binding browser checks cover percent editing, slider keyboard input, undo,
save and disabled state. Mixed per-point fill behavior, gradient/pattern/picture
authoring, fill/effect tabs and exact pane layout remain incomplete.

Validation: all 6,116 XLSX core tests passed; the expanded native regression
checks and 23 focused UI/locale tests passed. Core/UI typechecks and builds
passed. The full production browser run passed 87 checks; its one failed
new-workbook navigation reported `ERR_NO_BUFFER_SPACE` and passed on retry.

## Gradient series fill authoring

The pane can create a linear gradient and edit its angle, stop position, themed
color and transparency, add stops and remove stops down to the native two-stop
minimum. A shared DOM stop strip supports click and keyboard selection. Solid
controls are hidden when gradient controls apply. Core preserves imported
gradient flags, geometry, extensions and untouched chart parts; history and
save/reload use the existing chart edit pipeline. Imported path gradients retain
their geometry and allow stop edits, with the linear angle field disabled.

PowerPoint's stop sorting/update/removal helpers now live in shared core and its
existing picker delegates to them for all five bindings. Native Excel COM shows
that inserted stops can remain in insertion order in both the collection and
saved XML. The XLSX model keeps those identities; shared SVG/CSS painters sort
only the visual paint. This fixes imported unsorted gradients without assigning
an edit to a different stop.

Playwright MCP edited a full native workbook and downloaded it. Excel COM
reopened it with angle 54 degrees, two stops, the first at 23% with 37%
transparency, and the original red/white colors. Native regressions include
angle/position/transparency, insertion, deletion, minimum count, source flags
and unsorted-stop rendering. Six-binding browser checks cover color/opacity,
selection, add/remove, save, undo and disabled state. All 94 production checks
and 6,122 XLSX/shared-gradient core tests passed before the final pane visibility
adjustment; focused native-style browser checks cover that adjustment.

Preset/type/direction controls, brightness, stop dragging and paired stop sliders,
mixed point formatting, pattern/picture authoring and exact native pane layout
remain incomplete. New gradients start as a 90-degree primary-color-to-white
fill; this default does not claim full native preset parity. The source-alias
workflow now also resolves shared PowerPoint editor leaf modules, required by
embedded shared UI dependencies.

Final validation: all 26 native-style browser checks passed after the visibility
and compact stop-button adjustments. The 24 focused UI/locale checks, eight
PowerPoint gradient-picker/paint regressions, three source-alias regressions,
core/UI typechecks and builds passed. Clean-package checks imported 96 entries
and registered 45 custom elements. A newly created gradient downloaded from
Playwright MCP reopened in Excel with two stops at 90 degrees, red/white
colors, and the first stop's retained 37% transparency. The additional no-fill
creation regression verifies the chosen series' palette and unchanged siblings.

## Gradient stop brightness

The series pane now exposes brightness from -100% to 100%. Gradient edits and
SpreadsheetML theme tints reuse one format-neutral DrawingML luminance helper.
Native stop identity, theme choice, transparency, geometry and the other stops
survive edits, saving and undo. Zero brightness uses Excel's explicit 100%
luminance multiplier. Arbitrary imported luminance expressions are preserved;
the brightness field is blank and disabled when no canonical value is known.

Excel COM saved/reopened -42%, 0%, 100%, -100% and 37% cases independently.
Core round trips match their stop transforms exactly, and rendered stop colors
match the native RGB getters. Playwright MCP reviewed the pane and downloaded
a 37% edit: Excel reopened it with 37% brightness, 37% transparency, angle 54,
positions 23% and 56%, and the untouched green stop's 13% transparency.
Six browser checks cover all XLSX bindings, brightness extremes, invalid values,
retained opacity and read-only controls. The broader core selection passed
6,388 tests; 24 focused UI checks and core/UI typechecks passed.
Core/UI builds and clean-package checks also passed, importing 96 entries and
registering 45 custom elements.

This does not close the remaining gradient gaps: native presets, type and
direction galleries, dragging, paired position/brightness/transparency sliders,
mixed point formatting and exact pane layout. COM's explicit RGB property
replacement also reset brightness and opacity in the new recorder case;
the viewer's color-menu opacity retention is an existing policy whose native
interactive-menu equivalence remains unverified.

## Linear gradient direction gallery

The pane now offers eight standard linear direction tiles. It reuses the
PowerPoint/shared Office gallery's popup, focus, keyboard navigation and safe
SVG preview handling; the common chart painter generates each preview. Every
preview has its own ID, including across multiple editor instances. Direction
picks use the existing core angle command and preserve stops and opacity.
Imported path gradients keep this linear-only control disabled.

Excel COM saved/reopened angles 0, 45, 90, 135, 180, 225, 270 and 315. Core
editing round trips match these angles and retain the native stops and flags.
Playwright MCP reviewed all eight distinct previews and downloaded a 315-degree
edit: Excel reopened it with the original 37% brightness/opacity, 23% position
and the other stop's 56% position and 13% transparency. Six framework browser
checks cover keyboard selection, popup closure, undo and read-only behavior.
The focused UI/gallery selection passed 30 tests and the gradient core suite
passed 22 tests.
Core/UI typechecks and builds passed. Clean-package checks imported 96 entries
and registered 45 custom elements.

This closes linear direction authoring, not exact raster parity. A separate
Excel-exported 45-degree red/white chart revealed a diagonal paint difference;
gradient interpolation and endpoint behavior need a measured pixel corpus
before adjusting the shared renderer. The recorder now exports native direction
PNGs for that comparison. Radial/rectangular/path direction controls, exact
native gallery arrangement, preset/type galleries and stop dragging remain open.

## Gradient stop dragging and live preview

The shared stop strip now supports mouse dragging. The strip, position readout
and rendered chart preview the chosen stop's position while the workbook remains
unchanged. Release routes one edit through core, so one undo restores the whole
gesture. Crossing another stop retains native insertion-order identities; only
paint order changes. Pointer capture keeps the gesture active beyond the strip.
Escape, pointer cancellation, model/selection refresh and read-only transitions
discard the preview. Preview paints all frozen-pane copies of the chosen series
through the existing shared chart painter. Unresolved stop colors disable the
strip when painted stops cannot be mapped safely to model identities.

Playwright MCP reviewed a crossing drag from 23% to 85%: workbook positions
stayed `[23, 56]` during preview, became `[85, 56]` on release, and returned to
`[23, 56]` after one undo. Excel 16.0 build 20430 reopened the downloaded file
with the 85% stop, original 37% brightness/transparency, 54-degree angle and
the unchanged 56% green stop's 13% transparency. Independent COM captures
also cover 85%, 0% and 100%; core round trips match native stop identities,
colors, positions and opacity. The gradient core suite passed 26 tests.
Seven focused UI suites passed 27 tests, covering crossing, clamping, Escape,
secondary-pointer cancellation, read-only refresh, removal and SVG restoration.
All six final browser checks passed, including a read-only transition mid-drag.
Core/UI builds passed; clean-package checks imported 96 entries and registered
45 custom elements.

Mouse drag behavior is verified across all six bindings. Exact native fractional
drag quantization and touch/pen behavior remain unverified. Stop sliders,
preset/type galleries, path direction controls and diagonal raster parity remain
open. This increment adds interaction without changing gradient interpolation.

## Paired gradient sliders

Position, brightness and transparency now pair number fields with native range
controls. Their shared number/range binding also serves solid fill transparency.
The gradient fields follow Color, Position, Brightness, Transparency order.
Dragging previews chart paint, stop markers and color through prospective core
edits while preserving the workbook. Change commits one history step; Escape,
pointer cancellation, selection refresh and read-only transitions discard the
preview. Unknown stop colors and unsupported brightness transforms keep the
corresponding controls disabled.

Playwright MCP reviewed the pane and exercised slider keyboard endpoints.
Excel 16.0 build 20430 reopened its downloaded file with the edited stop at
100% position, -100% brightness and 100% transparency, retaining the other
stop's 56% position and 13% transparency and the 54-degree angle. Browser tests
exercise Home/End, undo and read-only controls in all six bindings. Eight focused
UI suites passed 28 tests, including cancellation and preview without model
mutation. Strict UI typechecks and core/UI builds passed.
All 26 chart browser checks passed; clean-package checks imported 96 entries
and registered 45 custom elements. A Playwright MCP input-event preview changed
the transparency readout to 37% while leaving the workbook unchanged; Escape
restored the original 100% readout.

This closes paired slider authoring. Exact native pane tabs/layout, presets,
path-gradient controls, mixed point formatting, fractional/touch behavior and
diagonal raster parity remain open.

## Measured scaled gradient raster correction

DrawingML's linear `scaled` flag now survives parsing, model edits and saves.
Scaled chart gradients extend their vector across the normalized bounding box;
the previous unit-length diagonal vector clipped the outer colors too early.
Opaque two-endpoint scaled linear gradients use the native sigma/gamma-2.2
paint curve extracted from Visio into shared color logic. Source/model stops
remain unchanged; only the rendered SVG uses additional sampled stops.
Newly authored series gradients explicitly save the scaled flag.

Excel COM recorded ten angles in square and wide chart areas, with independent
SaveCopyAs/reopen checks and valid PNG exports. Core regressions parse the
native saved fill XML and match 500 background sample pixels within two RGB
levels. Browser raster checks use the actual chart DOM definitions and match
the same corpus across all six bindings. Playwright MCP reviewed a wide native
45-degree workbook, confirming corner-to-corner endpoints and 256 paint stops.
The gradient's lighter intermediate colors now match native Excel's profile.
Excel reopened the browser download with exactly two original red/white stops
at 0%/100% and the unchanged 45-degree angle; sampled paint is not serialized.

The core selection covering chart/color/DrawingML/XLSX and relevant Visio
regressions passed 8,372 tests (13 skipped). Core/UI typechecks and builds
passed, as did 11 focused UI tests and four PowerPoint gradient regression
tests. The 26 existing chart browser checks passed; the six new raster checks
passed after synchronizing each file load with its new gradient geometry.
Clean UI package checks imported 96 entries and registered 45 custom elements.
The clean core package check imported every entry point.

The measured raster correction covers opaque endpoint-pair scaled linear
gradients. Transparent, interior/multiple-stop, unscaled and path profiles
still need independent raster measurements. PowerPoint's existing descriptors
keep their prior paint behavior until their native profile is measured and
explicitly enabled. Presets, path-gradient controls and native pane layout
remain open; this evidence does not establish full Excel UI parity.

## Extended linear gradient measurements

Excel COM now captures six additional scaled linear profiles at ten angles in
square and wide bounds: transparent endpoint pairs, opaque interior pairs,
opaque three-stop fills, translucent three-stop fills, crossed pairs and
coincident pairs. Each capture saves/reopens its own workbook and records the
saved fill XML and native PNG alpha alongside its pixels. The combined corpus
contains 140 native cases and 3,500 points, including the original opaque pair.
Native stop insertion order, opacity and geometry round trips are checked for
every profile; renderer sampling leaves the editable fills unchanged.

The existing straight SVG path passes the new transparent two-stop, interior,
crossed and opaque three-stop comparisons. Alpha is checked separately within
two byte levels; colors are compared after premultiplication, with the original
two-level opaque tolerance and three levels for translucent pixels. This avoids
interpreting unpremultiplication noise at low alpha as a paint algorithm.

Two native differences remain explicit. At `(306, 66)` in the 600x600 native
135-degree translucent three-stop capture, Excel reports RGBA `[114,138,0,193]`
while Playwright MCP reports `[118,136,0,193]`, a red premultiplied error just
over three levels. A wide 135-degree coincident-stop boundary also has native
edge coverage absent from the SVG hard step. The fixtures retain both problems;
strict expected-failure repros distinguish them from passing comparisons.
This increment expands the native evidence, not the renderer's parity claim.

Core raster checks report 124 passing tests and 17 expected failures. Across
all six bindings, 2,975 points per binding pass strict browser comparisons;
twelve additional browser repros remain expected failures (the two gaps in
each binding). Strict core typechecks pass. Unscaled/path gradients, additional
translucent profiles, presets and native pane layout still need work.

## Native named gradient presets

The series fill pane now offers all 24 standard named Office gradient presets,
using the shared Office gallery, keyboard navigation and gradient thumbnails.
The palette and factory live in the format-neutral diagram area; the XLSX
editing command applies a complete preset as one undoable change. Selecting
a preset retains other series and clears individual point overrides consistently
with other series fill edits. The pane fits without horizontal scrolling.

Excel COM captures all 24 presets using horizontal style 1, variant 1, then
saves and reopens each workbook. Tests compare the factory's exact saved stops,
colors and scaled 90-degree geometry with those captures, and check round trip,
undo, redo and repeated no-op application. All six bindings exercise every
preset, popup keyboard selection, read-only behavior and download serialization.
Playwright MCP downloaded the Rainbow preset; Excel COM recognized preset 16,
with its seven expected stops and 90-degree angle after reopening.

The 32 chart browser checks pass, including the existing gradient drag/preview
regressions. The controls reuse the same gallery as PowerPoint and the same
core paint helpers as other Office consumers. This covers the standard named
presets, not the modern theme-dependent preset menu, other native preset style
variants, path-gradient controls or a pixel comparison of every preset. The
previous translucent and coincident-edge raster gaps remain open.

## Rectangular path-gradient rendering

The shared chart resolver previously converted every DrawingML path gradient
into an elliptical SVG radial gradient, even when Excel saved `path="rect"`.
Chart paint now preserves the path and target rectangle and uses the existing
PowerPoint nested-rectangle painter, extracted into the shared diagram area.
PowerPoint retains its existing defaults through a thin compatibility entry.
XLSX uses higher-resolution bands; an independent alpha mask prevents
translucent bands from accumulating opacity. Existing stop editing previews
and cancellation now update and restore the rectangular pattern's image.

Excel COM supplies 24 independent opaque/translucent center/corner captures
in square and wide bounds, with 600 RGBA points. Core and all six browser
bindings match these captures within the existing tolerances. The focused
core selection passes 230 tests, with the existing 17 expected raster failures;
PowerPoint's two affected rendering suites pass 21 tests. Browser coverage
passes 50 ordinary checks plus the existing 12 expected-failure repros.
Six of the ordinary checks cover rectangular stop-drag previews without
premature workbook edits, Escape restoration, transparency changes and undo.

Playwright MCP independently measured a translucent center fill with alpha
error of one byte and premultiplied color error below 2.3 levels. Excel COM
reopened its viewer download as center style 7, variant 1, with the expected
37%/13% stop transparency. Native type/direction authoring controls remain
to be implemented; circle/shape paths, arbitrary target rectangles and
additional multi-stop path profiles remain unverified. The earlier linear
translucent and coincident-edge discrepancies are still explicit.

## Rectangular gradient authoring controls

The series fill pane now switches between Linear and Rectangular types and
offers the measured center and four corner directions through the existing
Office gallery. It preserves the selected stop, all stop values and color
transforms. Each geometry change uses one core history operation; repeated
selection of the current type/direction is a no-op. Switching back to Linear
clears the path target and resets tile insets instead of carrying corner
extension flags into the linear fill.

DrawingML tile insets are now typed data in the shared model and patched by
the shared writer. The format-neutral geometry helper retains original root
flags and extensions. Tests compare target/tile values against all 24 native
captures and verify saving, undo and redo. All six bindings exercise both
types, every rectangular direction, selected-stop retention and read-only
guards. Imported Radial, Path and unknown types remain selected without
conversion; Radial/Path authoring choices are disabled until independently
measured and implemented.

Playwright MCP reviewed the popup and exported a five-stop bottom-right fill.
Excel COM reopened it as corner style 5, variant 4, with all five original
positions and RGB colors. The pane has no horizontal overflow. The broader
diagram/chart/XLSX core sweep passes 8,074 tests, with the existing 17 expected
raster failures. The 62 browser checks pass, including the existing 12
expected-failure repros. These results establish the implemented geometry and
editing behavior, not full native pane layout, every gradient type or every
multi-stop raster profile.

## Native circular and rectangular-shape chart backgrounds

Imported `path="circle"` backgrounds now remain circular in physical paint
bounds instead of stretching into ellipses. Their radius reaches the farthest
corner and their SVG transform preserves the native center/corner focus.
Chart areas supply chart bounds; plot areas supply their final bounds after
title, legend and axis layout. Imported `path="shape"` backgrounds use the
existing shared rectangular path painter when the painted outline is a chart
or plot rectangle. Legacy radial descriptors keep their existing behavior.

The extended Excel COM recorder captures 48 opaque/translucent center/corner
profiles across square and wide bounds, with 1,200 RGBA samples. Circle/shape
paths are imported as explicit OOXML, resaved by Excel and rendered natively
on Excel 16.0 build 20430. All six browser bindings meet the unchanged native
pixel tolerances. The affected core selection passes 2,090 tests plus the
existing 17 expected raster failures; strict core/UI typechecks, builds and
clean package import checks pass. PowerPoint's shared gradient consumers
retain their existing compatibility descriptors; their affected suites pass
12 tests. The browser selection passes 30 ordinary checks plus the existing
12 expected-failure repros.

Playwright MCP reviewed an Excel-saved wide circular fill and downloaded it.
Excel COM reopened and resaved the download with `path="circle"`, two stops,
positions 0/1 and the original red/white RGB values. This validates the fill
round trip, not every surrounding chart element's native layout.

Radial/Path authoring remains disabled. Series marks still need their own
physical gradient bounds and independent native comparisons; nonrectangular
shape outlines, arbitrary target rectangles, other stop profiles and full
native pane layout remain open. Existing translucent three-stop and
coincident-edge discrepancies remain explicit expected failures.

## Circular and shape fills on bar/column series

Each bar/column mark now receives its own circle/shape paint using its actual
rounded bounds. A series with unequal values retains circular fills on every
mark instead of sharing one stretched gradient. Point fills keep their own
paint and focus; solid point overrides are retained. Clustered, stacked,
percent-stacked, negative, zero and missing values have geometry regressions.
The shared core geometry command now authors circle and shape paths while
preserving stops, root flags and extensions. The shared Office gallery supplies
five directions for Rectangular/Radial/Path and retains eight linear directions.
Preview gestures update all inherited marks without changing the workbook until
release. Undo, cancellation and read-only behavior remain guarded.

Radial and Path choices are enabled for bar/column series rectangles. Other
chart types retain disabled choices while keeping imported paths visible.
The existing PowerPoint gallery and rectangular painter are reused, together
with the shared chart resolver and stop sampler; no second gradient engine
was added to XLSX UI.

Excel 16.0 build 20430 supplies 96 column/bar captures with 2,400 native RGBA
points, including opaque/translucent center and corner variants at two chart
sizes. Explicit OOXML paths are imported and resaved in Excel before capture.
A second solid fill measures the mark's full bounds, avoiding gridlines and
white endpoint plateaus. Browser raster checks match physical mark aspects
independently of surrounding chart layout. Of the 96 profiles, 94 meet the
unchanged tolerances; two remain strict expected-failure repros: translucent
top-right circle at pixel 391,479 differs by about 3.04 premultiplied levels
(limit 3), and opaque centered shape at 302,376 differs by 3 levels (limit 2).
These are sampling/quantization gaps, not grounds to claim exact native paint.

The diagram/chart/XLSX core sweep passes 8,323 tests plus the existing 17
expected raster failures. The six-binding browser selection passes 42 ordinary
checks plus 24 expected-failure repros, including the two new series gaps.
Playwright MCP reviewed the Path gallery without
horizontal pane overflow and downloaded a bottom-right series fill. Excel
COM reopened and resaved it with `path="shape"`, the requested target point,
two stops, positions 0/1 and the original red/white RGB values. Native type
menu/pane layout certification, nonrectangular marks, arbitrary path targets,
additional stop profiles and the earlier raster gaps remain open.

The fill comparison exposed an axis-visibility export gap, now corrected by
the shared chart axis writer. New charts, kept chart parts and type/grouping/
series-count regeneration serialize explicit axis and label visibility.
Unchanged imports retain their original boolean spellings and high/low label
positions, including hidden axes whose effective labels are invisible.
Other direct formatting edits, including fonts, outlines and effects on
regenerated chart parts, still need export support.

Repeated save/reload regressions cover column, bar, line, area, scatter and
radar charts. Excel 16.0 build 20430 reopened and resaved twelve exports with
the requested `delete` and `tickLblPos` flags intact. Six browser bindings
verify category-label rendering and UI export for three visibility states.
Playwright MCP reviewed a 600 by 400 hidden-axis chart and downloaded it;
Excel COM reopened and resaved that download with zero visible axes and both
`delete="1"`, `tickLblPos="none"` pairs intact. This verifies visibility and
serialization, not native chart layout or complete Excel parity.

Direct chart-element fill export now uses the shared DrawingML fill writer.
New charts, regenerated parts and kept-part patches retain supported chart
area, plot area, title and legend fill metadata. Kept-part tests preserve
outlines, effects and extensions, including newly inserted legends and
replaced referenced titles. Existing series fill editing uses the same shared
DOM replacement helper. Unsupported picture/pattern fill editing, generated
font/outline/effect preservation and title/legend background painting remain
open. Background pane authoring is covered by the next iteration below.

The diagram/chart/XLSX sweep passes 8,349 ordinary tests plus the 17 existing
expected raster failures, with an additional focused title/legend insertion
regression passing afterward. Six browser bindings verify no fill, solid and
four gradient geometries through UI export. Excel 16.0 build 20430 reopened
and resaved all six variants on four chart elements. Playwright MCP reviewed
and downloaded a circular chart/plot background; Excel COM retained circle
paths, center focus, red/white colors and endpoint positions on all four
fill-bearing elements after resave. These checks verify fill metadata and
background rendering, not exact native layout or full formatting UI parity.

Chart/plot background formatting now has a docked pane opened from Chart
Design or a painted background double-click. Its target selector, no-fill,
solid and gradient authoring reuse the series controls. All model mutations
stay in core, and stop/range previews affect only the selected painted
background until commit. Target changes and pane closure cancel previews.
The pane follows undo/redo, disables editing in read-only mode and closes
when the series pane opens. Transparent plot hit regions, other formatting
targets, borders/effects and exact native pane layout remain open.

A shared-gallery correction prevents imported radial focal points from
making all direction thumbnails appear centered. Playwright MCP reviewed the
corrected five previews without horizontal pane overflow. Excel 16.0 build
20430 reopened/resaved a workbook authored through that UI and its save API:
a bottom-right seven-stop circle on chart area retains 37 percent first-stop
transparency, and a top-left two-stop shape on plot area retains its red/white
colors. Paths, stops, color transforms and target/tile rectangles match the
pre-resave export exactly. This proves those edits and their serialization,
not native full-chart raster fidelity or complete formatting parity.

The same fill pane now includes present chart titles and visible legends.
Their backgrounds paint behind text, and their gradient geometry uses their
own layout boxes rather than the whole chart. Five legend positions have
paint-order regressions. Target removal disables its option and falls back
to chart area; painted title/legend double-clicks open the corresponding
pane. Eighteen browser checks cover all four fill targets and existing
series controls in six bindings.

The repeatable COM recorder `scripts/record-xlsx-chart-decoration-fill.ps1`
captures five native legend positions, PNGs, fills and font/box measurements.
The JSON reference records Excel 16.0 build 20430. It exposed a concrete
gap: native title font is 18 points and legend font is 10 points, while the
renderer fallbacks were 14 and 10 pixels. Built-in text defaults are now
resolved as described below. Root chart text inheritance, actual font
measurement, manual/automatic title and legend layout, borders/effects and
transparent hit targets remain incomplete.
Playwright MCP reviewed and edited the native import. Excel reopened/resaved
the editor's export with a bottom-right circle title (37 percent first-stop
transparency) and top-left shape legend, retaining all modeled fill metadata
exactly. These checks establish fill authoring and preservation, not exact
native geometry or whole-product parity.

Built-in chart text defaults now use the shared chart style model and workbook
theme resolver. The supported Office 2010 style choice takes precedence over
its legacy fallback; external styles and direct element formatting retain
their precedence. New charts use modern style 102, and regenerated imported
charts retain their known built-in style. Unknown styles receive no invented
font defaults. Ninety-six independently reopened Excel references cover legacy
styles 1-48 and modern styles 101-148, including dark-style text colors.

The existing COM style recorder supports this reference workflow; preparation
uses the shared XML parser/writer to isolate built-in styles from direct and
external text formatting. New and replaced chart titles reuse one rich-text
writer with the empty property nodes Excel needs for inherited fonts. Without
those nodes, regenerated titles reopened with invalid font sizes.

The chart/XLSX sweep passes 8,399 ordinary tests plus 17 existing expected
raster failures. All 56 focused browser checks pass, covering built-in fonts
before and after chart type edits and existing fill controls in six bindings.
Playwright MCP reviewed the native import:
the modern title renders at 24 pixels and legend/axes at 13.33 pixels. Excel
reopened and resaved its UI export as a line chart with the expected 18-point
bold Aptos Narrow title and 10-point legend/axis fonts. Root chart text
inheritance, mixed rich-text runs, East Asian/complex-script font choices,
native text measurement/layout and built-in paint defaults remained open at
that revision. Chart-space inheritance is addressed below.

Chart-space text bodies now supply inherited font size, weight, italic, face
and color through the shared chart resolver. An inherited title uses 1.2 times
an explicit chart-space size; direct title/legend/axis properties override it.
The reference preparation/COM recorder now covers 24 cases across legacy,
modern and external chart styles, with partial defaults and direct rich-text
overrides. The recorder reads the title's actual TextFrame2 text-range font:
its title-wide Font property can report the inherited size while a rich run
displays a different explicit size.

Chart regeneration preserves imported text bodies and the first title
paragraph/run properties using the existing XML element mapping. This retains
root, title, legend and axis fonts through chart type edits. It does not add a
second text engine, and source-only formatting metadata is not a new font
authoring UI. Mixed-run/multiple-paragraph titles, font measurement, text
effects authoring and native title/legend layout remain incomplete.

The chart/XLSX sweep passes 8,461 ordinary tests and 17 existing expected
raster failures. All 74 focused browser checks pass across six bindings.
Playwright MCP reviewed inherited italic Arial and explicit title formatting
through chart type edits and the save API. Excel 16.0 build 20430 reopened
both line-chart exports, resaved them, and reopened those copies: inherited
title 24 points, legend/axes 20 points, all italic Arial and color #123456;
explicit title 32 points, bold, major theme font and color #595959, with
9-point legend/axes. These checks establish those font inheritance and
preservation paths, not exact raster or complete Excel parity.

Imported mixed-format chart titles now reuse the shared DrawingML text body
reader and chart/theme style resolver. Runs keep their own size, font, weight,
italic, color and simple underline; explicit line feeds and paragraph boundaries
render on separate lines. Paragraph defaults apply before run properties. The
shared reader also retains cached fields and breaks in their original order.
Unchanged title text retains its entire rich body when chart topology is
regenerated, instead of copying only the first run's properties. Plain title
replacement still uses the existing first-run formatting behavior.

The native recorder creates single-line and line-feed mixed titles and measures
individual characters on reopen. A shared-XML preparation step isolates a third
explicit-paragraph reference, which Excel reopens and measures independently.
The test-only chart parts and measurements are in
`src/core/chart/excel-chart-title-text.json`. Per-character comparison verifies
the resolved runs, and browser tests reuse the existing native chart shell in
all six bindings. The diagram/chart/XLSX sweep passes 8,534 ordinary tests plus
17 existing expected raster failures; all 92 focused browser checks pass.

Playwright MCP reviewed the title and changed the native paragraph reference
to a line chart. Excel 16.0 build 20430 reopened its UI export, resaved it, and
reopened the copy: all 23 characters retain size, font name, bold, italic,
underline and color. The visual review still exposes approximate horizontal
text measurement and line spacing. Automatic wrapping, native text box layout,
advanced underline/baseline/spacing effects, list-level defaults, mixed-format
font authoring UI and full Excel parity remain unfinished.

Chart titles and legends now accept host font measurements through
`renderChartSvg(..., { measureText })`. The grid supplies its existing cached
canvas measurer, using the same CSS family, size, weight and italic style as
SVG painting. Rich-run advances and centering, plain title background bounds,
legend widths and truncation share one core helper. Headless callers keep the
existing estimate when no measurer is supplied; invalid widths also fall back.
Measured truncation retains complete Unicode code points.

Playwright MCP observed the native paragraph reference before/after the change:
the Revenue advance is about 133.37 CSS pixels. Previously the following space
started only 123.2 pixels later, overlapping Revenue; it now starts about
133.39 pixels later. Forecast moves from x=261.33 to x=271.85 and is centered
using its measured advance. Native Excel COM reopened all three references:
the title box is 141.62496 points wide, 30.59504 points high for the single
line and 54.04504 points high for either two-line variant. Screenshot review
still shows different title padding/placement and vertical spacing. This change
does not establish native text box geometry, wrapping or complete chart parity.

The follow-up diagram/chart/XLSX sweep passes 8,539 ordinary tests with the same
17 expected raster failures. Core/UI builds and typechecks pass; all 92 focused
chart browser checks pass across six bindings. Browser rich
text regressions wait for SVG layout, then compare actual rendered advances
and centering across six bindings before/after chart edits.

Chart title vertical layout now reuses the grid canvas's font bounding ascent
and descent through the optional `measureFont` SVG hook. Mixed lines align
all runs on one baseline and advance by the previous line's largest descent
plus the next line's largest ascent. The title's height sums natural font
boxes plus the existing padding; plain title backgrounds use the same metrics.
The first baseline keeps its existing anchor. Missing/invalid metrics retain
the previous headless behavior. No second canvas or viewer layout algorithm
was introduced.

The native recorder captures title boxes on independent reopen. Three measured
references are retained in test-only `excel-chart-title-geometry.json`. With
font boxes measured in Chromium on Windows, single/multiline title heights are
within 2 CSS pixels of Excel 16.0 build 20430's boxes (about 0.79 and 1.06 pixels
respectively), compared with the previous larger font-size estimates. These
are bounded comparisons for the three mixed references, not a general tolerance
claim for other fonts or platforms.

Playwright MCP imported the native paragraph reference: Forecast's baseline
moves from 83.47 to 72 CSS pixels while Revenue remains at 40. In the owned
Excel PNG (960 pixels wide), green title pixels occupy y=82..111; normalized
to the viewer's 640-pixel chart this is y=54.67..74. The previous browser crop
had y=66..88 and the updated crop has y=55..76. These color extents verify an
improvement for that text, not complete chart raster parity. The chart's plot
position, borders/gridlines, title padding, automatic wrapping, paragraph
spacing and manual layout remain incomplete.

The final diagram/chart/XLSX sweep passes 8,543 ordinary tests with the same
17 expected raster failures. All 92 focused chart browser cases pass, including
actual mixed-line baseline checks before/after edits in all six bindings.
Core/UI typechecks, builds and package import/registration checks pass. The
native geometry fixture is absent from the published build output.

The shared DrawingML reader now retains `a:lnSpc`, `a:spcBef` and `a:spcAft`
in their original points/percentage units. XLSX title views carry that metadata
into rich rendering, including single-line titles with spacing properties.
The painter resolves percentage spacing against natural font line height and
point spacing through the existing chart unit converter. Literal title line
feeds and explicit paragraphs retain their spacing during line expansion.
The existing rich-body XML preservation path keeps those properties through
chart regeneration; no viewer-side spacing parser or serializer was added.

`scripts/record-xlsx-chart-title-spacing.ps1` creates 12 owned Excel references:
default, expanded/compressed percentage and fixed-point line spacing,
point/percentage before/after spacing, explicit paragraphs, and single-line
variants. Independently reopened parts and title rectangles are retained in
`src/core/chart/excel-chart-title-spacing.json`, outside the published output.
Core comparisons use independently observed Arial canvas font boxes; title
heights differ from native by less than 2 CSS pixels for these references.
Malformed spacing is ignored by the parsed view and remains in source XML.

Playwright MCP reviewed the 75% and before/after references. The latter keeps
baselines 40/85.67 CSS pixels before/after the actual Change Chart Type flow.
Excel 16.0 build 20430 reopens its line-chart export, resaves and reopens the
copy with SpaceWithin=1, SpaceBefore=6, SpaceAfter=8 and title height 72.4 points
unchanged. Independently reopened native PNGs match the authoring PNG positions.
The compressed-spacing comparison still shows a few pixels of first-line
placement difference. Centered leading is approximate; automatic wrapping,
empty/break/field/list-level spacing cases, manual text box layout, native
padding and complete chart raster parity remain open. Paragraph-spacing
editing controls are not implemented by this import/render/preservation work.

The diagram/chart/XLSX sweep passes 8,560 ordinary tests with the same 17
expected raster failures. All 164 focused chart browser cases pass, including
72 new native-spacing cases across six bindings. Core/UI typechecks, builds
and package checks pass. The spacing fixture is absent from published output.

### Automatic chart-title wrapping (8 October 2026)

Long plain and rich chart titles now wrap using the shared core text helper
extracted from PowerPoint's static SVG converter. Words spanning formatting
runs stay together; XLSX splits oversized words at grapheme boundaries. The
host still supplies font measurements, and the existing appearance resolver
supplies run fonts/styles. Soft breaks do not modify title text or rich XML.
Before/after paragraph spacing applies at the automatically wrapped paragraph
edges, while explicit title line feeds retain the previously verified behavior.

`scripts/record-xlsx-chart-title-wrap.ps1` owns hidden Excel instances and
creates ten references at chart widths from 160 to 480 points. Independent
reopen captures are in `src/core/chart/excel-chart-title-wrap.json`.
`excel-chart-title-wrap-browser-metrics.json` records actual Chromium 154
canvas measurements for the reference fonts/text. Core comparisons verify
native line counts, selected native-image line contents, title widths within
2 CSS pixels and heights within 4, plus exact rich-body preservation through
chart regeneration. These are bounds for these references, not universal
font/layout tolerances. A separate regression covers long/short plain titles.

Playwright MCP found that rounding the 200-point chart to 267 CSS pixels
changed a word break. The grid now passes its fractional dimensions into the
existing painter. The corrected reference displays `InternationalBusin`,
`essRevenueForec`, `ast`, matching the native PNG before/after an actual type
change. Excel reopens that UI export as type 4 (Line), resaves and reopens it
with the full title and native 152.07 by 65.1-point title rectangle unchanged.
The screenshot review still exposes different plot geometry/gridline paint.
The six-binding 200-point browser case also resizes the drawing with its
pointer grip, verifies the native 320-point word breaks, then checks undo,
redo and restoration before the type change and save.

The expanded `vitest run diagram chart xlsx text
pptx/converter/svg-text-lines.test.ts` sweep passes 9,938 ordinary tests plus
17 expected raster failures. The subsequent plain-title regression and updated
width assertions pass in the 20-test focused wrapping/PowerPoint run. Browser
runs pass 212 checks: 176 native-style/title/spacing/wrapping checks and 36
built-in-font/inheritance checks across six bindings. Core/UI typechecks,
builds and package import/registration checks pass. Both wrapping JSON fixtures
are absent from the published chart output. Manual title layout, autofit,
unverified scripts/fonts, advanced text cases and full Excel parity remain open.

### Manual title placement and layout preservation (8 October 2026)

Regenerating a chart after a type edit previously dropped its manual layout
and overlay properties. XLSX now preserves the entire imported title,
plot-area and legend layout/overlay subtrees, including unknown extensions.
Shared chart formatting exposes typed layout values. The parser is extracted
from PowerPoint core and rectangle geometry from PowerPoint UI into one core
chart module; the PowerPoint adapters continue to consume those shared helpers.
The source structure is documented in Microsoft's
[Open XML manual-layout reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.manuallayout).

Imported title positions now reach the SVG painter. Overlay titles do not
reserve the ordinary title band, and moved/overlay titles paint above the
plot. Seven independently reopened Excel references cover automatic and
automatic-overlay titles, left/right movement, a moved overlay, a wrapped
title and a smaller chart. The owned recorder retains title/chart/plot
rectangles and source parts in `excel-chart-title-layout.json`. Core tests
verify source layout/overlay preservation, title positioning and overlay
reservation before/after regeneration. The fixture's drawing-canvas coordinates
are four points beyond the reported COM title Left/Top coordinates; native
PNG inspection confirms that mapping for this corpus.

Browser checks exposed a CSS border that shrank the chart SVG viewport by
two pixels and scaled its coordinate system. The frame now uses an outline,
keeping its declared SVG dimensions. This fixes the actual rendering issue
without relaxing the positioning comparison. The native workbook's two-cell
anchor still produces a 393-CSS-pixel frame instead of the normalized
400-pixel chart reference, so native grid/frame fidelity remains open.

Playwright MCP reviewed moved and overlay title imports through the actual
type-change dialog. Excel reopens the Line exports, resaves and reopens them:
Left=40 points, Top=60/120 points and IncludeInLayout=true/false are unchanged.
The native and browser images still differ in plot geometry, borders and
gridline appearance. Manual title dimensions, manual plot/legend rendering,
position authoring UI, autofit and complete chart parity remain unfinished.

The diagram/chart/XLSX and PowerPoint layout-parser sweep passes 8,581 ordinary
tests and 17 existing expected raster failures. Updated title-layout checks
pass all seven native references; the extracted PowerPoint UI integration
passes 16 checks. All 254 focused chart browser cases pass across six bindings,
including 42 new native-layout cases. Core/UI typechecks, builds and package
checks pass, and the native layout fixture is absent from published chart output.

### Imported manual legend rectangles (8 October 2026)

Imported legend layout and overlay metadata now reach the XLSX view model and
SVG painter. Rectangle geometry reuses `chart/resolveManualLayoutRect`, already
shared with PowerPoint. The existing automatic legend painter moves to a
dedicated XLSX layout module without changing its ordinary positioning path.
Moved non-overlay legends retain the automatic side reservation; overlays
release it. Manual and overlay legend markup paints after the plot.

Eight independently reopened Excel 16.0 build 20430 references cover right and
bottom automatic, moved and overlay legends plus wide/tall overlays. The
recording scripts capture legend position, dimensions, IncludeInLayout and
source parts. Native PNGs demonstrate horizontal wide-box flow and vertically
spaced tall-box flow. Core checks compare manual bounds against COM measurements
and source fractions, verify source layout/overlay preservation through type
regeneration, and bound horizontal label origins within 6 CSS pixels of these
native PNG measurements using captured Chromium Arial 16px widths. This bound
is specific to this corpus and does not establish exact label or raster parity.

Playwright MCP reviewed actual tall/wide imports and changed the tall chart to
Line through the shared dialog. Excel reopens that UI export, resaves and
reopens it with Left=40pt, Top=100pt, Width=90pt, Height=150pt, Arial 12pt and
IncludeInLayout=false unchanged. The actual native workbook still paints into
a 393px-high frame rather than the normalized 400px reference. Native/browser
images differ in chart geometry, gridlines, background layering and text.

The broad diagram/chart/XLSX and PowerPoint layout-parser sweep passes 8,591
ordinary tests with 17 existing expected raster failures. The updated eight
legend core checks also pass. Core/UI types, builds and package smoke checks
pass. Focused browser runs pass 302 unique cases across six bindings, including
48 new legend cases. The final horizontal-flow correction passes the eight
native core cases and all 48 legend browser cases again. Automatic legend
geometry, arbitrary multi-row/column packing, entry text
overrides, line/marker swatches, position-authoring UI and complete chart parity
remain open.

### Gradient raster gap follow-up (8 October 2026)

The core linear comparison now evaluates each native sample at its pixel
centre, as the browser repro already did through `getImageData`. With that
geometry the wide 135-degree coincident-stop capture explains itself: Excel
paints two stops saved at one position as a one-step ramp that starts at the
shared position. Pixels projecting to exactly 0.5 keep the earlier red stop
(both square diagonals), and pixels 0.000469 past it carry 30-31/255 of the
later white stop, which bounds the ramp to 1/261..1/251 of the gradient
vector. `chart/gradient-coincident-stops.ts` opens coincident linear stops by
1/256 (never past the next stop or the end), so every Office chart painter
(XLSX, PowerPoint charts, galleries) shares it. All 20 coincident captures stay
within 0.6 levels; the previous hard SVG step missed the wide boundary by
about 225 levels. Only linear fills were measured; radial and path stops keep
their offsets.

Sixteen translucent three-stop captures remain expected failures. Against the
straight-alpha SVG stops they miss by 3.22 to 3.67 premultiplied levels (limit
3), with alpha within 0.76; angles 180 and 270 peak at 2.40 to 2.74 and pass.
In the 0-56% red-to-green segment native red sits about 2.3 levels below the
model and green about 0.9 above: colour leans about 0.01 of the segment toward
the more opaque stop (premultiplied interpolation would lean about 0.08), plus
roughly one level of PARGB-to-ARGB export truncation (opaque green reads 254).
The passing two-stop translucent profile shows the same lean at under two
levels. An alpha weight of `a^0.05` would pass every case but is an
unexplained fit, so it is not applied. Separating the colour weighting from
export truncation needs native captures of a two-stop red-to-green fill with
unequal alphas and of the same fill with its alphas swapped.

The two series repros remain browser-only gaps; the core models pass them at
pixel centres and corners. The translucent top-right circle at 391,479 is
2.18 levels off in core (native green premultiplied 133.2, model 135.4) and
3.04 in the browser, so browser gradient quantization adds about 0.9. The
centred shape at 302,376 is 23-24 in core against native 23, but the ramp
there is about three levels per pixel, so a sub-pixel registration difference
of the per-mark pattern accounts for the 3-level browser error. Browser
per-pixel values with the rendered mark bounds, or native captures at a larger
chart size, are needed to close them. The browser specs were updated for the
coincident fix but were not executed in this change.
