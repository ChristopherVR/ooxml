# Excel parity review and implementation order

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

| Priority       | Area                    | Evidence and gap                                                                                                                                                                                                                                     | Next acceptance target                                                                                                                                                       |
| -------------- | ----------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| P0             | Everyday editing        | `edit/clipboard.ts` clipped every copied rectangle to used cells; `grid/grid-commands.ts` routed Fill Down/Right through series inference. Fixed in this slice.                                                                                      | Trailing blanks overwrite the full finite destination; date and weekday fills repeat content, retain formatting and translate formulas.                                      |
| P0             | Save fidelity           | The reader/writer retain unsupported XML and parts, and patch some carried references. This does not prove preservation across every edit. Default clipboard payloads also do not cover all Excel metadata such as conditional formats and drawings. | Excel-produced fixtures for metadata, structural edits, tables, drawings, external links and unsupported parts; compare package contents and reopen in Excel without repair. |
| P1             | Clipboard and selection | Compatible selection paste now repeats copied blocks, with size checks and one undo step. Whole-row/column copies are intentionally bounded to used content. Clipboard payloads are dense arrays.                                                    | Skip blanks, full metadata semantics and bounded large-selection performance.                                                                                                |
| P1             | Formula compatibility   | A substantial function catalogue and recorded Excel corpus exist; missing functions and cached external references remain documented.                                                                                                                | Expand real-Excel differential cases for coercion, errors, arrays, names, tables, dates and structural edits before treating any function as complete.                       |
| P1             | Grid and print fidelity | The core supplies grid metrics and cell views; printing uses a browser-rendered range. Page Layout and page-break preview are absent.                                                                                                                | Reference workbooks and screenshots for widths, fonts, wrapping, merges, panes and number formats; paginated print with repeating titles and scaling.                        |
| P2             | Charts and drawings     | Live chart rendering and limited chart edits exist; many drawing types remain placeholders. SmartArt primarily uses cached drawings.                                                                                                                 | Shared chart/DrawingML support for axes, labels, series formatting and geometry, with read/edit/save fixtures.                                                               |
| P2             | Pivot/data features     | Pivot caches/tables are retained without pivot refresh/rendering; slicers, timelines and sparklines are not drawn. Connections are cached.                                                                                                           | Implement bounded pivot aggregation and views, then filtering controls and refresh, with explicit unsupported-feature reporting.                                             |
| P2             | Collaboration           | The shared Yjs infrastructure exists; spreadsheet document mapping and product integration are missing.                                                                                                                                              | XLSX adapter with atomic edits, origins, shared undo rules, selections, reconnect and conflict tests.                                                                        |
| Separate scope | Excel platform features | VBA is carried but never executed; add-ins and Power Query are not implemented. XLSB/ODS are unsupported.                                                                                                                                            | Define platform and file-format requirements explicitly before claiming whole-product 1:1 parity.                                                                            |

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
