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

## Evidence required for parity

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
