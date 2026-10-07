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

| Priority       | Area                    | Evidence and gap                                                                                                                                                                                                                            | Next acceptance target                                                                                                                                                       |
| -------------- | ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| P0             | Everyday editing        | `edit/clipboard.ts` clipped every copied rectangle to used cells; `grid/grid-commands.ts` routed Fill Down/Right through series inference. Fixed in this slice.                                                                             | Trailing blanks overwrite the full finite destination; date and weekday fills repeat content, retain formatting and translate formulas.                                      |
| P0             | Save fidelity           | The reader/writer retain unsupported XML and parts, and patch some carried references. This does not prove preservation across every edit. Default clipboard payloads also do not cover all Excel metadata such as validation and comments. | Excel-produced fixtures for metadata, structural edits, tables, drawings, external links and unsupported parts; compare package contents and reopen in Excel without repair. |
| P1             | Clipboard and selection | Compatible selection paste now repeats copied blocks, with size checks and one undo step. Whole-row/column copies are intentionally bounded to used content. Clipboard payloads are dense arrays.                                           | Skip blanks, full metadata semantics and bounded large-selection performance.                                                                                                |
| P1             | Formula compatibility   | A substantial function catalogue and recorded Excel corpus exist; missing functions and cached external references remain documented.                                                                                                       | Expand real-Excel differential cases for coercion, errors, arrays, names, tables, dates and structural edits before treating any function as complete.                       |
| P1             | Grid and print fidelity | The core supplies grid metrics and cell views; printing uses a browser-rendered range. Page Layout and page-break preview are absent.                                                                                                       | Reference workbooks and screenshots for widths, fonts, wrapping, merges, panes and number formats; paginated print with repeating titles and scaling.                        |
| P2             | Charts and drawings     | Live chart rendering and limited chart edits exist; many drawing types remain placeholders. SmartArt primarily uses cached drawings.                                                                                                        | Shared chart/DrawingML support for axes, labels, series formatting and geometry, with read/edit/save fixtures.                                                               |
| P2             | Pivot/data features     | Pivot caches/tables are retained without pivot refresh/rendering; slicers, timelines and sparklines are not drawn. Connections are cached.                                                                                                  | Implement bounded pivot aggregation and views, then filtering controls and refresh, with explicit unsupported-feature reporting.                                             |
| P2             | Collaboration           | The shared Yjs infrastructure exists; spreadsheet document mapping and product integration are missing.                                                                                                                                     | XLSX adapter with atomic edits, origins, shared undo rules, selections, reconnect and conflict tests.                                                                        |
| Separate scope | Excel platform features | VBA is carried but never executed; add-ins and Power Query are not implemented. XLSB/ODS are unsupported.                                                                                                                                   | Define platform and file-format requirements explicitly before claiming whole-product 1:1 parity.                                                                            |

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
Skip blanks and complete clipboard metadata support remain outstanding.

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
