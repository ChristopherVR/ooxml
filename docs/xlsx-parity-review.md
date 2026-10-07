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
| P1             | Clipboard and selection | The UI pastes at the selection's start, without repeating a source over a larger destination selection. Whole-row/column copies are intentionally bounded to used content. Clipboard payloads are dense arrays.                             | Tiled paste, incompatible-size handling, skip blanks, full metadata semantics and bounded large-selection performance.                                                       |
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

## Implemented first slice

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
- Native Excel COM acceptance and a new real-Excel recording were not run for this slice.
