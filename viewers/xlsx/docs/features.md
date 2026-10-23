# Features and limitations

This is an early implementation. It is **not Microsoft Excel parity** and saving is not lossless for features the model does not cover. This page lists what the editor is built to handle and what it does not; when a workbook contains something unsupported, the editor reports it (`workbook-warning`) instead of hiding it.

The [current parity status and evidence](https://github.com/ChristopherVR/ooxml/blob/main/docs/xlsx-parity-status.md)
summarize implemented work, native Excel comparisons and the remaining gaps.

::: warning Early release
The workbook engine in `ooxml-core` and the `<xlsx-editor>` component are young. The browser tests (`e2e/xlsx/*.spec.ts`) cover opening `.xlsx`, `.xls` and `.csv`, typing values and formulas, recalculation, ribbon formatting, row insertion and deletion, sheet tabs, undo and redo, Find and Replace, Format Cells, chart selection, save and reopen, read-only mode, locales and all six framework bindings. Everything else below is implemented but less exercised; expect rough edges.
:::

## File formats

| Format             | Open | Save       | Notes                                                                                                                                                                                                                        |
| ------------------ | ---- | ---------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `.xlsx`, `.xltx`   | yes  | yes        | The main format.                                                                                                                                                                                                             |
| `.xlsm`            | yes  | yes        | The VBA project is carried through unchanged; macros never run.                                                                                                                                                              |
| `.xls` (97-2003)   | yes  | as `.xlsx` | Read through the shared ole2 codecs: values, formulas where they decode, basic formatting, merges, column widths, row heights, frozen panes.                                                                                 |
| `.csv`             | yes  | yes        | RFC 4180, delimiter auto-detected; fields beginning with = stay text on import. Export writes the active sheet's displayed values and protects against formula injection.                                                    |
| Password-protected | yes  | yes        | Encrypted .xlsx/.xlsm files prompt for a password. File > Info > Encrypt Workbook sets a password for subsequent Excel saves. Opening an encrypted file does not retain its password for saving; set it again before saving. |
| `.xlsb`, `.ods`    | no   | no         |                                                                                                                                                                                                                              |

## Grid and formatting

- Cell values, shared strings with rich text runs, booleans, errors and dates (1900 and 1904 date systems).
- Fonts, fills (solid, pattern, gradient approximated), borders, alignment, wrap, indent, rotation.
- Number formats: sections, conditions, colours, dates and times, fractions, scientific, text. A General number that does not fit its column drops decimals or switches to scientific notation, and other numbers show `###`, measured with the browser's fonts.
- Row heights follow wrapped or enlarged text after an edit unless the row has a custom height.
- Merged cells, column widths and row heights, hidden rows and columns, frozen panes, zoom, gridlines.
- Conditional formatting: cell rules, text rules, "A Date Occurring" (today, last 7 days, this month, ...), top and bottom, above average, duplicates, colour scales, data bars, icon sets, with a Rules Manager that edits, reorders and deletes rules.
- Data validation, with the list drop-down offered in the grid.
- Comments (shown, edited; threaded replies display as notes), hyperlinks.

## Formulas

The calculation engine lives in `ooxml-core` and recalculates dependents after each edit. Its catalogue holds 497 functions (counted from `FUNCTION_CATALOG` on 2026-10-15) across math and trigonometry, statistics (including the inverse distributions and the `LINEST`, `LOGEST`, `TREND` and `GROWTH` family), logic, text (including `REGEXTEST`, `REGEXEXTRACT` and `REGEXREPLACE`), dates, lookup (including `XLOOKUP` and dynamic arrays such as `FILTER`, `SORT`, `UNIQUE` and `SEQUENCE`), information, financial (including the bond, coupon, odd-period (`ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`), discount-security and Treasury bill functions, `AMORDEGRC`, `VDB` and `XIRR`), engineering (including `COMPLEX`, the `IM` functions, `CONVERT` and the `BESSEL` functions) and database (`DSUM`, `DCOUNT`, `DGET` and the rest). Insert Function lists them by category.

Not every Excel function exists: an unknown one shows `#NAME?`. These are not implemented: `FORECAST.ETS`, `ASC`, `DBCS`, `JIS`, `PHONETIC`, `INFO` and `WEBSERVICE`. `GETPIVOTDATA` reads the pivots of this editor (worksheet-range pivots) from their cached records and returns `#REF!` for anything it cannot resolve. `GROUPBY` and `PIVOTBY` are a subset: one aggregating function (a function name such as `SUM` or `PERCENTOF`, or a `LAMBDA`), grand totals only (total depth 0, 1 or -1, no subtotals), field headers 0, 1 or 3, no function arrays, and flat group output with the row labels repeated; `PIVOTBY` takes one column field. `TRIMRANGE` returns the trimmed values as an array rather than a reference. `BESSELJ(1.9,2)` returns the exact 0.3299257277, which differs from the 0.329925829 Microsoft documents (Excel itself is about 1e-7 off there). The odd-period security functions use the quasi-coupon method and reproduce Microsoft's examples; long odd first periods are not checked against Excel. External workbook references are kept but not resolved.

Circular references are reported and left at 0 unless iterative calculation is on. Turn it on in File > Options > Formulas (Enable iterative calculation, with Maximum iterations and Maximum change); the setting is saved with the workbook (`calcPr iterate`, `iterateCount`, `iterateDelta`). A circular group is then recalculated up to the maximum iterations and stops once no cell changes by more than the maximum change. Turning it on or off is not an undo step.

Calculation can be switched between Automatic and Manual (Formulas > Calculation Options, or File > Options); the mode is saved with the workbook. Calculate Now (F9) and Calculate Sheet (Shift+F9) recalculate on demand. Show Formulas (Ctrl+\`) displays formula text instead of results.

## Editing

Typing into cells (values and formulas), the formula bar and name box, Ctrl+; and Ctrl+Shift+; for the current date and time, fill, copy and paste (with the system clipboard as text and HTML), insert and delete rows, columns and cells, sorting, filtering, remove duplicates, find and replace, sheet management (add, rename, move, hide, colour; new sheets are named in the interface language, for example `Tabelle2` in German), the built-in cell styles, tables (style options, header and total rows, resize, convert to range), row and column outline groups, page setup and print options (gridlines, headings, centring), undo and redo with meaningful step names. Formula references follow structural edits.

What-If Analysis (Data > Forecast > What-If Analysis) has three tools. **Goal Seek** changes one input cell until a formula cell reaches a value, iterating the recalculation engine (up to the workbook's Maximum iterations, default 100, until within Maximum change, default 0.001); the Goal Seek Status dialog keeps the result with OK and restores it with Cancel, and the whole run is one undo step. **Data Table** turns a selection into a one- or two-variable `{=TABLE(row input, column input)}` table: the results are recomputed by substituting the header values into the input cells (the real inputs are never changed) after every calculation, the table is read and saved as `<f t="dataTable">` and editing part of it is refused ("Cannot change part of a data table."; selecting the whole table and clearing it removes it). **Scenario Manager** adds, edits, deletes and shows scenarios (up to 32 changing cells each, with comment and protection flags) saved as the sheet's `<scenarios>` element. Not implemented: Forecast Sheet (it stays disabled and says why), Scenario Manager's Summary report and Merge, Goal Seek's Step and Pause, data tables on sheets other than the input cells' own, and moving a data table with row or column inserts and deletes.

Sheet and workbook-structure protection can carry a password. It stops accidental edits; it is not security. The editor reads and checks both Excel's legacy 16-bit hash and its newer SHA-512 (also SHA-1, SHA-256 and SHA-384) hash, so Unprotect Sheet and Unprotect Workbook ask for the password whichever hash the file carries. A password set in the editor is written as the legacy hash. A protection whose hash uses a digest the core cannot compute is never unlocked by a password.

Pictures and charts are selected by clicking them: they move and resize with the mouse, nudge with the arrow keys and are removed with Delete; a selected chart shows the Chart Design tab.

## Charts and SmartArt

Common chart families (column, bar, line, area, pie, doughnut, scatter, bubble, stock
(high-low-close and open-high-low-close) and radar), bar/line combinations with a secondary
axis, and waterfall, histogram, treemap and sunburst charts (read-only, from their ChartEx parts)
render as SVG from live values. Volume-stock combinations and surface charts show a labelled frame.

**Insert > Chart** opens the shared chart picker, the same one Word and PowerPoint use. Every chart
is edited over its parsed chart model (`ChartSpace`, the shared model in `ooxml-core/chart`):

- **Chart Design tab:** Chart Styles, Change Colors and Chart Layouts (Quick Layout 1 to 11 for
  column, bar, line and area charts) galleries, Chart Title and Legend, Type, Series Options, and
  Delete. The galleries are the shared ones, so they look the same in every product.
- **Format pane:** a docked panel with the shared chart Format pane. It edits the border and
  outline of the chart, axes (bounds, units, scale, crossing, tick marks, labels, number format),
  gridlines, data labels, the legend and series options. Series fills and the chart area fill
  open the same panel from their Series Options commands.
- A type change keeps the formatting: titles, legend, axes and fills carry over to the new type.
- A chart part whose content the model cannot carry (for example detail below the top level, such
  as a surface chart's band formats) is never rewritten from the model. Its chart commands and
  Format pane are disabled, a workbook warning reports it, and it is saved from its source bytes,
  with only moved references patched in.

Not offered: effects, 3-D format, text options, pattern and picture fills, trendlines, error bars,
markers and axis titles in the Format pane (no shared intent writes them yet); chart styles are
curated recolour presets, so style definitions (`styleN.xml`) cannot be authored; surface charts are
not drawn. Imported chart styles, supported shadows and gradients, axis visibility, inherited fonts,
mixed-format title runs, automatic title wrapping and title paragraph spacing have rendering and
preservation support. Imported manual title positions and overlays render, and title, plot and
legend layout XML is kept through chart type changes. Exact chart text measurement, wrapping,
placement and raster fidelity remain incomplete, as do many axis, label and effect controls.
Native Excel checks cover specific edit and export paths (`src/core/scripts/xlsx-excel-acceptance.ps1`
opens the saved workbooks in desktop Excel); they are not a claim of native layout parity.

**Insert > SmartArt** opens the shared layout picker (35 layouts). Selecting a diagram shows the
SmartArt Design tab (Add Shape after, before, above and below, Promote, Demote, Move Up, Move Down,
Text Pane and Change Layout, the same command list as Word). The text pane opens beside the grid,
and double-clicking a diagram opens it too. Each node edit, and each keystroke in the text pane, is
one undoable step. A diagram that is inserted, edited or saved without a cached drawing is drawn
with an approximate layout, badged as not Excel-identical; Excel lays it out again when it opens the
file, and saved files open in desktop Excel without repair
(`src/core/scripts/xlsx/smartart-excel-acceptance.ps1`). SmartArt Styles, Change Colors,
per-shape formatting and rich text in the text pane are not implemented.

## Sparklines

Excel 2010 sparklines are drawn in their host cells, behind the cell's text:
line (with markers and high, low, first, last and negative points), column and
win/loss. The drawing follows the group's settings for empty cells (gap, zero or
connected), the vertical axis (each sparkline, the whole group or custom limits),
the horizontal axis at zero, the date axis, right-to-left plotting, hidden cells
and theme or RGB colours. A sparkline repaints when a cell in its data range
changes and resizes with its cell. The sparkline XML is saved exactly as loaded,
and inserted or deleted rows and columns move it with its cells.

Insert > Sparklines creates line, column and win/loss groups from a data range
and a location range. A vertical location range maps data rows to individual
sparklines; a horizontal location range maps data columns. Select a host cell
and choose Sparkline Settings to change the entire group's type, markers,
series and negative colors, line weight, empty and hidden cell handling, axis
visibility, direction and individual, shared or custom vertical limits.
Clear Selected Sparklines removes the selected hosts; Clear Sparkline Groups
removes their entire groups. These changes support undo and save/load.

Style galleries and grouping/ungrouping are not available. The date axis can be switched on and off in Sparkline Settings. Exact Excel pixel output (marker size, column gap and
line anti-aliasing) is not matched, and non-numeric cells in a data range are
treated as empty.

## Not supported (preserved where possible)

| Feature                             | Behaviour                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| ----------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Pivot tables and pivot caches       | Worksheet-range PivotTables (up to two row fields, one column field, a report filter and several value fields) are created, shown, refreshed and saved. A numeric row or column field can be grouped by number ranges (Group Field, Ungroup) and a date field by seconds up to years (Group Dates); hidden items and a timeline date range filter the report. Other pivots (external or OLAP sources, calculated fields, more than two row fields, custom subtotals, label or value filters, shared caches) are kept on save but not shown or refreshed.                                                                                                                                                                                                      |
| Slicers, timelines                  | Table and PivotTable slicers are drawn (caption header, Multi-Select and Clear Filter buttons, item buttons in columns with selected, unselected and no-data states, the 14 built-in styles) and filter the table or PivotTable on click, Ctrl, Shift and Multi-Select. Insert > Slicer and Insert > Timeline add them; the Slicer and Timeline tabs set caption, style, columns, size, sort order and time level. Timelines on a PivotTable date field show years, quarters, months or days and filter by click, Shift+click or drag. Slicers on OLAP or other unsupported sources, and slicers connected to several PivotTables, are kept on save but drawn as placeholders. Style colours follow Excel's styles but are not measured from a native render. |
| Sparkline authoring                 | Creation, group settings and clearing are supported (see above); style galleries and grouping/ungrouping are not implemented.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| Advanced charts and chart options   | Common families and supported formatting are described above. Advanced families, complete axes/labels/effects authoring and exact native layout remain incomplete.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
| Shapes, SmartArt, form controls     | Pictures display. SmartArt is drawn from its cached drawing (or an approximate layout) and its nodes, text and layout can be edited; no SmartArt Styles or Change Colors. Other drawings show as placeholders.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| Macros (VBA), add-ins, Power Query  | Never executed; the VBA project is carried through.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| External links and data connections | Kept; values are the cached ones.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| Real-time collaboration             | Cell values, formulas, formats, sheets, row and column sizes, merges and names are shared over `ooxml-core/collab` (File > Share or the `collaboration` property); remote selections show in the grid and undo is per person. Charts, drawings, comments, conditional formats, validation, tables and page setup stay local; the host owns the server and saving. See [collaboration](/collaboration).                                                                                                                                                                                                                                                                                                                                                        |
| Printing                            | The browser's print of the used range or print area. Page Layout view (pages, rulers, editable header and footer bands with &P &N &D &T &F &A codes, cell editing; no copy and paste in it) and Page Break Preview (print-area shading, page numbers, dashed automatic and solid manual break lines, drag a break or the print-area edge, with undo) are chosen from the status bar or View > Workbook Views and saved per sheet (`sheetView/@view`). Breaks cannot be dragged while Fit to pages is on. Custom Views stays disabled.                                                                                                                                                                                                                         |
| Document properties                 | Core, company, manager and custom properties are editable in File > Info with undo. Unknown custom-property types are kept and displayed read-only.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |

The core's round-trip tests and Excel acceptance checks live in the `ooxml` repository. Report a workbook that renders or saves wrongly as an issue with the file attached if you can share it.

Digital signatures are reported when opening a workbook and removed on save. The editor does not sign files. Hyperlinks follow the shared core policy; only web and e-mail links open externally.
