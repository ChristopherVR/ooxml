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

The calculation engine lives in `ooxml-core` and recalculates dependents after each edit. Its catalogue holds about 480 functions across math and trigonometry, statistics (including the inverse distributions and the `LINEST`, `LOGEST`, `TREND` and `GROWTH` family), logic, text (including `REGEXTEST`, `REGEXEXTRACT` and `REGEXREPLACE`), dates, lookup (including `XLOOKUP` and dynamic arrays such as `FILTER`, `SORT`, `UNIQUE` and `SEQUENCE`), information, financial (including the bond, coupon, discount-security and Treasury bill functions, `VDB` and `XIRR`), engineering (including `COMPLEX`, the `IM` functions and `CONVERT`) and database (`DSUM`, `DCOUNT`, `DGET` and the rest). Insert Function lists them by category.

Not every Excel function exists: an unknown one shows `#NAME?`. These are not implemented: `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `AMORDEGRC`, `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`, `FORECAST.ETS`, `ASC`, `DBCS`, `JIS`, `PHONETIC`, `BAHTTEXT`, `INFO`, `GETPIVOTDATA`, `ISOMITTED` and `WEBSERVICE`. External workbook references are kept but not resolved.

Circular references are reported and left at 0 unless iterative calculation is on. Turn it on in File > Options > Formulas (Enable iterative calculation, with Maximum iterations and Maximum change); the setting is saved with the workbook (`calcPr iterate`, `iterateCount`, `iterateDelta`). A circular group is then recalculated up to the maximum iterations and stops once no cell changes by more than the maximum change. Turning it on or off is not an undo step.

Calculation can be switched between Automatic and Manual (Formulas > Calculation Options, or File > Options); the mode is saved with the workbook. Calculate Now (F9) and Calculate Sheet (Shift+F9) recalculate on demand. Show Formulas (Ctrl+\`) displays formula text instead of results.

## Editing

Typing into cells (values and formulas), the formula bar and name box, Ctrl+; and Ctrl+Shift+; for the current date and time, fill, copy and paste (with the system clipboard as text and HTML), insert and delete rows, columns and cells, sorting, filtering, remove duplicates, find and replace, sheet management (add, rename, move, hide, colour; new sheets are named in the interface language, for example `Tabelle2` in German), the built-in cell styles, tables (style options, header and total rows, resize, convert to range), row and column outline groups, page setup and print options (gridlines, headings, centring), undo and redo with meaningful step names. Formula references follow structural edits.

Sheet and workbook-structure protection can carry a password. It stops accidental edits; it is not security. The editor reads and checks both Excel's legacy 16-bit hash and its newer SHA-512 (also SHA-1, SHA-256 and SHA-384) hash, so Unprotect Sheet and Unprotect Workbook ask for the password whichever hash the file carries. A password set in the editor is written as the legacy hash. A protection whose hash uses a digest the core cannot compute is never unlocked by a password.

Pictures and charts are selected by clicking them: they move and resize with the mouse, nudge with the arrow keys and are removed with Delete; a selected chart shows the Chart Design tab.

## Charts and SmartArt

Common chart families (column, bar, line, area, pie, doughnut, scatter and radar)
render as SVG from live values. Chart insertion, type/grouping, title, legend and
color changes use the shared core and Office gallery controls. Docked formatting
panes provide series fills, supported bar spacing, and fills for chart area,
plot area, title and legend. Supported gradient types, stops, transparency,
brightness, direction previews and native presets share the same fill controls.

Imported chart styles, supported shadows/gradients, axis visibility, inherited
fonts, mixed-format title runs, automatic title wrapping and point/percentage title paragraph spacing
have rendering and preservation support. Imported manual title positions and
title overlays render; title/plot/legend layout XML is preserved through
chart type changes. Manual dimensions, complete plot/legend placement and
position-authoring controls remain incomplete.
Native Excel comparisons verify specific edit/export paths. Exact chart text
measurement, wrapping, placement and raster fidelity remain incomplete, as do
many axes/labels/effects controls, advanced chart families and full native UI.

SmartArt uses the shared cached-drawing renderer with workbook theme fonts and
colors. It remains display-only: insertion, text editing, reflow and complete
layout/color/style galleries are not implemented. A missing cached drawing
produces a text placeholder.

## Not supported (preserved where possible)

| Feature                             | Behaviour                                                                                                                                                                         |
| ----------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Pivot tables and pivot caches       | Kept on save; not shown as pivots and not refreshed.                                                                                                                              |
| Slicers, timelines, sparklines      | Kept on save; not drawn.                                                                                                                                                          |
| Advanced charts and chart options   | Common families and supported formatting are described above. Advanced families, complete axes/labels/effects authoring and exact native layout remain incomplete.                |
| Shapes, SmartArt, form controls     | Pictures display. SmartArt uses the cached drawing through the shared renderer (display only); without a cached drawing, its text is listed. Other drawings show as placeholders. |
| Macros (VBA), add-ins, Power Query  | Never executed; the VBA project is carried through.                                                                                                                               |
| External links and data connections | Kept; values are the cached ones.                                                                                                                                                 |
| Real-time collaboration             | Not implemented yet, see [collaboration](/collaboration).                                                                                                                         |
| Printing                            | The browser's print of the used range or print area; no Page Layout view or page break preview.                                                                                   |
| Document properties                 | Core, company, manager and custom properties are editable in File > Info with undo. Unknown custom-property types are kept and displayed read-only.                               |

The core's round-trip tests and Excel acceptance checks live in the `ooxml` repository. Report a workbook that renders or saves wrongly as an issue with the file attached if you can share it.

Digital signatures are reported when opening a workbook and removed on save. The editor does not sign files. Hyperlinks follow the shared core policy; only web and e-mail links open externally.
