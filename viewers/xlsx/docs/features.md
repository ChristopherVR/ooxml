# Features and limitations

This is an early implementation. It is **not Microsoft Excel parity** and saving is not lossless for features the model does not cover. This page lists what the editor is built to handle and what it does not; when a workbook contains something unsupported, the editor reports it (`workbook-warning`) instead of hiding it.

::: warning Early release
The workbook engine in `ooxml-core` and the `<xlsx-editor>` component are young. The browser tests (`tests/*.spec.ts`) cover opening `.xlsx`, `.xls` and `.csv`, typing values and formulas, recalculation, ribbon formatting, row insertion and deletion, sheet tabs, undo and redo, Find and Replace, Format Cells, chart selection, save and reopen, read-only mode, locales and all six framework bindings. Everything else below is implemented but less exercised; expect rough edges.
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

The calculation engine lives in `ooxml-core` and recalculates dependents after each edit. It implements a large set of functions (math, statistics, logic, text, dates, lookup including `XLOOKUP` and dynamic arrays such as `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, information and basic financial functions). Not every Excel function exists: an unknown one shows `#NAME?`. Circular references are reported and left at 0; iterative calculation is not supported. External workbook references are kept but not resolved.

Calculation can be switched between Automatic and Manual (Formulas > Calculation Options, or File > Options); the mode is saved with the workbook. Calculate Now (F9) and Calculate Sheet (Shift+F9) recalculate on demand. Show Formulas (Ctrl+\`) displays formula text instead of results.

## Editing

Typing into cells (values and formulas), the formula bar and name box, Ctrl+; and Ctrl+Shift+; for the current date and time, fill, copy and paste (with the system clipboard as text and HTML), insert and delete rows, columns and cells, sorting, filtering, remove duplicates, find and replace, sheet management (add, rename, move, hide, colour; new sheets are named in the interface language, for example `Tabelle2` in German), the built-in cell styles, tables (style options, header and total rows, resize, convert to range), row and column outline groups, page setup and print options (gridlines, headings, centring), undo and redo with meaningful step names. Formula references follow structural edits.

Sheet and workbook-structure protection can carry a password. It is Excel's legacy 16-bit password hash: it stops accidental edits, it is not security. A file whose sheet is protected only with Excel's newer SHA-512 hash (no legacy hash) is a known gap: the core does not expose that hash, so Unprotect Sheet removes such protection without asking for the password.

Pictures and charts are selected by clicking them: they move and resize with the mouse, nudge with the arrow keys and are removed with Delete; a selected chart shows the Chart Design tab.

## Not supported (preserved where possible)

| Feature                             | Behaviour                                                                                                                                                                                                              |
| ----------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Pivot tables and pivot caches       | Kept on save; not shown as pivots and not refreshed.                                                                                                                                                                   |
| Slicers, timelines, sparklines      | Kept on save; not drawn.                                                                                                                                                                                               |
| Charts                              | Common types drawn as SVG from live values. Type, title and legend can be changed; formatting the chart's own parts (axes, series colours, labels) is not. A loaded chart keeps its part on save, patched by the core. |
| Shapes, SmartArt, form controls     | Pictures display. SmartArt uses the cached drawing through the shared renderer (display only); without a cached drawing, its text is listed. Other drawings show as placeholders.                                      |
| Macros (VBA), add-ins, Power Query  | Never executed; the VBA project is carried through.                                                                                                                                                                    |
| External links and data connections | Kept; values are the cached ones.                                                                                                                                                                                      |
| Real-time collaboration             | Not implemented yet, see [collaboration](/collaboration).                                                                                                                                                              |
| Printing                            | The browser's print of the used range or print area; no Page Layout view or page break preview.                                                                                                                        |
| Document properties                 | Core, company, manager and custom properties are editable in File > Info with undo. Unknown custom-property types are kept and displayed read-only.                                                                    |

The core's round-trip tests and Excel acceptance checks live in the `ooxml` repository. Report a workbook that renders or saves wrongly as an issue with the file attached if you can share it.

Digital signatures are reported when opening a workbook and removed on save. The editor does not sign files. Hyperlinks follow the shared core policy; only web and e-mail links open externally.
