# Comparing workbooks (`ooxml-core/xlsx`)

`compareWorkbooks(before, after, options?)` reports what changed between two versions of a
workbook: sheets, defined names, the structure of each sheet and its cells. It is DOM-free, does
not modify either workbook and returns a structured, deterministic result; `comparisonText`
turns that result into a plain-text report for a terminal or a code review. The code is in
`src/core/xlsx/compare/`.

A byte diff of two `.xlsx` files says nothing useful to a person, and the edit history of a
session lives only in memory. This compares the two workbook models instead, in the categories
Excel's Spreadsheet Compare uses.

```ts
import { compareWorkbooks, comparisonText, loadXlsx } from 'ooxml-core/xlsx';

const result = compareWorkbooks(await loadXlsx(v1), await loadXlsx(v2));
if (result.count) console.log(comparisonText(result));
```

```text
Workbook
  Sheet "Data" renamed to "Inputs"
  Name Rate changed: =Data!$A$1 -> =Inputs!$A$2
Sheet "Inputs" (was "Data")
  Merge added D1:E1
  A1 entered value: 1 -> "one"
  B1 formula: =A1*2 -> =A1*3
  C1 calculated value: 3 -> 4 (=A1+2)
  A2 format (number format): "General" -> "0%"
  A4:B5 entered value: 4 cells
Not compared: rowAndColumnLayout, rowAndColumnFormats, comments, ...
```

## What is compared

| Category            | Reported when                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| ------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `enteredValue`      | A constant changed, was typed into an empty cell or was cleared. `1` and `"1"` differ.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| `formula`           | A formula was added, removed or edited, or the range of an array formula changed. Formula text is compared exactly. A what-if data table result shows its `TABLE()` formula, so a new input cell is a formula change. A formula switched between dynamic-array and legacy (implicit intersection) semantics is a formula change where that changes its result.                                                                                                                                                                                                                                                                                      |
| `calculatedValue`   | The cached result of an unchanged formula changed, or a spilled, array formula or what-if data table result did.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `format`            | The effective format differs, with the attributes that changed: `numberFormat`, `font`, `fill`, `border`, `alignment`, `protection`, and `richText` when the runs of a rich-text string differ on the same stretch of text as written, against the cell font and as painted. So unchanged runs under a new cell font, a run that repeats the cell font and runs rewritten to paint the same are not run changes, while `bold: false` in a bold cell is; a string typed into or cleared from a cell is a value change only; neighbouring runs that look the same are joined, and when the text was edited too the sequence of run looks is compared. |
| `sheet`             | A sheet was added, removed, renamed, moved or shown/hidden.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| `name`              | A defined name was added, removed or changed (formula, hidden, comment). Sheet-scoped names follow their sheet.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| `merge`             | A merged range was added or removed.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `table`             | A table (matched by name) was added or removed, or its range, columns, header or totals row or style changed.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| `dataValidation`    | A validation rule was added or removed. A changed rule shows as removed and added; spelled-out defaults (`between`, `stop`, flags that are off, empty texts) are not a change.                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `conditionalFormat` | A conditional format rule was added or removed, or moved above or below other rules (`conditionalFormatReordered`). Each rule is compared (a rule with formulas together with the first cell of its format's first range, where its relative references are anchored), and reported as a one-rule format, with the ranges of its format, so how a file groups rules into formats does not matter. Rules are matched by content and priorities by their order, so renumbering priorities, storing rules in another order, adding a rule or spelling out a default (`stopIfTrue: false`, `showValue: true`, an empty `font`) is not a change.         |

Cell formats are compared by their resolved attributes, never by style index: two files (or one file
before and after a save) can number the same formats differently. Defaults count as absent (`bold:
false`, `horizontal: 'general'`, `locked: true`), RGB colours are compared with their alpha
(`FF0000` equals `FFFF0000`) and `font.family` is ignored. Colours are compared as stored (a theme
slot is not resolved to RGB). A font without a name or size takes the default format's, as the grid
paints it, so changing style 0 shows in the cells that rely on it. The named cell style a format
derives from (`cellStyleName`), the quote prefix and the pivot button flag are not compared. A cell
present in only one workbook is compared with the format of the other workbook's row, or else
column, at that address. Conditional format rules compare their colours the same way, but keep their
other settings as stored: in a differential format `bold: false` switches bold off, so it is not a
default.

Everything the comparison does not look at is listed in `result.notCompared`, together with the
categories turned off in the options, so a report can say so: row heights, column widths and
hidden rows or columns, the formats of whole rows and columns (used only as the format of a cell
stored in the other workbook), comments, the Excel 2010 extension details of conditional formats
(`x14` data bar settings), hyperlinks, drawings and charts, pivot tables, sparklines,
slicers, auto-filters, sheet views, page setup (including print areas), sheet protection (cell
locking is compared, as a format), document properties, the theme, workbook settings (date
system, calculation mode, structure protection) and parts the model keeps without modelling
(VBA, custom XML).

## Matching

Cells are matched by address, as Spreadsheet Compare does. **Inserting or deleting a row or
column therefore shows as a change to every cell below or to the right of it.** Detecting
inserted rows and columns is not implemented.

Sheets are paired by `sheetMatch`:

- `auto` (default): by name (case-insensitive, as Excel compares sheet names), then by `sheetId`
  among the sheets left over, which finds renamed sheets. A producer that renumbers sheet ids
  on save still pairs by name.
- `name`: by name only; a renamed sheet shows as removed and added.
- `sheetId`: by `sheetId` only, which survives renames within one file's history.

Every pair says how it was matched (`matchedBy`). Added and removed sheets are reported in
`result.workbook`; their cells are not listed. A sheet counts as moved when it left the longest
run of sheets that kept their relative order, so adding or removing a sheet does not move the
others.

## Result

`result.workbook` holds sheet and name differences. `result.sheets` has one entry per paired
sheet, in the after workbook's order, with `structure` (merges, tables, validation, conditional
formats), `cells` in row-major order and `ranges`: the cells grouped into rectangles of
neighbouring cells with the same kind of difference (and, for formats, the same attributes), so
a block that changed reads as one line. Each cell difference carries both sides' value, formula
and array range, and for formats the effective `CellStyle` of each side.

## Options

- `categories`: compare only these (default all).
- `maxDifferences`: stop after this many differences and set `truncated`. The cells are walked in
  row-major order, so the kept differences are the same on every run.
- `sheetMatch`: see above.

`comparisonText(result, { cells: true })` prints one line per cell instead of one per range.

## Cost

The comparison visits each stored cell of both workbooks once, sorting each sheet's row numbers
and each row's column numbers; it never builds a dense grid, and formats are reduced to
canonical keys once per style, not once per cell. On a shared 32-core Linux machine under load,
a pair of 1,000,000-cell workbooks (10,000 rows by 100 columns) compares in 0.3 to 0.5 s when
identical, 0.6 to 0.8 s when every cell differs and about 1.2 s against an empty workbook; with
`maxDifferences: 1000` it stops after 1 to 2 ms.
