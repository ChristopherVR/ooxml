# Supported features and preservation guarantee

This page is the written list that [the stability policy](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)
requires of a Stable product. It says what the Excel editor models and edits, what it keeps
untouched through an edit and a save, and what it drops or refuses (with the message it shows).
[Features and limitations](./features.md) describes the user-facing behaviour, and the
[function list](./functions.md) the calculation engine. Nothing here claims that the editor matches
Microsoft Excel or that a saved file is identical to the source: the claim is narrower, below.

## The guarantee

Content the editor does not model (unknown elements, extension lists, parts, relationships,
custom XML and the like) survives opening a workbook, editing it and saving it, or the editor
reports that it cannot keep it. A silent loss is a bug and is fixed in a patch release.

"Survives" means the XML or bytes come back unchanged, in a valid position. It does not mean the
whole package is identical: the parts the editor regenerates from its model (listed below) are
rewritten, and Excel is free to rewrite them again on its next save.

## What is modelled and edited

These are read into the model, edited through the edit session, and written back from the model.
Attributes and children inside them that the model does not carry are kept where the table says
so.

| Area             | Modelled and written from the model                                                                                                                                                                                  |
| ---------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Cells            | Values, formulas (shared formulas are expanded), array and dynamic array formulas with their metadata, data table formulas, rich text strings, errors, dates in both date systems.                                   |
| Rows and columns | Heights, widths, hidden, outline levels, collapsed, row and column formats, thick top/bottom row hints.                                                                                                              |
| Formats          | Fonts, fills, borders, number formats, alignment, protection, named cell styles, differential formats.                                                                                                               |
| Sheet            | Merges, freeze panes, views, zoom, gridlines, tab colour, state (visible, hidden, very hidden), protection with password hashes, page setup, print options, print area, headers and footers.                         |
| Rules and lists  | Conditional formatting (the supported rule types), data validation, hyperlinks, auto filters, tables with their filters and totals row, defined names.                                                               |
| Annotations      | Notes, threaded comments with their persons.                                                                                                                                                                         |
| Drawings         | Charts (through the shared chart model; a chart the model cannot hold is kept as read), pictures, shapes it understands, SmartArt (cached drawing regenerated, source data kept), sparklines, slicers and timelines. |
| Analysis         | Worksheet-range PivotTables the editor supports, what-if data tables, Scenario Manager scenarios.                                                                                                                    |
| Workbook         | Sheet order and names, active sheet, defined names (with the attributes of macro and function names), calculation settings, workbook protection, date system, theme palette and fonts.                               |
| Properties       | Core, extended and custom document properties.                                                                                                                                                                       |

## What is kept unchanged through an edit and a save

The writer copies a part of the source package byte for byte (with its relationships and every
part they reach) when nothing in the model regenerates it. The source for this is
`PackageWriter.carry` in `src/core/xlsx/write/package-writer.ts`, driven by `saveXlsx`
(`src/core/xlsx/write/save.ts`) and `writeWorksheet` (`src/core/xlsx/write/worksheet.ts`).

| Kept                                                    | How                                                                                                                                                                                                                                                                                                                                  |
| ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| VBA project (`xl/vbaProject.bin`) of an `.xlsm`         | Carried with its relationship and the macro-enabled content type. Macros never run. Saving the workbook as `.xlsx` removes the project, because that format cannot hold it.                                                                                                                                                          |
| Custom XML (`customXml/*`) and their property parts     | Carried with the workbook relationship and the relationships between the parts.                                                                                                                                                                                                                                                      |
| External links (`xl/externalLinks/*`)                   | Carried with their (external) relationship targets; the `externalReferences` element stays in `workbook.xml`. The links are not updated (reported on load).                                                                                                                                                                          |
| Connections, query tables, and other workbook parts     | Any workbook relationship the writer does not regenerate is kept with its target part, whatever its type.                                                                                                                                                                                                                            |
| Printer settings, control properties, custom properties | Sheet relationships the writer does not regenerate are kept with their target parts (`printerSettings`, `ctrlProps`, `customProperty`, embedded objects, images).                                                                                                                                                                    |
| Worksheet elements the model does not represent         | `sheetCalcPr`, `phoneticPr`, `rowBreaks`, `colBreaks`, `customProperties`, `cellWatches`, `ignoredErrors`, `smartTags`, `oleObjects`, `controls`, `picture`, `webPublishItems`, `extLst` and the others of the schema are kept as XML and written back in schema order. Unknown extension entries inside `extLst` stay as they were. |
| Source attributes of modelled worksheet elements        | `sheetPr`, `sheetViews`, `sheetFormatPr`, `sheetProtection`, `autoFilter`, `pageMargins`, `pageSetup` and `headerFooter` keep the attributes the model does not carry when the modelled part did not change.                                                                                                                         |
| Workbook elements the model does not represent          | `fileVersion`, `fileSharing`, `functionGroups`, `externalReferences`, `oleSize`, `customWorkbookViews`, `smartTagPr`, `smartTagTypes`, `webPublishing`, `fileRecoveryPr`, `webPublishObjects`, `extLst`, every window of `bookViews` and the unmodelled attributes of `workbookPr`, `workbookView` and `calcPr`.                     |
| Chart sheets, dialog sheets, macro sheets               | Carried whole (shown as empty sheets and reported on load); their references follow sheet renames and deletions.                                                                                                                                                                                                                     |
| Unsupported PivotTable parts, pivot caches              | Carried with their relationships; cache definitions follow sheet renames.                                                                                                                                                                                                                                                            |
| Drawings and charts the editor did not touch            | The drawing, chart, style, colour and media parts of an unedited object are carried byte for byte.                                                                                                                                                                                                                                   |
| Theme                                                   | Carried unless the theme palette was changed.                                                                                                                                                                                                                                                                                        |

`src/core/xlsx/write/preservation.test.ts` checks the first rows of this table: it builds a package
with a VBA project, custom XML, an external link with an external target, a connections part,
printer settings, a control property part, an unknown relationship, unknown `extLst` entries in the
workbook and the sheet, every worksheet element listed above and a second workbook window, edits
one cell, saves, and asserts that each part is byte-identical, each relationship and content type is
present, and each XML fragment is back in schema order. A second save of the result is stable.
`save.test.ts` does the same with Excel's and openpyxl's own files.

## What is dropped, and why

| Dropped                                                                  | Why                                                                                                                                                             |
| ------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `xl/calcChain.xml`                                                       | Rebuilt by Excel; the workbook is saved with `fullCalcOnLoad` so every value is recalculated when it opens.                                                     |
| Digital signatures (the origin, signature parts and their relationships) | Any change invalidates them. Reported on load: "The workbook is digitally signed ... saving removes it". Excel drops them too.                                  |
| The VBA project when saved as `.xlsx`                                    | The format cannot hold it. The `.xlsm` format is kept as loaded.                                                                                                |
| `x15ac:absPath` (where the file was last saved) and `xr:revisionPtr`     | Authoring hints for Excel, regenerated by it. The absolute path would also leak the author's folder layout.                                                     |
| `xr:uid` attributes, row `spans`, `x14ac:dyDescent`                      | Revision identifiers and layout hints that Excel regenerates.                                                                                                   |
| Row and cell `ph` (show phonetic), cell `cm` without a metadata part     | Not modelled. Phonetic guides stay in the shared strings, only the flag that shows them is lost.                                                                |
| Shared formulas and inline strings as such                               | Written as the equivalent plain formulas and shared strings (the same values).                                                                                  |
| Conditional formatting rule types the editor does not support            | Reported on load: `Conditional formatting rule type "..." is not supported and was dropped.` The other rules of the sheet are kept.                             |
| A table or chart part that cannot be read                                | Reported on load ("A table part could not be read and was dropped", "Chart ... could not be read; it is kept but not shown": the chart part itself is carried). |
| A worksheet element outside the schema                                   | It cannot be placed in a valid position, so it is reported on load: `the element "..." is not part of the worksheet schema and is dropped on save.`             |

## What is refused or reported

Nothing is refused silently. The editor reports through `workbook.warnings` (shown by the hosts as a
notice) and by throwing:

- an encrypted package without a password asks for one (`load/detect.ts`); a wrong password is an error;
- an unsupported or damaged part is reported by name ("... is kept but not shown", "... could not be read");
- a workbook with VBA macros or external links says they are kept but not run or updated;
- a digitally signed workbook says the signature is removed on save;
- a chart that cannot be edited without losing detail is marked `editBlocked` and shown and saved unchanged;
- `.xlsb` and `.ods` are not opened; legacy `.xls` is converted to `.xlsx` on save, with a warning for each feature the binary reader cannot decode;
- editing part of a data table is refused ("Cannot change part of a data table.").

## Known limits of the guarantee

- A part kept byte for byte is not rewritten when the cells it refers to move. Charts, chart sheets, pivot caches and the other parts the writer knows are patched for sheet renames and deletions and for the moves of their ranges; a custom XML, connection or external part that mentions a cell range keeps the old text.
- Concurrent editing (the collaboration area) carries the modelled subset in the shared document; a peer without the original package cannot write the unmodelled parts back.
- The native check that Excel opens a saved file without repair is manual (`src/core/scripts/xlsx-excel-acceptance.ps1`); the date of the last passing run belongs in the project status, not in this list.
