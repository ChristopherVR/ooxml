# Supported features and the preservation guarantee

This page is the written feature list and preservation guarantee that the
[stability policy](../../../docs/stability.md) asks of a product before it is called Stable. It
covers Word (`docx-core` and the editor). It says what is modeled and editable, what is carried
through a save without being understood, and what is refused, and with which diagnostic. It does not
claim that Word and this editor lay a document out alike, or that a saved file is the same as the one
Word would write.

## How a save works

A file loaded with `loadDocx` keeps its original package. `save` opens that package again and changes
only the parts the edits reach:

- `word/document.xml` is parsed to a DOM and written back from it. The body is patched block by
  block: a paragraph, table or section the model did not change keeps its source XML, and the markup
  the model does not hold inside a changed one (unknown elements, extension attributes, `mc:`
  alternate content, bookmarks, tracked-change wrappers) stays where it was.
- A part is rewritten only when the model changed what it holds: a header, footer, footnote or
  endnote part, `word/comments.xml` (and `commentsExtended.xml`), `word/numbering.xml`, the style
  parts for edited paragraph style formatting, `docProps/core.xml`, `word/settings.xml` flags the
  editor owns, the chart and SmartArt parts that were edited, and the bound `customXml` item of an
  edited data-bound control.
- `[Content_Types].xml` and the relationship parts change only to register a part or relationship the
  save adds.
- Every other entry is written back as it was read: same name, same bytes. A save of a model that
  did not change returns the original bytes.

A package created by the editor (not loaded) has no source to preserve and is written from the
model with the built-in default styles.

## Kept verbatim

The round-trip test [`preservation.test.ts`](../../../src/core/docx/preservation.test.ts) loads a
package carrying these, edits text, a table cell and the comments, saves, and asserts that every
entry the edit did not own is byte-identical (the entry list is also unchanged apart from a created
`comments.xml`):

| Kept                                                                                                                                                                                                            | Notes                                                                                                                       |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `customXml/*` items, their properties and relationships                                                                                                                                                         | Rewritten only for a data-bound control the user edited.                                                                    |
| `word/glossary/document.xml` and its parts (building blocks)                                                                                                                                                    | Read only to show content control placeholder text.                                                                         |
| `word/vbaProject.bin` and the macro-enabled content type                                                                                                                                                        | Kept, never run. The extension of the saved file is the host's concern.                                                     |
| `word/people.xml`, `word/commentsIds.xml`, `word/commentsExtensible.xml`                                                                                                                                        | See "Comments" below.                                                                                                       |
| `word/fontTable.xml` and embedded fonts (`word/fonts/*.odttf`)                                                                                                                                                  | Not used for rendering; kept with their relationships.                                                                      |
| `word/settings.xml` (zoom, `w:extLst`, compatibility and protection elements)                                                                                                                                   | Only the flags the editor owns are patched (track changes, even/odd headers, hyphenation, ...).                             |
| `word/theme/theme1.xml`, `word/webSettings.xml`, `docProps/custom.xml`, thumbnails                                                                                                                              | Theme editing is refused (below).                                                                                           |
| Pictures and other media that are not replaced                                                                                                                                                                  | A picture the user replaces gets a new part.                                                                                |
| Unknown parts and relationships of any type                                                                                                                                                                     | Anything the package holds that the editor has no reader for.                                                               |
| Inside `document.xml`: `mc:AlternateContent`, extension lists (`w:extLst`), `w14:`/`w15:` markup, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, bookmarks, field codes, unknown run and paragraph properties | In paragraphs and tables the user did not restructure. Edited paragraphs keep their unmodeled `w:pPr` and `w:rPr` children. |

## Modeled and editable

Text and character formatting, paragraph formatting and styles (use and per-style formatting edits),
lists and numbering definitions, sections and page setup, headers, footers and footnotes, fields
(editable results, Update Fields), comments and tracked changes, content controls and form fields,
tables (below), hyperlinks and bookmarks, pictures, simple shapes and text boxes, SmartArt, classic
charts and admitted ChartEx, document properties, and the editing restrictions Word reads (form
protection is enforced; other restrictions are kept but not enforced).

### Tables

Row and column commands, merge and split, and the merged-cell layout are edited through the core
table structure engine (`ooxml-core/docx`, `table-structure.ts`) and written by the merge-aware
writer (`write-table-structure.ts`). The engine follows Word: its results for insert row, insert
column, delete row, delete column, merge and split were recorded from Word 16 on small tables
(`src/core/docx/__fixtures__/word-table-structure.json`, produced by
`src/core/scripts/docx/table-structure-word-probe.ps1`) and a test checks every one.

- A row or cell that survives an edit keeps its XML element, so `tblPr`, `trPr` and `tcPr` children
  the model does not hold, extension attributes and row `w14:paraId` values stay. Only `tblGrid`,
  `gridSpan`, `vMerge`, `tcW`, the cells an edit creates and the rows it adds or removes change.
- New rows copy the row they are made from (cell layout, height, repeat-as-header); new cells copy
  the formatting of the cell beside them. Deleting the first row of a vertical merge moves the merge
  to the next row and empties its text, as Word does.
- Nested tables are edited in place: an edit of the inner table changes only the inner `w:tbl`.
- A text-only edit of any table never rewrites its structure.

Refused (the command is disabled, and a model change the writer cannot express raises the error
below). Each refusal is a table Word itself turns into something a plain grid cannot hold:

| Edit                                                                                                                                                                                                                                                                          | Why                                                                                                                                                                        |
| ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Insert a column to the left of, or through, a merged cell (a cell spanning several columns)                                                                                                                                                                                   | Word adds `gridAfter` rows and a new cell as wide as the merged cell.                                                                                                      |
| Delete a column that only partly covers a merged cell                                                                                                                                                                                                                         | Word deletes the whole merged cell and leaves the row short (`gridAfter`).                                                                                                 |
| Split a cell spanning several columns into a different number of columns, or into more columns than it spans when it spans more than one                                                                                                                                      | Word re-grids the table. Splitting into the cell's own columns, and splitting a one-column cell, work.                                                                     |
| Split a vertical merge into a row count other than its height                                                                                                                                                                                                                 | Word accepts only divisors of the height; only the full unmerge is implemented.                                                                                            |
| Merge a range with a cell that sticks out of it, or cells holding nested tables                                                                                                                                                                                               | The result would lose content or the merge's shape.                                                                                                                        |
| Any structure edit of a table with row offsets (`gridBefore`, `gridAfter`), tracked row or cell changes (`w:ins`, `w:del`, `trPrChange`, `tcPrChange`), `hMerge`, content controls around rows or cells, or hyperlinks, fields, bookmarks or other inline markup in its cells | The load warning "Merged, nested, or complex tables can be read, but their row and column structure cannot be edited safely." says so; text in their cells stays editable. |

`saveDocx` throws `Cannot change the structure of a merged, nested, or complex table. The original
DOCX package remains unchanged.` for a structural change to such a table; the package is not
modified.

Table-level properties other than the margins the Table Properties dialog edits (width, alignment,
indent, borders, style, look) are rendered and preserved but not editable: changing them in the model
raises `Cannot edit table grid widths, width, alignment, indent, borders, or style; ...`.

## Refused with a diagnostic

Saving throws, naming the reason and leaving the source package untouched, when the model:

- edits the paragraph, character or table style catalog (`Editing the paragraph style catalog is not
  supported; source styles.xml is preserved unchanged.`, and the character and table variants), apart
  from the per-style formatting the editor owns;
- edits the theme (`Editing the theme catalog is not supported; source theme1.xml/settings.xml are
  preserved unchanged.`);
- repeats a block or paragraph id (`Duplicate document block id` / `Duplicate document node id`);
- asks for a picture, hyperlink, shape or chart that needs a relationship the caller cannot allocate;
- fails model validation (`assertValidDocumentModel`: for example a document with no section).

Loading is limited to 50 MB per package, 5,000 entries and 200 MB expanded (`DOCX_ZIP_LIMITS`); a
package over a limit is rejected before any part is read.

## Reported at load, not changed by a save

`model.warnings` lists features that load but do not render or edit like Word, for example
embedded OLE objects (kept, drawn as a placeholder), complex `HYPERLINK` fields (kept as field codes),
table-structure tracked changes (kept in the XML, not in the model), picture bullets and some
style-linked numbering, and editing restrictions that are kept but not enforced.

## Known limits of the guarantee

- **Comments.** When the comment list changes, `comments.xml` and `commentsExtended.xml` are written
  again from the model. Formatting inside a comment that the model does not hold is not carried over,
  and the kept `commentsIds.xml` / `commentsExtensible.xml` / `people.xml` can then describe comments
  by their old ids. Word ignores entries it cannot match.
- **Inline markup in table cells** blocks structural edits of that table (above) but not text edits.
- **Layout is not preserved information.** Cached page-break hints (`w:lastRenderedPageBreak`) in
  edited paragraphs are dropped; Word recomputes them.
- Content in a part the editor has no reader for is kept, not validated. A damaged part stays damaged.
