# Comparing documents (`ooxml-core/docx`)

`compareDocuments(original, revised, options)` is Word's Review > Compare: it returns a new DOCX
package whose content is the revised document, with everything that changed since the original
recorded as tracked changes by one author. Accepting every change gives the revised body text;
rejecting every change gives the original body text. Only the body is compared (see
[Not compared](#not-compared)): rejecting does not bring back the original's headers, footers,
notes, paragraph formatting, or the pictures and note references the result could not carry. The
code is in `src/core/docx/compare/`; it is DOM-free and changes neither input.

```ts
import { compareDocuments, listRevisions } from 'ooxml-core/docx';

const result = await compareDocuments(v1Bytes, v2Bytes, {
	author: 'Reviewer',
	granularity: 'word', // or 'character'
});
await writeFile('compared.docx', result.bytes);
console.log(result.summary); // { insertions, deletions, formatChanges }
listRevisions(result.model); // the same changes, as the editor lists them
```

`result.model` is the saved package loaded again, so the model and the bytes agree. `warnings`
names content the comparison could not carry over, and `notCompared` lists every part it never
looks at (see below).

## How it compares

The rules follow Word 16's observable output (`Application.CompareDocuments`, new document):

- **Tracked changes in either input are accepted first**, as Word does by default ("Word will
  consider these changes accepted for the purpose of the comparison").
- **Paragraphs with the same text are aligned first** (Myers' O(ND) difference over the
  paragraphs). The paragraphs between two aligned ones are compared as one stream of words (or
  characters), paragraph marks included, so a changed paragraph shows only its changed words, a
  removed paragraph is deleted with its mark, and text merged from two paragraphs shows the
  deleted mark between them.
- **The last paragraph mark of the body or a table cell is never marked**: a paragraph added at
  the end marks the previous paragraph's mark as inserted, as Word does.
- **Word granularity** splits text into words, single spaces and punctuation; spaces left between
  two changes join them (`Alpha [BETA gamma][beta GAMMA].`, not two replacements around a kept
  space), and an insertion or deletion that could sit in several places is placed with its edges
  at paragraph marks, then at the start of a sentence, then at spaces. **Character granularity** marks single characters
  (`colo[u]r`).
- Deletions come before insertions in a replaced span. Each contiguous span is one revision; ids
  are numbered on save by the same allocator the editor uses.
- **Formatting changes** of unchanged text are recorded as `w:rPrChange` with the original run
  properties (direct run formatting is compared; `compareFormatting: false` turns this off).
- **Tables** are compared row by row: rows with the same text pair, the rows of a changed stretch
  pair by their words, paired rows are compared cell by cell, and the rest are tracked row
  insertions and deletions (`w:trPr/w:ins`, `w:trPr/w:del`) with their content marked. Rows
  with nested tables are kept as they are when their text is the same. A table the result cannot
  be saved over is shown deleted and inserted whole: one whose rows the writer cannot restructure
  (merged cells, nested tables, tracked changes inside its cells in the source), or one that
  gained or lost a column. A deleted or inserted table with merged cells shows as deleted or
  inserted paragraphs.
- Pictures, field markers, breaks and other non-text runs are compared whole, never split. A
  simple field (`w:fldSimple`) is marked through its result runs only: rejecting its insertion
  in Word removes the result text but keeps the field, which Word then updates.
  Pictures and note references deleted from the original are not carried into the result (a
  warning says so); deleted text keeps its direct formatting and any style or list the revised
  document also defines. A revised paragraph whose runs carry formatting the writer cannot split
  (and, when the writer refuses a picture, every paragraph with a picture) is shown inserted
  whole with the original text deleted, so the result still saves; a warning says so.

## Not compared

`notCompared` (`UNCOMPARED_DOCUMENT_PARTS`) lists them in every result:

| Part                  | What the result shows                                                          |
| --------------------- | ------------------------------------------------------------------------------ |
| `headersAndFooters`   | The revised document's headers and footers.                                    |
| `notes`               | The revised document's footnotes and endnotes.                                 |
| `comments`            | The revised document's comments; anchors in deleted text are dropped.          |
| `textBoxes`           | Text boxes, shapes and drawings compared as whole objects only.                |
| `moves`               | Moved text as a deletion and an insertion, never `w:moveFrom`/`w:moveTo`.      |
| `paragraphFormatting` | The revised paragraph properties; no `w:pPrChange`.                            |
| `tableFormatting`     | The revised table, row and cell properties; no `w:tblPrChange`/`w:tcPrChange`. |
| `sections`            | The revised section and page setup.                                            |

## Checking against Word

`src/core/docx/__fixtures__/compare/generate-word-compare.ps1` (Windows, Word 16, manual) builds
original and revised pairs with Word and saves Word's own comparison of each (`*-word.docx`).
`compare.word.test.ts` checks that `compareDocuments` marks the same inserted, deleted and
reformatted spans (`revisionSpans`) as Word for eight pairs: word edits with paragraphs added
and removed, formatting changes, inputs with their own tracked changes, character granularity,
table rows, a paragraph split in two, an inserted sentence that repeats the words around it and
a deleted last paragraph.
The pairs were compared with Word's moves detection off (`CompareMoves` false), since moves are
not compared here. The results also open in Word with the same revision count, and rejecting
them all in Word gives the original text. Each fixture result's accept-all and reject-all in the
model give the revised and original body text (`compare.test.ts`).
