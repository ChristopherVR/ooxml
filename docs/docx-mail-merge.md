# Word mail merge API

`ooxml-core/docx/mail-merge` exposes mail merge for hosts that work with a `DocumentModel`. The host
loads recipients, selects records, previews a record, and saves the merged document models.
The Word editor's Mailings tab supports letters, directories, envelopes and labels, local CSV/TSV and first-sheet
XLSX files, recipient filtering and sorting, field insertion and matching, record preview,
validation, and combined DOCX and browser print output. A host can set
`docxEditor.mailMergeDataSource` to offer its own provider. The cancelable `file-command` event
with `command: 'openMergedDocument'`, `fileName` and `bytes` lets a host open the output as a
separate session. Without a handler the editor downloads it.

The editor reads a local CSV or TSV file in 1 MiB slices with an incremental parser
(`mailMergeDelimitedChunks`), so the file text is never held whole, and builds each record as
its row completes. Records are kept in memory once read, so every source (local files and the
host's provider) is bounded: at most 100,000 records, 255 fields, 4 million values in all
(for example 100,000 records of 40 fields), 16,384 characters per value and 32 Mi characters of
values in all. A delimited file may be up to 32 MiB. An `.xlsx` workbook is a zip package that
has to be in memory to be read, so it is capped at 20 MiB, its rows are then read one at a time.
Edit Recipient List shows 20 records per page and Preview Results builds one record at a time.
Finish & Merge still holds the merged document in memory, so its separate work limit
(`mailMergeWithinWorkLimit`) can ask for a smaller selection when a large template meets a long
list.

Address Block and Greeting Line insert native fields using the selected matches. Finish & Merge
asks for ASK and FILL-IN answers and can cancel the entire merge from a prompt. Preview and
Check for Errors use the saved defaults. Field insertion targets a paragraph in the document
body, and the Mailings controls require an editable, unshared document.

## Filling a template

`listMergeFields` and `mergeFields` fill the MERGEFIELD fields of one document from values, for
a service that generates documents from `.docx` templates in Node or a worker, without a DOM or
an editor. They sit next to `listFormFields` and `fillFormFields` in `ooxml-core/docx` (and
`docx-core`), and `ooxml-core/docx/mail-merge` exports the same two functions with the same
signatures. Both load their code on first use, so they return promises; the main entry only
carries the two loaders, which import `ooxml-core/docx/mail-merge` by name. Any pairing of
entries works (`loadDocx` from `ooxml-core/docx` or `loadDocument` from `ooxml-core/docx/load`,
either `mergeFields`, ESM or CommonJS): what the merge records for the writer is shared between
the entries' bundles. The CommonJS build loads the merge with a dynamic `import()`, as the
`math` and Visio chart loaders do, so a CommonJS test runner must allow it (Jest needs
`--experimental-vm-modules`).

```ts
import { listMergeFields, loadDocx, mergeFields } from 'ooxml-core/docx';

const template = await loadDocx(templateBytes);
const { fields, regions } = await listMergeFields(template.model);
const { model, results } = await mergeFields(
	template.model,
	{
		values: { 'First name': 'Ada', 'client.2.name': 'Lovelace', Notes: '<p>Hi <b>you</b></p>' },
		regions: {
			Items: [
				{ Name: 'Pen', Price: '10' },
				{ Name: 'Ink', Price: '20' },
			],
		},
		html: ['Notes'],
	},
	{ unlink: true },
);
const bytes = await template.save(model);
```

`listMergeFields(model)` returns `fields`, one entry per MERGEFIELD, story by story (the body,
then headers and footers, footnotes and endnotes; within a story the fields of table cells and
text boxes come before those of the paragraphs around them): its `name` as
written, its `instruction`, `story` (`body`, `header`, `footer`, `footnote`, `endnote` or
`textBox`; a table cell or content control reports the story it sits in), in a header or
footer its section `slot` (`default`, `first` or `even`), `kind` (`simple` for `w:fldSimple`, `complex` for
`w:fldChar`), its `switches` as written, the `region` it repeats in, whether it is `locked` and
whether it is `nested` in the code of another field. `regions` lists each region with its name,
`kind` (`rows` or `paragraphs`) and story.

`mergeFields(model, data, options)` returns a new model for the loaded template's `save()` (the
input model is not changed), with:

- `results`: one entry per MERGEFIELD evaluated, with its `name`, `story` and `slot`, `status`
  (`filled`, `missing`, `invalid`, `locked` or `unsupported`), a `reason` for every status but
  `filled` (such as `no value named "Cell" in row 1 of region Rows or in values`, the row counted from one in
  the text), and, inside a
  region, the `region` and the zero-based `row`. A `filled` field also has a `reason` when its
  value does not match its picture switch and is written as it is (`the value is
  not a number, so the \# picture was not applied`, and the same for `\@` and dates);
- `regions`: each region with the number of `rows` written for it;
- `htmlIssues`: what HTML values lost on the way in (`image-refused`, `link-refused`,
  `nested-table-flattened`).

Save the returned model with the template's own `save()`: repeated blocks and table rows are
written as copies of their source XML, which only that writer has.

### Values

Names match exactly as written in the field, including dots, digits, underscores, hyphens and
quoted names with spaces (`MERGEFIELD "First name"`). A value is text: a line break in it is
written as `w:br` and a tab as `w:tab`. Fields are filled in every story: the body, tables
(nested ones too), headers and footers of every type, footnotes, endnotes, DrawingML text
boxes and content controls, whose `w:sdt` and tags are kept. Simple and complex fields are
filled, including a code split across several `w:instrText` runs and a result split across
several runs. The MERGEFIELD switches apply as in the editor's merge (`\b`, `\f`, `\*`, `\@`
with `options.locale`, `\#`), and the result takes its formatting as described under
[Fields and settings](#fields-and-settings).

| Option                  | Default   | Effect                                                                                                                                                                                                                         |
| ----------------------- | --------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `unlink`                | `false`   | `true` replaces each field by its result, as Word's Unlink Fields (Ctrl+Shift+F9). `false` keeps each field and gives it a new cached result.                                                                                  |
| `missing`               | `'empty'` | A field whose name has no value shows nothing; `'keep'` keeps the result it has (unlinked, that result becomes the text).                                                                                                      |
| `removeEmptyParagraphs` | `false`   | Removes the paragraphs whose merge fields were all empty and that show nothing else (a table cell keeps one paragraph).                                                                                                        |
| `updateFields`          | `true`    | Refreshes the computed fields afterwards (DOCPROPERTY, document information, IF, `=`, numbering fields). PAGE, NUMPAGES and TOC keep their saved results (they need page layout), and so does REF, which is not evaluated yet. |
| `locale`                | `'en-US'` | The locale of `\@` date pictures.                                                                                                                                                                                              |

Locked fields (`w:fldLock`) are left as they are and reported as `locked`. A MERGEFIELD in the
code of an IF field is filled and the IF evaluated, as in the editor's merge; in the code of any
other field (QUOTE, for instance) the outer field keeps its result and is reported, under its
own name, as `unsupported`. So is a MERGEBARCODE field, which keeps its saved result.

A kept field (`unlink: false`) holds the value as its cached result, and Word shows it on
opening. Word's Update Field (F9) on such a document shows `«Name»`, because the document has no
data source attached (checked with Word 16); unlink the fields when the output must survive F9.
Fields in headers and footers are always replaced by their values, even when fields are kept:
Word 16 prints a MERGEFIELD kept in a header or footer of a document without mail merge settings
(which the merged document is) as `«Name»`, not its result (#368). The result of such a field
gives that as its `reason`.

### Repeating regions

A region opens with `MERGEFIELD TableStart:<name>` and closes with `MERGEFIELD TableEnd:<name>`.
Its content repeats once per row of `regions[<name>]`; each row's keys fill the fields inside it,
and a name the row does not have falls back to `values`, then to `missing`. The markers never
show in the output.

- Markers in the rows of one table (in one row, or spanning rows) repeat those rows. Each copy
  is a copy of the source row's XML, so its properties, borders, shading and merges are kept.
- Markers in the paragraphs of one story (the body, a header, a footer, a note or a text box)
  repeat the blocks from the paragraph holding `TableStart` to the one holding `TableEnd`,
  including when both are in the same paragraph.
- An empty or missing row list removes the region with its markers. A table that loses every
  row is removed; the rest of a table stays valid.

Regions inside other regions are not supported: the inner markers are only removed.

### HTML values

Only the names listed in `html` are read as HTML; any other value is text, whatever it holds.
HTML is read without a DOM and nothing is fetched. It goes through an allow-list: paragraphs and
headings (`Heading1` to `Heading6` when the document has those styles, bold otherwise), bold,
italic, underline, strikethrough, superscript, subscript, colour, line breaks, links (`http`,
`https`, `mailto` and `#anchor`), bulleted and numbered lists (each a new list definition) and
tables (`border` gives the table Word's default borders, `colspan` spans columns; `rowspan` is
not read). Every other element gives only its text; scripts, styles and comments are dropped,
and pictures and other links are refused and reported. Invalid HTML gives its text, never a
failed merge. One merge takes at most 1,024 distinct HTML values (more throw a `RangeError`)
and reads at most 20,000 tokens of each; the rest of a longer value is kept as text.

- Inline HTML (no block element) replaces the field in place, keeps the text around it and
  takes the field's run formatting where it sets none.
- Block HTML replaces the host paragraph when the field is alone in it, and is inserted after the
  paragraph otherwise. A field holding block HTML is unlinked even when fields are kept, since a
  field result here cannot hold paragraphs. Inside a table cell or a text box, an HTML table
  becomes one paragraph per row, its cells separated by tabs.

## Envelopes and labels

Headless hosts can call `await loadMailMergeStationery()` for the same label catalogue and
document operations. The catalogue (`labelCatalogue`, `searchLabels(query, vendor)`,
`findLabel(product, vendor)`) lists 39 sheet products from Avery US Letter, Online Labels,
Avery A4, Avery Zweckform and HERMA with the dimensions their vendors publish, and their inkjet
and colour variants as aliases (8160 finds 5160, J8160 finds L7160). It loads with the rest of
the stationery code, on demand. `createLabels({ product: 'custom', spec })` takes any other
layout, including labels with a vertical gap, which get spacer rows as in Word. Label Options
searches by product number, alias or description within the selected vendor.

Use Envelopes to add an opening envelope section to an existing document. The section records
the page size, address positions, and optional first-page printer tray. The print driver resolves
the tray number and physical feed settings. Labels opens a separate document using the selected
vendor and product. Use the full sheet or select a row and column for a single label.

From an empty document, Start Mail Merge can create an envelope template or a sheet of labels.
In a label template, insert content in the first label and choose Update Labels. It copies the
first label into the remaining label cells with a native NEXT field at the start of each copy.
Finish & Merge fills the sheet in row order and leaves unused labels blank. The preset catalog
covers Avery US Letter 5160 and 5163 and HERMA A4 4464.
See [HERMA's label dimensions](https://www.herma.com/office-home/labels/self-adhesive-labels/).

## Email output

Finish & Merge > Send Email Messages asks for the recipient column, subject and output format.
The browser fallback supports up to five recipients from a plain-text template and limits each
`mailto:` URL to 1,800 characters. The user opens each draft individually in their local mail
app and sends it there.
Email output uses one source record per message. A NEXT or NEXTIF rule that would
advance to another recipient stops email preparation and appears in Check for Errors.

Hosts can assign `editor.emailProvider` to an object with `send(message)`. Word calls it only
after the user reviews the recipient addresses and confirms Send messages. The method receives
`{ recipients, subject, text, attachments }` for text or
`{ recipients, subject, html, attachments }` for HTML. HTML includes escaped text, basic
emphasis and tables. DOCX attachments carry `fileName`, `contentType` and `bytes` as a
`Uint8Array`. The attachment mode supports up to ten recipients and email output is limited to
16 MiB per batch. The host owns transport and credentials. If delivery fails, Word stops and
reports the number accepted by the provider. Set `editor.emailProvider = false` to disable email.

## Recipient data

Each `MailMergeDataSource` has a `read()` method that returns `MailMergeData` or a promise
for it. Supply `records` and, optionally, the ordered `fields` list. A record maps field names
to strings, numbers, booleans, dates, or `null`. Field names match the source headers exactly.
Supply `fields` when a column can be absent from every record or when ODSO mappings use
column positions.

| Source               | API                                                                 | Data interpretation                                                                                                                                                                                                                                |
| -------------------- | ------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| CSV or TSV           | `mailMergeDelimitedText(text, { delimiter })`                       | The first row supplies trimmed field names. The delimiter defaults to auto-detection. Values stay strings.                                                                                                                                         |
| CSV or TSV, streamed | `mailMergeDelimitedChunks(() => chunks, { delimiter, dataLimits })` | The same rules, read from an iterable of text chunks (a `File` read in slices). `dataLimits` can bound records, fields, values, characters and value length while reading.                                                                         |
| XLSX                 | `mailMergeXlsx(bytes, { sheet })`                                   | The first row supplies trimmed field names. `sheet` accepts a zero-based index or a name and defaults to the first sheet. Formulas use cached values. Date-formatted cells supply ISO text when the XLSX reader can represent the value as a date. |
| Host records         | `mailMergeRecords({ records, fields })`                             | The host supplies the field names and typed values. A custom provider can implement the same interface.                                                                                                                                            |

Blank headers are ignored and duplicate headers raise an error. The core reads bytes and
records supplied by the host. The host owns credentials, source selection, requests, and
storage for any external provider.

`selectMailMergeRecipients(data, { filter, sort })` returns a new list. `filter` receives the
record and its zero-based source index. `sort` receives a pair of records and uses comparator
ordering. The filtered and sorted list determines the one-based record numbers.

## Preview, check, and finish

```ts
import { loadDocx } from 'ooxml-core/docx';
import {
	checkMailMerge,
	finishMailMerge,
	mailMergeDelimitedText,
	previewMailMerge,
	selectMailMergeRecipients,
} from 'ooxml-core/docx/mail-merge';

const template = await loadDocx(templateBytes);
const source = mailMergeDelimitedText(csv);
const data = selectMailMergeRecipients(await source.read(), {
	filter: (record) => record.Status === 'Active',
	sort: (a, b) => String(a.Name).localeCompare(String(b.Name)),
});

const issues = checkMailMerge(template.model, data);
if (data.records.length && issues.length === 0) {
	const preview = previewMailMerge(template.model, data, 1);
	showPreview(preview.document);
	const result = finishMailMerge(template.model, data);
	if (result.document) await writeOutput(await template.save(result.document));
}
```

- `previewMailMerge(model, data, recordNumber, options?)` returns a view model with evaluated
  fields replaced by their values. `skipped` reports a matching `SKIPIF`. The input model stays
  available for editing and saving.
- `checkMailMerge(model, data, options?)` runs the merge rules without creating output files.
  It reports an empty recipient list, unknown field names, unsupported fields, invalid
  instructions, and record rules in unsupported locations. Field issues include the block ID,
  field name, and one-based record position. Missing values in known columns produce blanks.
- `finishMailMerge(model, data, options?)` returns one combined document and a `mergedCount`.
  Letters start each copy on a new page. Directory copies start with a continuous section.
  Templates with internal section breaks keep
  those breaks. A template that ends with a table gains a paragraph when it needs a break.
- `finishMailMergeIndividually(model, data, options?)` returns `{ recordNumber, document }`
  entries in merge order. Both finish functions return the same field issues as Check for
  Errors. For output addressed to individual recipients, set `allowRecordAdvance: false`
  and reject any reported issues before sending. An empty list or a merge that skips
  every record produces no output documents.

Pass each finished model directly to the loaded template's `save()` method. This gives the
writer access to the original package, including styles, media, and source XML for repeated
body blocks. Use `saveDocx` for templates created in memory. Finish removes the mail merge
settings from the output. Check for Errors covers field evaluation. Package serialization can
still report errors from the DOCX writer.

## Fields and settings

| Field                  | Evaluation                                                                                                                                                                                              |
| ---------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `MERGEFIELD`           | Looks up the named column. `\b` and `\f` add text to a nonempty value. `\m` uses the ODSO mapping. `\v` converts ASCII letters, digits, and hyphens for vertical East Asian text.                       |
| `IF`                   | Evaluates nested merge fields in the comparison and the chosen branch.                                                                                                                                  |
| `NEXT`, `NEXTIF`       | Advance the current recipient inside the same output copy. If the list ends, later merge values are blank.                                                                                              |
| `SKIPIF`               | Discards the current output copy when the comparison is true. The next copy starts after the last record consumed.                                                                                      |
| `MERGEREC`, `MERGESEQ` | Show the current record position and the current output copy's sequence number, respectively.                                                                                                           |
| `ADDRESSBLOCK`         | Substitutes mapped address placeholders from `\f`. Supports country inclusion and exclusion with `\c` and repeated `\e` switches and selected locality orders with `\d` or `\l`.                        |
| `GREETINGLINE`         | Substitutes mapped name placeholders from `\f` and uses the `\e` fallback for missing names.                                                                                                            |
| `ASK`, `FILLIN`, `REF` | `ASK` assigns an answer to a bookmark name for later `REF` fields. `FILLIN` displays the answer. The `\d` switch supplies a default, or the cached result is used. The `\o` switch asks once per merge. |

Headless callers can supply `options.answerPrompt(request)` to either Finish API. Return a
string to use that answer or return `undefined` to use the saved default. For an asynchronous
prompt handler, use `resolveMailMergePrompts(model, data, ask)` first and pass its returned
options to Finish. An `undefined` answer from the asynchronous handler cancels the merge.

General `\*` switches reuse the Word field code's case and number formats. A merge result
takes its formatting as Word gives it (measured with Word 16's Update Field): with
`\* MERGEFORMAT` the formatting of the old result, with `\* CHARFORMAT` that of the first
letter of the field name, and otherwise that of the last run of the field code.
`\@` accepts ISO dates and times, `M/d/yyyy` text, and JavaScript `Date` values. Date pictures
use UTC, with `options.locale` for month and day names. `\#` supports numeric padding, decimal
and grouping marks, signs, literals, and sections for positive, negative, and zero values.
[Word field formatting](https://support.microsoft.com/en-us/word/format-field-results)

`DocumentModel.mailMerge` exposes `mainDocumentType`, `dataType`, `query`, the `dataSource`
and `odsoSource` relationship references, and the ordered `fieldMappings`. Edit the model
and use the normal DOCX save path to update these settings. Other settings and ODSO children
remain in the source package. New source targets default to external relationships. Queries
and targets are metadata for the host. `options.fieldMappings` can override the saved mappings
for a merge. [Open XML mail merge settings](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.mailmerge)

The engine covers body paragraphs, table cells, headers, footers, footnotes, endnotes and the
DrawingML text boxes (`wps:txbx`) in any of them. Record rules run at the top level of the body;
fields in text boxes merge but do not advance the record. A complex field whose code or result
spans paragraphs (an `IF` whose branch text holds paragraph marks, or a merge field whose saved
result does) merges as one field: the paragraphs it covers are evaluated together and split
again at the paragraph marks that remain, each taking the formatting of the paragraph whose mark
ends it, as in Word. A field that crosses a table, never closes or spans more than 256
paragraphs keeps its source result, and so do locked fields. A merged text box writes its new
text into its `w:txbxContent` and drops its VML fallback copy rather than leave it stale;
legacy VML-only text boxes (`w:pict` without DrawingML) are kept unmerged. Desktop Word
compatibility remains open for the full set of regional address and greeting formats,
printer-specific envelope settings, barcodes, caption picture tokens and the full range of
template objects.
