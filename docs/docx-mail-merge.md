# Word mail merge API

`ooxml-core/docx/mail-merge` exposes mail merge for hosts that work with a `DocumentModel`. The host
loads recipients, selects records, previews a record, and saves the merged document models.
The Word editor's Mailings tab supports letters, directories, envelopes and labels, local CSV/TSV and first-sheet
XLSX files, recipient filtering and sorting, field insertion and matching, record preview,
validation, and combined DOCX and browser print output. A host can set
`docxEditor.mailMergeDataSource` to offer its own provider. The cancelable `file-command` event
with `command: 'openMergedDocument'`, `fileName` and `bytes` lets a host open the output as a
separate session. Without a handler the editor downloads it. The editor keeps recipient data in
memory and accepts up to 1,000 records, 64 fields and 5 MiB per local file. Large merges may need
a smaller selection.

Address Block and Greeting Line insert native fields using the selected matches. Finish & Merge
asks for ASK and FILL-IN answers and can cancel the entire merge from a prompt. Preview and
Check for Errors use the saved defaults. Field insertion targets a paragraph in the document
body, and the Mailings controls require an editable, unshared document.

## Envelopes and labels

Headless hosts can call `await loadMailMergeStationery()` for the same presets and document
operations.

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

| Source       | API                                           | Data interpretation                                                                                                                                                                                                                                |
| ------------ | --------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| CSV or TSV   | `mailMergeDelimitedText(text, { delimiter })` | The first row supplies trimmed field names. The delimiter defaults to auto-detection. Values stay strings.                                                                                                                                         |
| XLSX         | `mailMergeXlsx(bytes, { sheet })`             | The first row supplies trimmed field names. `sheet` accepts a zero-based index or a name and defaults to the first sheet. Formulas use cached values. Date-formatted cells supply ISO text when the XLSX reader can represent the value as a date. |
| Host records | `mailMergeRecords({ records, fields })`       | The host supplies the field names and typed values. A custom provider can implement the same interface.                                                                                                                                            |

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

General `\*` switches reuse the Word field code's case and number formats. Merge results
take the cached result run's formatting, or the instruction run's formatting for `CHARFORMAT`.
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

The engine covers body paragraphs, table cells, headers, footers, footnotes, and endnotes.
Record rules run at the top level of the body. Locked fields and cross-paragraph fields keep
their source results. Desktop Word compatibility remains open for the full set of regional
address and greeting formats, printer-specific envelope settings, custom label sizes,
barcodes, text boxes, caption picture tokens and the full range of template objects.
