# Word mail merge API

`ooxml-core/docx` exposes mail merge for hosts that work with a `DocumentModel`. The host
loads recipients, selects records, previews a record, and saves the merged document models.
The Word editor's Mailings tab still needs UI integration.

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
import {
	checkMailMerge,
	finishMailMerge,
	loadDocx,
	mailMergeDelimitedText,
	previewMailMerge,
	selectMailMergeRecipients,
} from 'ooxml-core/docx';

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
  Each output copy starts a new page section. Templates with internal section breaks keep
  those breaks. A template that ends with a table gains a paragraph when it needs a break.
- `finishMailMergeIndividually(model, data, options?)` returns `{ recordNumber, document }`
  entries in merge order. Both finish functions return the same field issues as Check for
  Errors. An empty list or a merge that skips every record produces no output documents.

Pass each finished model directly to the loaded template's `save()` method. This gives the
writer access to the original package, including styles, media, and source XML for repeated
body blocks. Use `saveDocx` for templates created in memory. Finish removes the mail merge
settings from the output. Check for Errors covers field evaluation. Package serialization can
still report errors from the DOCX writer.

## Fields and settings

| Field                  | Evaluation                                                                                                                                                                        |
| ---------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `MERGEFIELD`           | Looks up the named column. `\b` and `\f` add text to a nonempty value. `\m` uses the ODSO mapping. `\v` converts ASCII letters, digits, and hyphens for vertical East Asian text. |
| `IF`                   | Evaluates nested merge fields in the comparison and the chosen branch.                                                                                                            |
| `NEXT`, `NEXTIF`       | Advance the current recipient inside the same output copy. If the list ends, later merge values are blank.                                                                        |
| `SKIPIF`               | Discards the current output copy when the comparison is true. The next copy starts after the last record consumed.                                                                |
| `MERGEREC`, `MERGESEQ` | Show the current record position and the current output copy's sequence number, respectively.                                                                                     |
| `ASK`, `FILLIN`, `REF` | `ASK` assigns the `\d` default to a bookmark name for later `REF` fields. `FILLIN` displays the default. A missing default uses the cached result.                                |

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
their source results. Desktop Word compatibility remains open for address and greeting
blocks, barcodes, interactive prompts, text boxes, caption picture tokens, and the full range
of locale-specific formatting and template objects.
