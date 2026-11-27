# Content control templates

`fillContentControls`, `readContentControls` and `flattenContentControls` fill a `.docx` template
whose variable parts are tagged content controls (`w:sdt` with a `w:tag`), read the values back
after a person edited the result in Word, and fill the same document again with new data. They
work on a `DocumentModel` in Node or a worker, without a DOM or an editor. They are the content
control counterpart of the MERGEFIELD merge (`docs/docx-mail-merge.md`), and sit next to
`listFormFields` and `fillFormFields`, which set single controls and legacy form fields by id.

`ooxml-core/docx/content-controls` exports the three functions synchronously. `ooxml-core/docx`
(and `docx-core`) exports the same functions loaded on first use, so there they return
promises; the main entry carries only the loaders, which import `ooxml-core/docx/content-controls`
by name. What a fill records for the writer (the source of each repeated block and row) is shared
between the entries' bundles, so any pairing of entries, ESM or CommonJS, saves the copies with
their source formatting.

```ts
import { loadDocx } from 'ooxml-core/docx';
import { fillContentControls, readContentControls } from 'ooxml-core/docx/content-controls';

const template = await loadDocx(templateBytes);
const { model, results, htmlIssues } = fillContentControls(
	template.model,
	{
		values: { customer: 'Ada Lovelace', notes: '<p>Paid <b>in full</b></p>', note_optional: '' },
		html: ['notes'],
		repeat: {
			items: [
				{ desc: 'Pen', qty: '2' },
				{ desc: 'Ink', qty: '5', discount: '10%' },
			],
		},
	},
	{ optionalSuffix: '_optional', lock: { customer: 'sdtContentLocked' } },
);
const bytes = await template.save(model);

// Later, after the document was edited in Word:
const filled = await loadDocx(bytes);
const values = readContentControls(filled.model, { as: 'html' });
const again = fillContentControls(filled.model, newData, { template: template.model });
```

## Values

`values` are by tag. Each control with that tag is filled, in every story: the body, headers,
footers, footnotes and endnotes, table cells and tables nested in cells. Text boxes are not
filled.

- A rich text control takes the value as plain text, or as HTML when its tag is in `html`. A
  plain text value's lines become paragraphs in a block or cell control and line breaks in an
  inline one.
- HTML is read through the merge's allow-list (`mail-merge/html-content.ts`): paragraphs and
  headings, bold, italic, underline, strikethrough, super- and subscript, colour, line breaks,
  links (http, https, mailto and anchors), lists and tables. Every other element gives its text;
  scripts, styles, comments and pictures are dropped and nothing is fetched. HTML that is not
  well formed is read as far as it goes and falls back to its text, never failing the fill.
  `htmlIssues` says what was refused or flattened (a table in a cell or in an inline control
  becomes one paragraph or line per row).
- The control's own typography wins: new runs take the font, size and other run properties of
  the control's first text run (the Placeholder Text style is dropped), and the HTML adds only
  emphasis, colour, links and structure. `font-family` and `font-size` in the HTML are ignored.
  New paragraphs of a block control keep the paragraph formatting of its first paragraph
  (alignment, spacing, indents, style), unless the HTML sets an alignment, a heading or a list.
- In an inline control, several HTML paragraphs become one run sequence separated by line breaks.
- An empty value (only spaces, `&nbsp;`, or tags without text) clears the control to one empty
  paragraph or run, and the control stops showing its placeholder.
- A plain text control (`w:text`) takes text only: HTML gives it its text, and the result says so.
- Check boxes take `true`/`false` (also `1`, `0`, `yes`, `no`, `x`), drop-down lists and combo
  boxes an entry's value or display text (a combo box also other text), dates a day as
  `YYYY-MM-DD`, shown with the control's date format. A value that does not fit is not written,
  and its result says why.
- A control bound to a custom XML part (`w:dataBinding`) is kept as it is. A control whose content
  is locked (`contentLocked`, `sdtContentLocked`) is filled only when `options.lock` names its tag
  (it is locked again after). A temporary control (`w:temporary`) is removed once filled, as Word
  does, leaving its new content. Group, picture, citation, equation and building-block controls
  are not filled with a value; the controls inside a group are.

A control around other controls reads through to them: when its own tag has no value, the
controls inside it are filled.

## Repeating

`repeat` is by tag of the control to repeat, one record per copy. A record's keys are the inner
controls' tags as they are in the template (`item_template_qty`), or what follows the `template`
part of the tag (`qty`); a tag a record lacks falls back to `values`.

- Word's repeating section (`w15:repeatingSection`, at block or row level) gives one copy of its
  first item (`w15:repeatingSectionItem`) per record, inside the section. Later items (the copies
  of an earlier fill) are replaced.
- A plain block control, or a row control, named in `repeat` is copied whole per record. Copies
  keep the control's tag; later copies of an earlier fill that follow it are replaced.
- A table row is repeated when a control in one of its cells has the repeat's tag, or a tag
  `<key>_template_...` (with `_`, `.` or `-` between the parts). The rows right after it whose
  tagged controls are all `<key>_template_...` are its sub-rows (a discount, a subtotal): a copy
  includes a sub-row only when its record fills one of the sub-row's controls.
- The controls inside each copy are retagged with its number (`item_template_qty` becomes
  `item_1_qty`, `item_2_qty`; a tag without a `template` part gets `_1`, `_2`), so each copy is
  addressable when reading back. Every copy's controls get new ids.
- Copies in a story or of a table row keep their source formatting: the writer copies the source
  paragraph, table or row (`w:trPr`, borders, shading, heights, fonts) and writes only what the
  fill changed. Copies inside table cells (a repeated block control in a cell) are written from
  the model.
- An empty record list removes the repeated control and leaves a valid table (a table left
  without rows is removed).
- Rows of a table nested in a cell are not repeated or removed (the result says so).

## Optional controls

With `optionalSuffix`, a control whose tag ends with it stays only when its own value, or a
control inside it, shows a value (not its placeholder, not blank, a checked box). Otherwise it is
removed with its content: a block control with its paragraphs, an inline control with its runs, a
row control with its rows, a cell control's content (the cell stays, empty). A plain row emptied
this way (no text and no control left) is removed, and a cell or story always keeps a paragraph.

## Locks

`options.lock` sets locks by tag after filling: `sdtLocked` (the control cannot be deleted),
`contentLocked` (its content cannot be edited), `sdtContentLocked` (both) or `none`. It applies to
every control with the tag at any level, copies included.

## Filling an already filled document

Pass the template the document was filled from as `options.template`. In each story (the body,
each header and footer), every top-level block that holds a content control comes from the
template, and every block without one comes from the document: an optional block a previous fill
removed is back at its place, repeated copies are replaced by the new items, and the paragraphs
the user typed outside the controls stay between the controlled blocks they followed (a block the
document lost comes back after the text typed before the next block both have). The result is the
one filling the template with the same data gives, plus that text. The template's blocks are
written from the model, so XML the model does not hold of those blocks (unmodeled paragraph or
table properties) comes from neither file; text edited inside a controlled block is replaced too.
Without `template`, copies are retagged and cannot be matched to the records again.

## Reading back

`readContentControls(model, { prefix, as })` returns each tagged control's content by tag: the
body first, then headers, footers and notes, in reading order; the first control of a tag wins
(copies of a repeated control share its tag; their inner controls are numbered). A control around
others reads through to them, table cells included. `as: 'html'` writes paragraphs, headings,
emphasis, colour, links, line breaks, lists and tables, without fonts, so reading and filling
again keeps the template's typography; an inline control reads as runs without a paragraph. A
control showing its placeholder reads as empty; check boxes read as `true` or `false`, dates as
`YYYY-MM-DD`, list controls as the chosen entry's value.

## Flattening

`flattenContentControls(model, { removeHiddenParagraphs })` replaces every control (inline, block,
row and cell, nested ones too) by its content, which keeps its formatting; a control showing its
placeholder leaves the placeholder text, as Word does. With `removeHiddenParagraphs`, paragraphs
whose paragraph mark is hidden text and whose runs are all hidden go too (a paragraph with visible
text is kept), before an export to PDF.

## Results

`results` has one entry per control acted on: `tag`, `status` (`filled`, `cleared`, `repeated`,
`removed`, `kept`, `locked`, `skipped`, `missing`), the `story` and header or footer `slot`,
`items` for a repeat, `item` inside a copy (its number) and a `reason` for `skipped` (bound,
locked, the kind, a value that does not fit) and `missing` (a tag in the data that no control
has).

## Not supported

- Inline controls nested in inline controls: the model keeps only the innermost, so the outer one
  is neither filled nor read (the XML is kept).
- Controls in DrawingML text boxes are not filled.
- Pictures in HTML values, `rowspan`, and nested HTML tables (flattened).
- A refill without `template`.
- Data-bound controls are not synchronized with their custom XML part.
