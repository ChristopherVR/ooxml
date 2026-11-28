# Word layout issue fixtures

Borris-pro attached the original generated samples to the issues below. `cell-end-marker.docx` is a generated follow-up for #330. The anonymized production documents were used locally for comparison.

Word 16.0.20625 exported each fixture through `Document.ExportAsFixedFormat`. The fixture tests exercise the same DOCX import and layout engine as Print Layout and PDF export.

| Issues                                                                                                               | Word observations                                                                                                                                                                                                                                                                                                                                            |
| -------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| [#329](https://github.com/ChristopherVR/ooxml/issues/329)                                                            | Three 0.25 pt outlines and a label sized independently of the legacy group coordinates.                                                                                                                                                                                                                                                                      |
| [#330](https://github.com/ChristopherVR/ooxml/issues/330)                                                            | With the current Normal template, both styleless 12 pt samples have seven pages and 43 lines on page one. With an explicit styles part and single spacing, Times New Roman advances 13.8 pt per line and has 50 lines on page one. At 12 pt, Word modes 11, 12 and 14 required 286 pt to fit the test phrase on a justified line. Mode 15 fit it at 279 pt.  |
| [#331](https://github.com/ChristopherVR/ooxml/issues/331), [#332](https://github.com/ChristopherVR/ooxml/issues/332) | Both tall-footer samples have three pages. The bordered footer is indented and has one outer box with continuous side borders.                                                                                                                                                                                                                               |
| [#333](https://github.com/ChristopherVR/ooxml/issues/333)                                                            | The picture starts at x=420.75 pt in the header and x=421 pt in the body control.                                                                                                                                                                                                                                                                            |
| [#334](https://github.com/ChristopherVR/ooxml/issues/334)                                                            | Triple borders have three strips. The thin/thick variants preserve the number, order and relative widths of the strips.                                                                                                                                                                                                                                      |
| [#335](https://github.com/ChristopherVR/ooxml/issues/335), [#336](https://github.com/ChristopherVR/ooxml/issues/336) | The signature keeps its full height with the preceding text readable. The wrapped line ends near the left edge of the picture while retaining right alignment.                                                                                                                                                                                               |
| [#337](https://github.com/ChristopherVR/ooxml/issues/337), [#338](https://github.com/ChristopherVR/ooxml/issues/338) | The missing font is substituted with Cambria. The omitted and explicit 10 pt controls use the same size.                                                                                                                                                                                                                                                     |
| [#339](https://github.com/ChristopherVR/ooxml/issues/339), [#340](https://github.com/ChristopherVR/ooxml/issues/340) | The percentage shades export as solid grays. Highlight exports in yellow and small capitals use reduced glyphs for the lower-case part of the text.                                                                                                                                                                                                          |
| [#341](https://github.com/ChristopherVR/ooxml/issues/341), [#342](https://github.com/ChristopherVR/ooxml/issues/342) | The uncached merge field has no displayed result. Footer tabs advance to both the document's default intervals and the explicit 5 cm stop.                                                                                                                                                                                                                   |
| [#343](https://github.com/ChristopherVR/ooxml/issues/343), [#344](https://github.com/ChristopherVR/ooxml/issues/344) | The two groups of empty lines use their respective 2 pt and 8 pt paragraph marks. The hanging-indent border extends 284 twips toward the start margin.                                                                                                                                                                                                       |
| [#346](https://github.com/ChristopherVR/ooxml/issues/346)                                                            | Mode 15 retains all three reported final words and uses two pages. Mode 14 wraps them and uses three pages.                                                                                                                                                                                                                                                  |
| [#347](https://github.com/ChristopherVR/ooxml/issues/347)                                                            | Implicit and explicit Normal suppress the same gaps, including those beside empty paragraphs. Both samples fit on one page.                                                                                                                                                                                                                                  |
| [#348](https://github.com/ChristopherVR/ooxml/issues/348)                                                            | Zero character scale uses the same text widths and two pages as the unscaled control.                                                                                                                                                                                                                                                                        |
| [#349](https://github.com/ChristopherVR/ooxml/issues/349)                                                            | Inline title and closing tags occupy space in both dimensions. The tagged file occupies three pages and the default-appearance control occupies two.                                                                                                                                                                                                         |
| [#353](https://github.com/ChristopherVR/ooxml/issues/353)                                                            | TableGrid supplies single spacing with zero after-spacing inside its cell. The table matches the height of the direct-formatting control and fits on one page.                                                                                                                                                                                               |
| [#359](https://github.com/ChristopherVR/ooxml/issues/359)                                                            | Contextual spacing drops only the flagged paragraph's own space beside a paragraph of the same style. The one-sided sample has two pages (33 paragraphs on page one) and the control one.                                                                                                                                                                    |
| [#360](https://github.com/ChristopherVR/ooxml/issues/360)                                                            | Multiple line spacing puts its extra space below the text: the first 1.15 line after a single line steps 14.64 pt, later ones 16.80 or 16.92 pt. The last line's extra space may pass the bottom margin, so 42 lines fit on page one, and the control fits too.                                                                                              |
| [#383](https://github.com/ChristopherVR/ooxml/issues/383)                                                            | Century Gothic 10 pt at 1.08 lines steps 13.23 pt, so page one ends after 48 lines: one line of the eighth paragraph without widow control, none with it (three pages). The single-spaced control keeps 52 lines.                                                                                                                                            |
| [#391](https://github.com/ChristopherVR/ooxml/issues/391)                                                            | Times New Roman 12 pt with a Garamond comma steps 14.364 pt (largest ascent, largest descent, the line gap above), 48 lines on page one; the one-font control steps 13.802 pt.                                                                                                                                                                               |
| [#393](https://github.com/ChristopherVR/ooxml/issues/393)                                                            | A run holding only a space in a larger font (12 pt bold in 11 pt Calibri) does not raise its line: the first baseline stays at 760.7 pt as in the 11 pt control, 26 paragraphs on page one.                                                                                                                                                                  |
| [#395](https://github.com/ChristopherVR/ooxml/issues/395)                                                            | A Symbol bullet over 11 pt Calibri: the first line is Symbol's ascent plus Calibri's descent (14.01 pt), the next 13.44 pt; 25 items on page one.                                                                                                                                                                                                            |
| -------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| [#406](https://github.com/ChristopherVR/ooxml/issues/406)                                                            | A space next to a non-breaking space is no break opportunity: with `John` NBSP space `SMITH,` or `John` space NBSP `SMITH,` Word moves `John` down with `SMITH,` in ten of the twelve paragraphs. `issue-406-space-next-to-nbsp.docx` and `issue-406-two-spaces-control.docx` are rebuilt from the issue's paragraphs by `generate-line-break-fixtures.mjs`. |
| [#410](https://github.com/ChristopherVR/ooxml/issues/410)                                                            | A start tag that directly follows a word is no break opportunity: Word keeps `Name/Field`, both start tags and `Value` together and ends line 1 at `amet`. With a space before the tags it breaks at the space. `issue-410-start-tag-after-word.docx` and `issue-410-start-tag-after-space.docx` are rebuilt from the issue's paragraph.                     |

`issue-359-contextual-spacing-rows.docx` is a synthetic file with eleven pairs of Normal
paragraphs that set contextual spacing, space before and space after on either side. Word's gaps
for each pair, from `Range.Information(6)`, are in `issue-359-contextual-spacing-rows.json`. When the
upper paragraph drops its space after, Word still subtracts that space after from the lower
paragraph's space before. `issue-360-line-spacing-rules.docx` has three one-line paragraphs for each of single, 1.15,
1.5, double and 0.8 spacing, exact 20 pt and at least 20 and 10 pt, in Times New Roman.
`issue-360-triple-spacing-page-fit.docx` has 120 one-line paragraphs at triple spacing: Word
puts 17 on each page because the extra space of each page's last line passes the bottom margin.
Word's PDF baselines for both and for the reporter's #360 samples are in
`issue-360-line-spacing.json`. Regenerate the synthetic files with
`python src/core/docx/layout/fixtures/word-layout-reports/generate-spacing-fixtures.py`.

`issue-353-table-style-precedence.docx` is a synthetic Lorem ipsum document without personal
metadata. Normal sets 12 pt after and left justification, while the table style sets zero after
and center justification. Its first-row condition adds 4 pt before. The cell covers implicit and
explicit Normal, a named style and direct formatting. Normal sets 11 pt for non-complex text and
12 pt for complex text. The table style sets 16 pt and 18 pt respectively. The named style inherits
Normal's non-complex size and sets its own 10 pt complex size. The last paragraph sets a direct
9 pt non-complex size. Tests save and reload the 10 pt Normal and compatibility-setting variants.
The expected precedence follows MS-DOCX 2.3.1. Regenerate it with
`bun src/core/docx/layout/fixtures/word-layout-reports/generate-table-style-precedence.mjs`.

For #330, Word moved all three lines of a paragraph together when widow and orphan control
would have left one line on either page. At a 453.5 pt text width, an italic Times New Roman
control wrapped before its final word with kerning omitted and fit that word with kerning
enabled. A 12 pt Arial list marker raised the baseline of a 12 pt Times New Roman line
without increasing its 13.8 pt height. In `cell-end-marker.docx`, Word omitted the empty final
paragraph in a cell with `w:hideMark`. An explicit break or space in the cell still took room.

The follow-up #330 check compared all 2,238 rendered contract body and table line starts and
their page assignments to Word COM. It compared the visible prefix length of each Word range
so hidden field instructions did not affect the offsets. All 600 lines of the two single-spacing
samples also matched Word. The contract's four hidden end-of-cell marks occupy no line.

Word COM wrapping probes are captured by `word-corpus-rules.test.ts`. Word broke after an
unspaced en or em dash, preserved leading spaces after explicit breaks, and used the line's
remaining width for a long word that spans two formatting runs. In mode 15, the amount a
justified line could give up depended on the spaces available and the width of the next word.

The synthetic #349 follow-up `block-content-control-tags.docx` covers a control spanning three
paragraphs and a control spanning two table rows, nested with a control spanning two cells.
The reporter's samples contain inline controls only. Regenerate the follow-up with
`python src/core/docx/layout/fixtures/word-layout-reports/generate-block-content-control-tags.py`.

For #349, `measured-inline-content-control-tags.docx` is a synthetic Lorem ipsum fixture
with no personal metadata. It combines the four Arial sizes from the Word COM probes.
Each size has an absent title, a short title, a long title and a bounding-box control.
A manual break follows the control so the next line start gives the tagged line height.
The Node generator and full COM coordinates and PDF exports are in the issue's scratch
directory `word-measure/issue-349`.

Word 16.0.20625.20108 Print Layout measurements use points from
`Range.Information(5)` and `Range.Information(6)`. The two edge advances are measured
between the one-character ranges straddling each marker.

| Body pt | Start, absent | Start, A | Start, Lorem Ipsum Dolor | Inline end | Tagged line | Bounding-box line |
| ------- | ------------- | -------- | ------------------------ | ---------- | ----------- | ----------------- |
| 8       | 18.5          | 18.5     | 81.5                     | 11.5       | 18.5        | 9                 |
| 11      | 18.5          | 18.5     | 81.5                     | 11.5       | 18.5        | 12.5              |
| 16      | 18.5          | 18.5     | 81.5                     | 11.5       | 19          | 18.5              |
| 24      | 18.5          | 18.5     | 81.5                     | 11.5       | 27.5        | 27.5              |

Times New Roman at 11 pt gave the same figures as Arial. Toggling `ShowXMLMarkup` left
the original figures unchanged. Both bounding-box and hidden appearances gave zero
advances and a 12.5 pt line at 11 pt.

| Structural probes, 11 pt | COM start advance | PDF start outline | PDF end outline |
| ------------------------ | ----------------- | ----------------- | --------------- |
| Lorem Block              | 56.5              | 56.66             | 56.64           |
| Lorem Row                | 53.5              | 53.3              | 53.3            |
| Lorem Cell               | 51                | 50.78             | 50.76           |
| Absent title             | 18.5              | 18.5              | 18.48           |

Word's print-optimized PDF contains both tag outlines and 8 pt Tahoma text. The inline
closing tag is unlabelled. Structural tags repeat the title on the closing edge. The
tag boxes are 13.44 pt high and the title baseline is 3 pt above the body baseline.
The PDF advances for `A` and `Lorem Ipsum Dolor` are 4.83 and 68.06 pt at the label size.
The reporter's tagged and bounding-box files had 3 and 2 pages in both
`ComputeStatistics(2)` and PDF, with 240 and zero PDF tag outlines respectively.

The tag outlines in the PDF content streams are 227-point polylines on a 0.12 pt grid,
filled and stroked with one `B*` and a round join (`1 j`). Comparing the tagged and
bounding-box exports isolates them; they were the same at 8, 11, 16 and 24 pt and in
Times New Roman. `content-control-tag-shape.ts` draws them with cubic arcs that stay
within 0.13 pt of Word's vertices.

| Outline property      | Word PDF                                                          |
| --------------------- | ----------------------------------------------------------------- |
| Inline fill           | `0.925 0.949 0.969 rg` (#ECF2F7)                                  |
| Block, row, cell fill | `0.988 g` (#FCFCFC)                                               |
| Stroke                | `0.627 0.745 0.843 RG` (#A0BED7), 1.08 pt                         |
| Title                 | Tahoma 8.04 pt, `0 g`, baseline 4.68 pt above the box bottom      |
| Box                   | 13.44 pt, from 1.68 pt below to 11.76 pt above the body baseline  |
| Side edges            | Arcs of a circle of radius 11.04 pt, bowing 2.28 pt across        |
| Opening tag           | Both edges bow left; the outline spans the whole advance          |
| Closing tag           | Both edges bow right; inline, 9.84 pt of the 11.5 pt advance      |
| Title start           | 6.72 pt inline, 9 pt block opening, 5.04 pt closing               |
| Block marker          | Arrow and 0.6 pt bar, 3.36 pt high, from 5.1 pt below the box top |
| Marker stroke         | Filled and stroked in the outline colour, 0.14 pt                 |

The opening block marker is an arrow pointing left, tip 2.34 pt and base 4.5 pt from the
left edge, with the bar from 5.58 pt. The closing marker points right, tip 2.82 pt and
base 4.98 pt from the right edge, with the bar from 6.66 pt. The page's 11 pt opening
tag runs from x 79.224 to 97.704 and y 729.6 to 743.04 on a 731.28 baseline; its
closing tag runs from 127.1 to 136.94 and the next word starts at 138.62.

In the 11 pt Arial tagged probe, COM put `Amet` and the wrapped `in` at (45, 63.5)
and (45, 157.5). The PDF glyph boxes began at (45, 64.05) and (45, 158.15).
The bounding-box probe broke at `luctus` at (45, 133.5) in COM and (45, 133.91) in PDF.

Run the Word export again with `pwsh -File src/core/scripts/docx-layout-reports-word.ps1`.

`generate-line-height-fixtures.py` writes the #391, #393 and #395 documents (`mixed-font-line-height*`, `larger-space-run*`, `symbol-bullet-line-height*`) from the XML those reports quote; the reporter's own files were not available.
