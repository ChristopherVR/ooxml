# Word layout issue fixtures

Borris-pro attached the original generated samples to the issues below. `cell-end-marker.docx` is a generated follow-up for #330. The anonymized production documents were used locally for comparison.

Word 16.0.20625 exported each fixture through `Document.ExportAsFixedFormat`. The fixture tests exercise the same DOCX import and layout engine as Print Layout and PDF export.

| Issues                                                                                                               | Word observations                                                                                                                                                                                                                                                                                                                                           |
| -------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [#329](https://github.com/ChristopherVR/ooxml/issues/329)                                                            | Three 0.25 pt outlines and a label sized independently of the legacy group coordinates.                                                                                                                                                                                                                                                                     |
| [#330](https://github.com/ChristopherVR/ooxml/issues/330)                                                            | With the current Normal template, both styleless 12 pt samples have seven pages and 43 lines on page one. With an explicit styles part and single spacing, Times New Roman advances 13.8 pt per line and has 50 lines on page one. At 12 pt, Word modes 11, 12 and 14 required 286 pt to fit the test phrase on a justified line. Mode 15 fit it at 279 pt. |
| [#331](https://github.com/ChristopherVR/ooxml/issues/331), [#332](https://github.com/ChristopherVR/ooxml/issues/332) | Both tall-footer samples have three pages. The bordered footer is indented and has one outer box with continuous side borders.                                                                                                                                                                                                                              |
| [#333](https://github.com/ChristopherVR/ooxml/issues/333)                                                            | The picture starts at x=420.75 pt in the header and x=421 pt in the body control.                                                                                                                                                                                                                                                                           |
| [#334](https://github.com/ChristopherVR/ooxml/issues/334)                                                            | Triple borders have three strips. The thin/thick variants preserve the number, order and relative widths of the strips.                                                                                                                                                                                                                                     |
| [#335](https://github.com/ChristopherVR/ooxml/issues/335), [#336](https://github.com/ChristopherVR/ooxml/issues/336) | The signature keeps its full height with the preceding text readable. The wrapped line ends near the left edge of the picture while retaining right alignment.                                                                                                                                                                                              |
| [#337](https://github.com/ChristopherVR/ooxml/issues/337), [#338](https://github.com/ChristopherVR/ooxml/issues/338) | The missing font is substituted with Cambria. The omitted and explicit 10 pt controls use the same size.                                                                                                                                                                                                                                                    |
| [#339](https://github.com/ChristopherVR/ooxml/issues/339), [#340](https://github.com/ChristopherVR/ooxml/issues/340) | The percentage shades export as solid grays. Highlight exports in yellow and small capitals use reduced glyphs for the lower-case part of the text.                                                                                                                                                                                                         |
| [#341](https://github.com/ChristopherVR/ooxml/issues/341), [#342](https://github.com/ChristopherVR/ooxml/issues/342) | The uncached merge field has no displayed result. Footer tabs advance to both the document's default intervals and the explicit 5 cm stop.                                                                                                                                                                                                                  |
| [#343](https://github.com/ChristopherVR/ooxml/issues/343), [#344](https://github.com/ChristopherVR/ooxml/issues/344) | The two groups of empty lines use their respective 2 pt and 8 pt paragraph marks. The hanging-indent border extends 284 twips toward the start margin.                                                                                                                                                                                                      |
| [#346](https://github.com/ChristopherVR/ooxml/issues/346)                                                            | Mode 15 retains all three reported final words and uses two pages. Mode 14 wraps them and uses three pages.                                                                                                                                                                                                                                                 |
| [#347](https://github.com/ChristopherVR/ooxml/issues/347)                                                            | Implicit and explicit Normal suppress the same gaps, including those beside empty paragraphs. Both samples fit on one page.                                                                                                                                                                                                                                 |
| [#348](https://github.com/ChristopherVR/ooxml/issues/348)                                                            | Zero character scale uses the same text widths and two pages as the unscaled control.                                                                                                                                                                                                                                                                       |
| [#349](https://github.com/ChristopherVR/ooxml/issues/349)                                                            | Inline title and closing tags occupy space in both dimensions. The tagged file occupies three pages and the default-appearance control occupies two.                                                                                                                                                                                                        |
| [#353](https://github.com/ChristopherVR/ooxml/issues/353)                                                            | TableGrid supplies single spacing with zero after-spacing inside its cell. The table matches the height of the direct-formatting control and fits on one page.                                                                                                                                                                                              |

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

Run the Word export again with `pwsh -File src/core/scripts/docx-layout-reports-word.ps1`.
