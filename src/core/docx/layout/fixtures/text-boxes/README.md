# Word-saved text boxes

`word-saved.docx` is `textBoxLayoutFixture()` (`src/core/docx/test-support/text-box-layout-fixture.ts`)
opened in a hidden instance of Microsoft Word 16 (build 16.0.20527) through COM, exported to PDF
and saved again with `SaveAs2` (format 16). Only the author name in `docProps/core.xml` was replaced
afterwards. Word wrote its own markup around the boxes: `mc:AlternateContent` with a VML fallback, and
the linked chain renumbered to `wps:txbx id="7"` / `wps:linkedTxbx id="7" seq="1"`. It kept every
`wp:extent` as authored, including the too-small height of the `spAutoFit` box.

Arial 11 pt, single spacing, no paragraph spacing; every box floats in front of the text at a fixed
page offset, with a 0.75 pt outline and the default insets. What Word's PDF shows (positions in CSS
pixels from the page's top-left, read with PyMuPDF):

| Box             | Settings                       | Word 16                                                                                                                               |
| --------------- | ------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------- |
| `Wrapped`       | 2 x 1 in, `noAutofit`          | Lines `Wrapped word1 word2`, `word3 word4 word5`, `word6 word7 word8`, `word9 word10 word11`, `word12 word13 word14`; the rest hidden |
| `Columns`       | 3 x 1.2 in, `numCol="2"`       | One column across the box: `Columns word1 word2 word3 word4`, `word5` to `word9`, `word10` to `word14`; the seventh line cut          |
| `Vertical`      | 0.9 x 2.4 in, `vert="vert"`    | Turned clockwise, first line at the right: `Vertical word1 word2 word3`, `word4` to `word8`, `word9` to `word12`                      |
| `Vertical 270`  | 0.9 x 2.4 in, `vert="vert270"` | Turned anticlockwise, first line at the left, reading upwards; the same breaks                                                        |
| `Autofit`       | 2.5 x 0.4 in, `spAutoFit`      | Drawn 83.8 pt high (six lines of 12.65 pt, the insets and the outline), breaks `word4`, `word9`, `word13`, `word17`, `word21`         |
| `No wrap`       | 1.5 x 0.4 in, `wrap="none"`    | Drawn 200.4 pt wide: one line, left edge kept                                                                                         |
| `Linked first`  | 2 x 0.8 in, `wps:txbx id`      | Three lines, to `word8`: a fourth would end 0.2 pt below the text rectangle                                                           |
| `Linked second` | 2.5 x 1.2 in, `wps:linkedTxbx` | From `word9`, four words a line; the seventh line cut at the text rectangle, the rest hidden                                          |
| `Shrink`        | 2 x 0.6 in, `normAutofit`      | Neither resized nor shrunk: three lines, the rest hidden                                                                              |
| `Formatted`     | 2.5 x 1.2 in, `anchor="ctr"`   | Right-aligned bold 14 pt heading over a 8 pt italic paragraph with a bottom border, centred in the box                                |

Probes of the same generator (not kept) showed: the text rectangle is the preset's (`ellipse`: inset by
(1 - cos 45°) / 2 of the size) less the insets and half the outline (a 6 pt outline wraps a line that
fits without one); `spAutoFit` also shrinks a box to one line; a column break in a box is drawn as a
line break; a three-box chain flows in `seq` order; `vert` with `anchor="ctr"` or `"b"` centres or
moves the lines across the box's width.
