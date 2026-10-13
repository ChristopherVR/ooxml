# Native nested table frames

Synthetic fixtures were regenerated and saved by an owned Microsoft Word 16.0 instance,
build 16.0.20430, using `scripts/record-word-nested-table-frames.ps1`. No existing user
Word documents are opened. The four cases cover one row, two differently shaded rows,
a centered 120 point table inside a 160 point cell, and a split 20 line nested row.

Run the PowerShell generator with an output directory, then run
`python scripts/measure-word-nested-table-frames.py <output-directory>` to extract
PDF points with PyMuPDF. `native-measurements.json` records native colored fills,
red border rectangles and text baselines. The native PDFs stay in the output directory.
The DOCX fixtures are native saved OOXML; source styles and note references are retained.

This slice paints resolved nested borders and shading, including physical fragments,
without introducing duplicate text or a second editable nested story. Nested text remains
read-only and enclosing cell structural edits remain guarded. Native borders reserve
approximately one point of text inset and row height, which this projection does not yet
reserve. Exact glyph metrics, vertically merged inner cells and repeated inner headers
are not asserted as native parity. Center alignment uses source grid width, while native
Word also includes the external border extent. The tests deliberately preserve existing
text and note pagination rather than claiming those remaining offsets are resolved.
