# Native nested table frames

Synthetic fixtures were regenerated and saved by an owned Microsoft Word 16.0 instance,
build 16.0.20430, using `scripts/record-word-nested-table-frames.ps1`. No existing user
Word documents are opened. The initial four cases cover one row, two differently shaded rows,
a centered 120 point table inside a 160 point cell, and a split 20 line nested row.

Run the PowerShell generator with an output directory, then run
`python scripts/measure-word-nested-table-frames.py <output-directory>` to extract
PDF points with PyMuPDF. `native-measurements.json` records native colored fills,
red border rectangles and text baselines. The native PDFs stay in the output directory.
The DOCX fixtures are native saved OOXML; source styles and note references are retained.

This slice paints resolved nested borders and shading, including physical fragments,
without introducing duplicate text or a second editable nested story. Nested text remains
read-only and enclosing cell structural edits remain guarded. Uniform single-line nested frame borders now reserve measured text insets and row heights using their
actual source widths. Horizontal margins interact with half the border width; vertical
margins add to the top edge, while the final bottom/right edges extend the grid bounds.
Exact glyph metrics, vertically merged inner cells and repeated inner headers remain
outside the measured parity claims. Source XML and structural guards are preserved.

Three exact-height (30 point) cases measure vertical alignment. Native first text baselines
are 50.66, 59.18 and 67.58 points for top, center and bottom respectively. Projected text
shares the painter's vertical offset. Native split fixtures contain top and bottom red
horizontal borders on both physical fragments, supporting border painting at each cut.
Four additional profiles cover 0.5 and 2 point borders, 6 point margins with 1 point borders,
and 6 point margins with 3 point borders. PDF dimensions are rounded by native Word,
so geometry comparisons use a small point tolerance rather than exact PDF coordinates.

Mixed border widths, missing edges and non-solid border styles retain approximate
occupancy and report a layout warning until those profiles have native evidence.

Browser CSS can snap thin border strokes to device pixels; these comparisons assert
source geometry and text placement, not identical rasterization or glyph metrics.

An unequal 20/18-line case records native parallel continuations: 16 lines per cell on
page one, followed by four lines in the first cell and two in the second. Native row
bands contain the longer story and the associated note starts with N0L00 on page one and continues with N0L01/N0L02 on page two. A final
outer paragraph can mask the shared-layout last-box height defect in parsed OOXML;
the direct split regression checks both paragraph orders, including margins and frame
edges, independently of that trailing paragraph.
