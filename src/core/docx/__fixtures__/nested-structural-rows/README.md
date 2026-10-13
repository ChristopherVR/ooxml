# Native automatic nested row structure

This corpus replaces the row-count whitelist with the existing structural guard:
at least two automatic top-aligned rows, exactly two direct paragraph cells per
row, resolved positive width, complete consistent solid shared edges, and a single
leading nested table in a direct, note-free enclosing story. Every other refusal
remains in force. The existing source-edge, row/token continuation and frame
helpers perform layout; there is no second layout engine.

Reproduce with `scripts/record-word-nested-multi-row.ps1 -Profile structural
-VerifyCore -OutputDirectory <new-temporary-directory>`. The recorder owns and
closes its new Word.Application and every synthetic document. Twenty-two four-
and eight-row cases cover short and long unequal cell stories, row boundaries,
first/middle/last cuts, all-long ten-page flow, one-line rows, and empty first,
middle, last or every row. Native Word 16 version/build, per-page marker sets,
fills, border rectangles and text baselines are in `native-measurements.json`.
The committed Word-saved source fixtures are immutable.

Across 48 native pages, physical fragments repeat the original outer top and
bottom edges, independently of source row count. The source widths are 2 and
0.5 points (native PDF rectangles quantize to 1.92 and 0.48 points). Before this
guard, four eight-row profiles misplaced a body line across physical pages and
had 11.45 to 12.98 point frame errors. The existing reflow now matches all page
counts and per-page body sets, with frame height error at most 0.156 points and
width error at most 0.046 points. Minimal and empty rows retain their layout
occupancy rather than being dropped.

The existing first-ink assertion remains below 0.15 PDF points, using the measured
9.6 point ascent of these 10 point Times New Roman probes. A separate all-line
relative ink bound is below 0.2 points: its measured maximum is 0.176 points,
also present in the unchanged three-row reference corpus. This does not replace
the first-ink assertion or establish glyph/raster parity. All-empty cases have
no inner ink; their frame and enclosing text positions are tested instead.

Forty-four core exports (22 no-op and 22 direct enclosing-text edits) pass 88
native open/save/reopen checks, recorded in `core-reopen-native.json`. Regression
tests check exact package entries and model after no-op export/reload, exact
nested XML and other parts after the direct edit, and source-edge geometry.
Shared DOM tests check every corpus page, noninteractive decoration layers and
exact body sets. Browser coverage includes four-row first cuts, eight-row short,
all-long and minimal cases with deterministic edits, rendered undo/redo,
structural refusal, source preservation and fresh print projection on reopen.

The earlier two- and three-row fixture bytes and historical evidence remain
unchanged. The structural rule is supported by representative 2/3/4/8-row native
corpora, not an exhaustive proof over all possible row counts or Office tables.
Notes, merges, conflicting/incomplete borders, heights, repeated headers,
paragraph keeping or breaks, non-top alignment, unresolved widths and deeper or
multiple nested tables remain outside this admission. General multi-row note
continuation and external note interactions are not established by these probes.

Validation passes 1,545 Word core tests, 64 shared DOM tests and 23 browser
workflows (eight repeated new cases plus the full 15-case two/three/structural
workflow suite). Strict core, full shared UI and focused browser-spec types,
core/source-aliased demo builds, targeted lint and formatting pass.
