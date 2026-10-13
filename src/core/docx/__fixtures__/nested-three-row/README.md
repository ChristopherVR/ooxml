# Native three-row nested table fragments

This corpus extends the measured two-row source-edge profile using the same native
generator, recorder, shared core row/token continuation and frame helpers. Reproduce
with `scripts/record-word-nested-multi-row.ps1 -Profile three -VerifyCore
-OutputDirectory <new-temporary-directory>`. The recorder owns one new Word.Application
instance and closes every synthetic document and its application.

Six profiles contain three automatic, top-aligned nested rows, two direct paragraph
cells per row, complete consistent solid edges, resolved width and direct note-free
enclosing text. Short and row-boundary cases plus first-, middle- and last-row cuts
include an all-long four-page case. Source fixtures and two-row fixtures are immutable
once committed. Native PDF text baselines, fills, red inner borders, page counts and
Word version/build are recorded in `native-measurements.json`.

Before admission, the shared engine already matched all per-page body line sets and
page counts but retained interior 1.5 point edges at some physical cuts and starts.
Native Word repeats the original 2 point top and 0.5 point bottom on every fragment.
Extending the existing bounded guard to three rows and reading the last source row's
bottom edge fixes roughly 0.57 to 0.9 point frame errors, without another layout engine.
All measured frame heights now differ by less than 0.2 PDF points and widths by less
than 0.1 point; relative projected ink baselines retain the existing 0.15 point rounding
bound and 9.6 point reference ascent for these 10 point Times New Roman probes.
These assertions do not establish glyph/raster parity.

Twelve core exports (six no-op and six direct text edits) pass twenty-four native
open/save/reopen checks, recorded in `core-reopen-native.json`. No-op exports preserve
all package entries, source XML and model. Direct enclosing-text edits preserve nested
XML and other package parts. The existing shared DOM/browser workflows verify geometry,
source preservation, deterministic paragraph-start edits, rendered undo/redo, structural
edit refusal and old editor/page detachment before reopening.

Notes, merged/conflicting borders, heights, headers, paragraph keeping, non-top cell
alignment, unresolved widths and deeper/multiple/four-or-more-row tables remain outside
admission. The one-row note continuation profile is unchanged; this corpus does not
establish multi-row note continuation, external note interactions or general table parity.

Validation passes 1,480 Word core tests, 42 shared DOM tests and twenty-eight
repeated browser cases: twenty-two source/geometry/history/reopen workflows across
both corpora and six existing nested-note controls. Strict core, full shared UI
and focused browser-spec types, source builds, lint and formatting pass.
