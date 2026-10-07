# Native paragraph formatting history

Recorded 2026-10-07 by a separate hidden Word instance, build 16.0.20430
(Click-to-Run 16.0.20430.20140). Installed licenses are perpetual editions;
these fixtures do not certify current M365 subscription behavior.

`scripts/record-word-review-paragraph-formatting.ps1` creates four synthetic
documents covering alignment, paragraph spacing, indentation and combined
changes including keep-with-next. Each case has before, tracked and native
rejected DOCX files. `reference.json` records Word's formatting and revision
type 10 (paragraph property change). Existing documents and Office profile
settings are untouched.

Regression tests cover preserving `w:pPrChange` and its complete prior `w:pPr`
through text edits, editor conversion and export, and core acceptance. Paragraph
format recording and prior formatting in Original display remain unfinished.
Shared editor Review commands support navigation, acceptance, rejection,
undo/redo and Yjs peer synchronization. Core rejection
restores prior properties and matches native rejected documents; later known
edits overlay the complete restored XML basis.

`core-export-reference.json` records Word opening all eight edited exports
(package-preserving and standalone for each case): one paragraph property
revision, edited text, and the expected current paragraph formatting remain.

`core-rejected-reference.json` records Word opening all eight rejected exports:
zero revisions and the same paragraph properties as native rejection.

`core-rejected-properties-reference.json` additionally checks the paragraph-mark
font in all eight rejected exports. Rejection retains the current paragraph-mark
and section properties when the prior paragraph snapshot omits them. The model
and editor retain complete source paragraph properties through text edits and
export, including opaque attributes and independently tracked properties.

# Tracked text export check

`tracked-text-export-reference.json` records desktop Word reopening four core
exports. Each export splits the first run after four characters, inserts a
tracked `!` by Codex, and preserves the original run formatting on all three
pieces. Word reports one paragraph-format revision (10) and one insertion (1),
retains the complex-script Arial font, and Reject All restores the original text
with zero revisions. The export comes from the corresponding tracked fixture via
`loadDocx(...).save(model)`. This does not certify glyph shaping or M365 licensing.
