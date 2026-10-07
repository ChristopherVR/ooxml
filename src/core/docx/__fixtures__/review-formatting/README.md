# Native run formatting changes

Recorded 2026-10-07 in a separate hidden Word instance, build 16.0.20430
(Click-to-Run 16.0.20430.20140). Installed Office licenses are perpetual editions;
these fixtures do not certify current M365 subscription behavior.

`scripts/record-word-review-formatting.ps1` creates two synthetic documents:
one changes bold, and one changes bold, italic, size and color. Each case has
the before, tracked and Word-rejected DOCX. `reference.json` records native
formatting and revision type 3 (property change). No existing documents or
Office profile settings are changed.

Regression tests first cover preserving `w:rPrChange` and its original `w:rPr`
snapshot through text edits, model conversion and export. Editor navigation,
acceptance and rejection now use the same core mapping, with undo/redo and Yjs
peer coverage. New formatting-change recording and prior formatting in Original
display still require implementation.
Core rejection now restores the complete prior properties. The installed Word
build opened four core-rejected exports (both package-preserving and standalone
for both cases), reporting zero revisions and the same fonts as native rejection.
`core-rejected-reference.json` records this additional interoperability check.

`overlapping-text-export-reference.json` records Word reopening four core
exports: each source tracked fixture is split after four characters, then receives
an inserted `!` or a deletion of `at`. The text revision is authored by Codex,
while each piece retains the source formatting revision and its prior properties.
The writer preserves `w16du:dateUtc` separately from `w:date` and renumbers the
new text revision ID. Word recognizes one formatting change and one insertion or
deletion, then Reject All restores `Format me`, baseline bold/color and Arial for
complex scripts, with zero revisions. The adjacent overlap regression tests
reproduce the core exports and compare rejection with the native rejected fixture.
