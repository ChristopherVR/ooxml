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
snapshot through text edits, model conversion and export. Editor recording,
navigation and rejection of formatting changes still require implementation.
