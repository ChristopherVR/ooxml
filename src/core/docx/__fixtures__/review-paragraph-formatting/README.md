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
format recording, editor review commands and rejection remain unfinished.

`core-export-reference.json` records Word opening all eight edited exports
(package-preserving and standalone for each case): one paragraph property
revision, edited text, and the expected current paragraph formatting remain.
