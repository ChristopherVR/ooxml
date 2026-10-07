# Native inline object formatting references

Recorded 2026-10-08 with Word build 16.0.20430 in an owned hidden instance.
The installed licenses are perpetual Office editions, so these cases do not
certify current Microsoft 365 subscription behavior.

`scripts/record-word-review-object-formatting.ps1` creates four synthetic cases:
an inline picture, a footnote reference, a page break and a PAGE field code.
Each has native before, tracked bold, accepted and rejected DOCX files. The
recorder changes no Office profile settings. `reference.json` records every
available main, note, header and footer story's text and revision count.

Component tests compare exported content and direct run properties with Word's
accepted/rejected files. Picture checks retain media bytes, exact drawing XML
and opaque `noProof` properties. Core tests retain independent insertion
history, comments, links and note labels, and cover complete prior properties
and isolated undo/redo. Yjs tests cover peer resolution and shared undo.
Browser checks cover all four cases in each of the six framework bindings,
including Original-mode properties compared with each native before reference.

`src/ui/scripts/write-word-review-inline-exports.mjs <directory> --object-formatting`
creates eight synthetic editor exports. Reopening them read-only with
`scripts/check-word-review-inline-exports.ps1 -IncludeStories` produces
`core-export-reference.json`: all story text matches Word's native results,
with zero revisions in every inspected story.

This establishes the covered review resolution behavior. It does not establish
pixel equality, atom formatting recording, OMML formatting, or full Microsoft
365 Word parity.
