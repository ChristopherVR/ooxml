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

`--record-object-formatting` instead starts from the native before files and
uses the shared core toggle and tracking commands to produce twelve tracked,
accepted and rejected exports. `recorded-export-reference.json` records Word's
revision counts and direct object bold properties. `recorded-rejection-reference.json`
uses Word's document-wide Reject All in memory before inspecting the same files,
without saving them: recorded changes clear and prior direct bold is restored.
Core tests compare native resolved content and preserve the earliest snapshot
through successive edits and full reversion. Yjs peers cover author identity,
convergence and one-step undo/redo; mixed text/object ribbon selections work
across all six browser bindings. Hidden field-code formatting uses a direct
command because the editing field guard trims hidden markers from selections.

This establishes the covered formatting and review behavior. It does not establish
pixel equality, mixed structural tracking, all advanced atom controls, OMML formatting,
hard-break formatting in Yjs, or full Microsoft
365 Word parity.
