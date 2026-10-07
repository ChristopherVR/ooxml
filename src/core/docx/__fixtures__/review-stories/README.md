# Native document-wide review references

Recorded 2026-10-08 in a separate hidden Word instance, build 16.0.20430.
The installed licenses are perpetual Office editions; this does not certify
current Microsoft 365 subscription behavior.

`scripts/record-word-review-stories.ps1` creates one synthetic document with
format revisions in the body, default header, default footer, footnote and
endnote. Both note types intentionally use ID 1. Formatting targets only the
literal text, excluding automatic note marks and paragraph boundaries.
The four DOCX files contain native before, tracked, accepted and rejected
results. The recorder uses Word's document-wide `AcceptAllRevisions` and
`RejectAllRevisions`, and changes no Office profile settings.

`reference.json` records text and revision counts for every available main,
note, header and footer story. Component tests compare saved text and direct
bold/italic properties with the native accepted/rejected files, check zero
pending revisions, and restore the original model with one undo operation.

`src/ui/scripts/write-word-review-inline-exports.mjs <directory> --stories`
creates two editor exports. `scripts/check-word-review-inline-exports.ps1`
with `-IncludeStories` reopens them read-only and writes
`core-export-reference.json`. Both exports retain the native story text and
have zero revisions in all inspected stories. Six browser bindings and Yjs
peers cover document-wide resolution and history. This does not establish
pixel equality, paragraph-mark formatting support, or complete native parity.
