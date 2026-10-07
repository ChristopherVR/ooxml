# Native inline object revisions

Recorded 2026-10-08 with a separate hidden Word instance, build 16.0.20430.
Installed licenses are perpetual Office editions. These references do not
certify current Microsoft 365 subscription behavior.

`scripts/record-word-review-inline.ps1` creates synthetic picture, footnote and
page-break insertion/deletion documents. Each case includes its native before,
tracked, accepted and rejected DOCX. `reference.json` records revision types,
text, picture/footnote counts and page counts. The recorder changes no Office
profile settings and closes only its own instance.

Shared editor regression tests compare body object/text content with native
accept/reject results for both picture cases, both note cases and page-break
deletion. This does not establish style or pixel equality, orphan-note cleanup,
or resolution of every related non-body revision. The insertion of a page break
also inserted two tracked paragraph marks in this Word build; the fixture is
retained for the separate paragraph-mark implementation, not claimed as covered.

`src/ui/scripts/write-word-review-inline-exports.mjs` writes the ten covered
editor exports using the shared adapter and commands. The separate hidden
`scripts/check-word-review-inline-exports.ps1` reopens them read-only.
`core-export-reference.json` records matching body text, pictures, footnote counts
and pagination, with zero body revisions. Retained notes still have one pending
revision in their own story; body acceptance/rejection does not yet resolve it.
Regression exports preserve the native picture bytes and `w:noProof` on retained
pictures, remove it with removed pictures, and keep it off neighboring text.
