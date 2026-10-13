# Native startup fit boundaries

This compact Word 16 compatibility-mode 15 record checks the existing keep-lines
startup calculation at its border and margin fit threshold. It adds no keeping
logic and leaves the preceding native corpus unchanged.

Twenty-four sources pair keep-on and keep-off controls at minimum minus 1 pt,
exact equality, and minimum plus 1 pt. Profiles use 3 pt borders in either kept
cell, 6 pt borders with 2 pt top/bottom margins, and 1.5 pt borders with 1.5 pt
margins. The six-line kept paragraph needs 72 pt of content. Both border widths
and both cell margins are already included in the startup minimum: 78 or 88 pt.
Below that threshold Word moves the whole row; at equality it starts both cells
on the occupied page. The existing core matches every per-page marker set.

native-boundaries.json records page/body sets, physical frame bounds, first ink
baseline, native version and hashes of measured input and accepted source files.
The original PDF/DOCX outputs are reproducible temporary evidence, rather than a
second committed copy of the source corpus. Tests construct the inputs from the
immutable table-keep-lines fresh and mirrored-fresh fixtures and compare the
recorded native results. Frame and first-ink tolerances remain 0.2 and 0.15 pt.
No glyph or raster parity is claimed. All no-op package parts and parsed models
are compared separately for exact preservation.

From the repository root on Windows with Word and PyMuPDF installed:

```powershell
./scripts/record-word-table-keep-boundaries.ps1 -OutputDirectory "$env:TEMP/word-keep-boundaries"
```

The recorder creates and closes its own empty invisible Word application. The
generator parameters and raw native measurements remain in the selected output
directory. Accepted package hashes can change after a different native save;
page/body/geometry comparisons establish the measured behavior.
