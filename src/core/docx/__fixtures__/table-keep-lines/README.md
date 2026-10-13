# Native ordinary-cell paragraph keeping

These immutable Word 16 compatibility-mode 15 sources measure Keep lines together
in one automatic, top-aligned, two-cell row. Each of nine profiles has a keep-on
source and an otherwise identical keep-off control: 18 documents and 45 pages.
The corpus includes keeping after a prefix, row startup on an occupied page,
oversized paragraphs, already consumed prefixes, and mirrored cell placement.

Native Word defers the entire starting row when either first kept paragraph does
not fit the occupied page. Within a row, an unstarted kept paragraph and its
following cell content defer independently while the other cell continues.
An oversized paragraph first defers when it follows content, then splits at a
fresh origin. A partially consumed suffix continues without keeping again.

The implementation uses existing paragraph boxes, source token offsets,
paragraph cuts, row reflow and PageCursor startup. Admission is limited to this
ordinary, note-free, single-column profile with direct paragraphs, exact positive
line spacing, zero paragraph spacing, disabled widow control and complete simple
cell borders. Nested content, merges, explicit heights (including zero), headers,
keep-next, page breaks, notes, floating content and complex frames remain outside
this admission. No font size or line-count constant controls keeping.

Core regressions compare exact page/body marker sets and frame extents within
0.2 pt. First ink remains within 0.15 pt; the separate all-line ink bound is
0.2 pt. These are source layout measurements, not glyph or raster parity.
Shared DOM and browser checks cover physical pages, independent content,
reachable paragraph controls, history and fresh print projection after reopening.

No-op export preserves every package entry byte-for-byte and the complete parsed
model. Edited paragraph export preserves all other package parts and all keep and
pagination properties. Exact edited XML comparison explicitly normalizes only
existing CRLF declaration normalization, a redundant WordprocessingML namespace
on the edited paragraph and its stale lastRenderedPageBreak cache. Removing that
cache can change the derived structural-editability diagnostic. This is semantic
edited-paragraph preservation, not byte-exact edited XML preservation.

Reproduce on Windows with an installed Word application, Python with PyMuPDF and
Bun, from the repository root:

```powershell
./scripts/record-word-table-keep-lines.ps1 -OutputDirectory "$env:TEMP/word-keep-probes" -VerifyCore
```

The recorder owns a new empty invisible Word application, closes each synthetic
document and releases its application. native-measurements.json records the PDF
geometry and native version. core-reopen-native.json records 72 native open,
save and reopen checks of 36 no-op and edited core exports. The generator reuses
the existing ordinary-table border fixture; sources are never regenerated during
normal tests.
