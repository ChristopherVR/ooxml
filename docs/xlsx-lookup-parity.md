# Excel lookup parity evidence

The target is current Microsoft 365 Excel, as requested on 7 October 2026.
These changes establish a measured lookup subset, not whole-product parity.
All calculation changes live in `src/core/xlsx/formula` and are inherited by
the shared editor and every viewer binding. No viewer-specific search engine
or new logic package was introduced.

## Implemented

- Exact XMATCH/XLOOKUP searches retain the full reference extent and can find
  trailing blank cells. Whole-column searches do not scan the implied tail.
  Blank references differ from explicit zero and empty-string values.
- MATCH/VLOOKUP/HLOOKUP reject scalar table arguments while retaining real
  one-cell references and arrays, with native coercion and error propagation.
- Linear approximate XMATCH/XLOOKUP searches compare across value types,
  ordered as numbers, text, logicals, then blanks. Search direction determines
  ties and exact blank matches without materializing the reference tail.
- Binary searches include sorted trailing blanks and choose native duplicate
  positions: first in ascending data, last in descending data.
- An explicitly omitted XMATCH search mode defaults to forward search. An
  empty referenced cell still coerces to zero and produces `#VALUE!`. The
  shared function-call default policy preserves that distinction before
  parameter lifting, without changing other functions' optional arguments.
- ADDRESS preserves omitted absolute/style defaults while coercing empty
  references normally. INDIRECT treats an explicitly empty style slot as
  R1C1 mode. OFFSET retains source dimensions only for omitted height/width
  slots; empty referenced dimensions produce `#REF!`. Computed defaults reuse
  the same call policy and derive their values from the source reference.
- INDIRECT now supports whole-row/column R1C1 references, signed relative
  coordinates and native wrapping at sheet edges. The exported
  `parseR1C1Range` address helper is independent of formula evaluation and
  reuses the existing grid limits and range normalization. Product sheet
  resolution and sparse aggregation remain in their existing core paths.

## Reproduce the native evidence

Run `scripts/record-xlsx-lookups.ps1 -OutputFile <temporary-json-path>` on
Windows with Excel installed. It creates its own hidden workbook, closes it
without saving, and records the installed version/build with 216 comparisons.
The committed fixture is
`src/core/xlsx/formula/__fixtures__/lookup-native.json`, replayed by
`lookup-native.test.ts`. It covers both sort orders, all implemented match
modes, forward/reverse linear searches, correctly sorted binary searches,
numbers, text, logicals, blanks and wildcard lookup values.

Additional focused tests record scalar argument rejection, omitted slots,
duplicate positions, blank/empty-string distinctions and bounded whole-column
reads. The recorded application was Excel 16.0 build 20430.

## Validation and remaining limits

- Final integration validation at `2cd583316` passed all 4,326 XLSX tests
  across 126 files, including the 216 native comparisons. Both core TypeScript
  projects, the complete core package build and clean-consumer imports of
  every ESM/CJS entry point passed.
- A generated seven-formula workbook was saved by the library, opened,
  recalculated and saved by Excel, then reloaded and recalculated by the
  library. All seven values, including errors, remained identical. This is
  bounded save acceptance evidence, not a general lossless-export claim.
- Existing formula regression cases continue to pass, including classic
  approximate lookup behavior, array lifting and return references.
- After adding reference-function fixes, 2,369 formula tests and both core
  TypeScript projects passed. The parser suite and sheet-qualified aggregation
  checks passed, followed by a save/reload/recalculation regression for a
  wrapped R1C1 reference (26 reference tests passed).
- An earlier broad rerun hit suite-load failures during severe shared-machine
  memory pressure and was stopped. The native TypeScript declaration emitter
  also failed to create temporary directories (`TS5033`). After memory recovered,
  the full suite and unchanged build passed when run separately. No build-script
  changes were required.
- Binary search requires the declared sort order. The corpus does not
  establish arbitrary unsorted-input behavior or every error placement.
- Current Microsoft 365 regex functions need further work. Microsoft
  documents all-match and capture-group array modes for REGEXEXTRACT and
  PCRE2 semantics; the current implementation supports first-match mode
  using JavaScript regular expressions. The installed native application
  returned `#NAME?` for REGEXEXTRACT, so it cannot validate that feature.
- Grid/print fidelity, metadata preservation, pivot/data features and
  spreadsheet Yjs mapping remain separate unfinished work in the
  [broader parity review](xlsx-parity-review.md).

Primary references:
[XMATCH](https://support.microsoft.com/en-gb/excel/functions/xmatch-function?nochrome=true),
[XLOOKUP](https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929),
[REGEXEXTRACT](https://support.microsoft.com/en-us/excel/functions/regexextract-function).
