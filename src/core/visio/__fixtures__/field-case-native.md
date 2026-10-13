# Native text-field casing

`scripts/visio-field-case-acceptance.ps1` uses owned Visio 16.0.20430.20146 automation to record
`FIELDPICTURE(38)` and `FIELDPICTURE(39)` output for every scalar value with a case
mapping in the generating Node runtime (Unicode 17.0). The JSON contains 16
chunks, each with input and both native outputs. The separate synthetic VSDX
contains inherited fields using normal, lower and upper case pictures 37-39.
No stencil content or Office implementation was copied.

Visio's field-specific mappings differ from JavaScript Unicode casing. Examples
include unchanged uppercase sharp S, dotted I becoming plain i, non-contextual
Greek sigma, and legacy Georgian mappings. The runtime table encodes observed
one-scalar mappings as sorted ranges; unmatched scalars stay unchanged. It does
not use the browser's potentially different Unicode case tables.

The corpus tests every captured input/output pair and the VSDX regression checks
inherited instance values through parsing and unrelated-edit save/reopen. This
evidence applies to this Visio 16 behavior, not all Office applications, locales,
versions, or numeric/date field pictures. The generator is separate from runtime
code so refreshing evidence requires deliberate review.

To regenerate the compressed table from reviewed evidence, run
`node scripts/visio-field-case-generate.mjs <corpus.json> <output.ts>`, format the
output with oxfmt, and compare it with `src/core/visio/text-field-case.ts`.

Microsoft documents the [VisFieldFormats constants](https://learn.microsoft.com/en-us/office/vba/api/visio.visfieldformats)
and [custom-field API](https://learn.microsoft.com/en-us/office/vba/api/visio.characters.addcustomfieldu).
