# Native Excel chart gradient presets

`excel-gradient-presets.json` was captured with Excel 16.0 build 20430 by
`scripts/record-xlsx-gradient-presets.ps1`. Each case applies
`FillFormat.PresetGradient(1, 1, id)`, saves a separate workbook, extracts its
chart XML and reopens it to record the native preset ID, angle and stops.
The corpus covers the 24 standard named Office presets, horizontal style 1,
variant 1. It does not measure theme-dependent presets or other variants.

Production definitions use exact saved OOXML stop positions and sRGB colors.
COM floating-point position getters are evidence, not a reason to round the
saved positions. The fixtures contain no native workbook binary or rendered
image. For the separate native pixel corpus and known paint discrepancies,
see `../../chart/__fixtures__/README.md`.

## chart-surface-band-formats.xml

A surface chart part with `c:bandFmts` (content the shared chart model does not hold), copied from
`src/core/pptx/__tests__/fixtures/e2e/three-d-parity/three-d-charts.pptx` (`ppt/charts/chart16.xml`)
with its `c:externalData` removed. `xlsx/edit/chart-space-blocked.test.ts` and the Excel acceptance
script use it to check that such a chart is saved from its source bytes and never rewritten.
