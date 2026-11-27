# Font subsets for the PDF export tests

Latin subsets (U+0020 to U+007E and U+00A0, no hinting, with their `kern` and `GPOS` kerning and
their vertical metrics) of free fonts, made with fontTools `pyftsubset` from the files installed
with them. The tests supply them as the families the documents name: `FixtureMetricSans` as
`Carlito` (renamed in its name table, because the Open Font License reserves the name "Carlito" for
the unmodified font), the others under their own names. No Microsoft font is included.

They are licensed under the SIL Open Font License 1.1 (`OFL.txt`, with each font's copyright).
