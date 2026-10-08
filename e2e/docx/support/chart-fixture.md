# chart.docx provenance

`chart.docx` is copied unchanged from `ooxml-core` (`src/core/docx/__fixtures__/chart.docx`),
where `scripts/docx/build-chart-fixture.mjs` builds it from
`src/core/xlsx/__fixtures__/excel-features.xlsx`: the chart part (an Excel-authored clustered
column chart, "Sales by region", two series of four regions), its relationships, its chart style
and colour style parts and the theme are copied byte for byte. `document.xml`, the relationships
and the content types are hand-built, and there is no embedded workbook, so the chart is drawn from
the values cached in its part. It was not produced by Word, so it shows that the chart part is
read and painted, not that Word renders it the same way.
