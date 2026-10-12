# Native slide clipboard acceptance

This optional Windows check requires installed Microsoft PowerPoint and no existing
PowerPoint process. It generates synthetic presentations, opens them through COM,
checks slide order, the source picture, bold speaker notes and independent theme
colors, saves a native copy and reopens it. A second scenario uses boundary IDs and
master parts with nonstandard names and directories.

The ID limits follow Microsoft's [Open XML SDK PresentationML schema](https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/schemas_openxmlformats_org_presentationml_2006_main.json):
slide IDs are 256 through 2147483647; master and layout IDs are 2147483648 through 4294967295.

Build core first, then run from the repository root:

```powershell
bun run build
bun e2e/pptx/native/generate-slide-clipboard.ts D:/Temp/pptx-native-acceptance
powershell -NoProfile -File e2e/pptx/native/verify-slide-clipboard.ps1 -ArtifactDirectory D:/Temp/pptx-native-acceptance
```

The script writes generated decks and `native-acceptance.json` into the supplied
directory. It opens only those decks and closes its own presentations and application.
Keep artifacts outside the tracked fixture tree. This checks package transplantation,
not PowerPoint's private clipboard formats or full Office parity.

## Chart title search and replace

The title check uses a synthetic classic chart with an embedded workbook and rich
title formatting. It verifies source, replacement, core resave and native resave,
checking title text, bold formatting, series name, categories and numeric values.
A linked title remains searchable and is rejected by the replacement operation.

```powershell
bun e2e/pptx/native/generate-chart-title.ts D:/Temp/pptx-chart-title-acceptance
powershell -NoProfile -File e2e/pptx/native/verify-chart-title.ps1 -ArtifactDirectory D:/Temp/pptx-chart-title-acceptance
```

Generated artifacts and `chart-title-acceptance.json` stay in the supplied directory.
Linked and ChartEx cached titles and modeled text runs in complex titles remain
searchable, but replacement requires PowerPoint or source-workbook editing. Field
values are not comprehensively indexed. Series/category find-and-replace is not implemented.
