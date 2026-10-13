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
values are not comprehensively indexed. Series/category find-and-replace is supported separately for validated literal source cells.

## Chart series and category label transactions

Build core, then run the synthetic label acceptance harness:

```powershell
bun e2e/pptx/native/generate-chart-labels.ts D:/Temp/pptx-chart-label-acceptance
powershell -NoProfile -File e2e/pptx/native/verify-chart-labels.ps1 -ArtifactDirectory D:/Temp/pptx-chart-label-acceptance
```

This opens only generated decks, activates their embedded workbooks in Excel,
checks changed source cells and preserved unrelated cells, refreshes the charts,
and saves and reopens a native copy. Existing PowerPoint or Excel processes cause
the isolated harness to stop before opening anything.

The search scope defaults off and covers series names and category strings,
including grouped charts. Numeric values are excluded. Editing supports direct
literal chart data or direct references to literal cells in a single embedded
`.xlsx` workbook with consistent chart caches. External and shared workbooks,
formula-dependent worksheets, rich worksheet strings, complex/missing ranges,
other chart formula consumers, date/multilevel categories and ChartEx sources
require source editing. All source ranges are validated before any workbook
bytes change; a failed reconciliation stops export instead of returning a
cache-only result. These cases do not establish full chart or Office parity.

## Shape background-only animation acceptance

Build core and the viewer core compatibility entries first. Generate a synthetic PresentationML `p:spTgt/p:bg` fixture with `bun e2e/pptx/native/generate-background-animation.ts output.pptx`, then run `verify-background-animation.ps1 -InputPath output.pptx -OutputDirectory acceptance`. The harness refuses an existing PowerPoint process, opens only its owned fixture, checks its single native effect and foreground text, exports a static PNG, resaves and reopens it, and verifies every saved shape target remains background-only. Native PowerPoint may add a visibility-setting behavior, hence more than one `p:spTgt` for the one effect. This proves package acceptance and preservation, not native slideshow pixel equivalence. The five-binding browser spec separately checks hidden paint, visible foreground text, click playback and exit.

Pass `emphasis` after the output path to generate background fill/stroke colour emphasis. Add `whole-opacity` to include an independent numeric whole-shape opacity ramp; verify that two-effect fixture with `-WholeOpacity`. The browser checks final paint colours, unchanged text colour, and whole-shape opacity. Discrete native opacity sets without CSS animation are not represented as committed opacity in the current playback state and are not covered by this slice.
