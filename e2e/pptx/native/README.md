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

Pass `emphasis` after the output path to generate background fill/stroke colour emphasis. Add `whole-opacity` to include an independent numeric whole-shape opacity ramp; verify that two-effect fixture with `-WholeOpacity`. The browser checks final paint colours, unchanged text colour, and whole-shape opacity. Background-only numeric sets are refused by the fixed-opacity resolver because native rendered composition differs from a simple paint/body CSS multiplication.

## Whole-shape fixed opacity acceptance

Build core and the viewer compatibility entries, then generate `hold.pptx` and `remove.pptx` with `bun e2e/pptx/native/generate-opacity-sets.ts output-directory`. Run `render-opacity-set.ps1 -InputPath output-directory/hold.pptx -OutputDirectory acceptance -FfmpegPath path/to/ffmpeg` (and repeat for remove in a separate output directory). The harness refuses an existing PowerPoint process, owns its application and presentation, checks the effect and black foreground text, resaves the package and renders native video frames. Video export includes a startup offset; inspect active and completed frames without equating timestamps to the authored clock. Package save/reopen preservation is tested separately.

Native PowerPoint 16 rendered an isolated whole-shape numeric-string set as a discrete opacity change. Missing effect/behavior fill restores the original fill and text; explicit hold on both retains the transparent fill and text. Explicit remove on either restores. Playback supports finite isolated whole-shape numeric sets with simple delay, including plain decimal numeric strings and the native legacy image hint. Scientific-notation strings are refused; finite numeric `fltVal` variants are accepted. Background, text-subtarget and hover sets, repeated/reversed/additive or mixed sets, event-driven child timing and indefinite intervals are refused and preserved. The timeline reports unsupported opacity profiles instead of substituting a transparency pulse.

## Imported whole-shape after-effect clearing

Build core, shared UI and the viewer compatibility entry first. Run `bun e2e/pptx/native/generate-after-effect-clears.ts output` and `verify-after-effect-clears.ps1 -Directory output -Render`. The synthetic native fixture has four effects: RGB dim, scheme dim, hide after animation and hide on next click. Four edited cases clear only the selected whole-shape action, including the scheme-color dim action. Native checks compare effect counts and after-state values, resave/reopen, media hashes and audio relationships. The owned PowerPoint process is closed in finally; an existing user process is refused.

For embedded sound coverage, run `generate-after-effect-sound.ps1 -SourcePath e2e/pptx/fixtures/animation-after-effect.pptx -OutputDirectory sound-source`, then pass its native-after-effect-with-sound.pptx as the generator's third argument in a fresh output directory. The sound fixture uses a generated silence WAV imported through Microsoft's [SoundEffect.ImportFromFile](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.soundeffect.importfromfile) API.

Native PowerPoint 16 video samples at 10 seconds show the source RGB-dim shape gray, while the cleared shape retains its blue fill and white foreground text. This measured clearing case does not establish general dim-to-color rendering fidelity. Core regressions separately verify exact no-op timing XML, explicit none persistence, unrelated retiming, sound bytes/relationships and opaque mixed sub-timing. The browser spec drives imported controls, undo/redo, export/reload and playback in all five demos.

Imported controls are bounded to a single whole-shape entrance/emphasis effect, matched by native node ID and shape target. Ambiguous multiple effects, background/text/subshape targets and unknown after-state behavior remain preserved and excluded from the imported control. Undefined means no edit; explicit none clears recognized after-state nodes without rewriting the entrance effect or removing sound/other sub-timing. An unresolved imported theme color offers action clearing but no invented RGB swatch. The browser fixture adds a separate following-click shape to exercise the hide-on-next-click boundary. Imported entrance effects without modeled duration currently use the viewer's existing 500 ms fade fallback; these checks do not establish native entrance timing fidelity.
