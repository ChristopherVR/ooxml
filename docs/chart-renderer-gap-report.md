# Chart renderer gap report: PowerPoint renderer vs core's neutral renderer

Goal (decided by the maintainer): one chart renderer, in `src/core/chart/render/`, that xlsx, Visio,
docx and pptx all draw through. This report is step 1: what the pptx renderer draws that core does
not, how big each piece is, and how hard it is to retarget. Step 2 ports the top gaps in slices (see
"Slices" at the end for what has landed).

## How the numbers were made

- `src/core/pptx/__tests__/integration/chart-renderer-gap.test.ts` (kept; runs with the core tests).
  It loads every chart part of the 23 committed decks (82 parts: 75 classic, 7 ChartEx), and for
  each runs core's own pipeline: `parseChartSpace`, `readChartFormatting` (the feed Word uses),
  `chartSummaryFromSpace`, `chartSummaryView`, `renderChartSvg`; ChartEx parts go through
  `chartExSummary`. It asserts nothing throws and that all 82 parts are reached.
  Since 2026-10-23 `chartExSummary` builds its summary from the neutral ChartEx model
  (`chart/ex`, `chart/render/summary-from-chartex.ts`), and the agnostic plan records that its outputs did not
  change; the tables below were measured on 2026-10-17 and were not re-run.
- Feature use is detected by regular expressions on the part text. "Core draws it" is measured by
  ablation: the feature is stripped from the part, the chart is rendered again, and if the SVG is
  identical core does not draw the feature. Set `CHART_GAP_REPORT=<file>` for the JSON table and
  `CHART_GAP_SVG_DIR=<folder>` to keep every SVG.
- pptx code sizes come from counting the non-test `chart-*.ts` files of `src/ui/src/pptx/render/`
  (322 files with tests; about 27,800 lines of source at the top level plus 5,000 in `chart-format/`)
  grouped by feature; a "pptx-type" file imports `PptxChartData` and friends from `ooxml-core/pptx`.

## What the harness found

**Baseline, measured 2026-10-17 before the slices below.** Rows such as data labels (now drawn on 27 of 27 parts, slice 1), markers and always-drawn major gridlines (slice 3) have since been ported, and later work covered radar axis scaling and many-point series; the harness test is still in the tree and the counts below have not been re-run. Read the table as the starting point, not the current state.

No part throws. Core renders 78 of 82 parts to a chart; funnel, box and whisker and region map parts
(and any pareto) have no summary and draw nothing (3 of the 7 ChartEx fixtures).

| Feature (fixtures: parts / decks)                                                                                                                                                              | Core draws it?                                                              | Measured by                                    |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------- | ---------------------------------------------- |
| Data labels, any kind (27 / 6)                                                                                                                                                                 | No: 0 of 27                                                                 | ablation                                       |
| Data label positions `c:dLblPos` (22 / 3), per-point `c:dLbl` (4 / 2), leader lines and callouts (29 / 7)                                                                                      | No                                                                          | follows from the line above (no label code)    |
| Major gridlines (32 / 11)                                                                                                                                                                      | Always drawn, whether or not the part has them (identical with and without) | ablation                                       |
| Markers (6 / 5)                                                                                                                                                                                | No: a fixed dot, symbol, size and fill ignored                              | ablation                                       |
| Manual layout (10 / 4)                                                                                                                                                                         | Title and legend only; plot area and label layouts ignored                  | ablation                                       |
| `c:txPr` text properties (52 / 12)                                                                                                                                                             | Partly: 34 of 52 change the output                                          | ablation                                       |
| 3-D view (20 / 4)                                                                                                                                                                              | No: drawn as flat 2-D                                                       | ablation                                       |
| Explosion (1 / 1)                                                                                                                                                                              | No                                                                          | ablation                                       |
| Trendlines (1 / 1)                                                                                                                                                                             | Yes                                                                         | ablation                                       |
| Picture fills (2 / 2)                                                                                                                                                                          | No                                                                          | code reading (no `blip` handling in `render/`) |
| Surface (3 parts)                                                                                                                                                                              | Labelled frame only                                                         | code reading                                   |
| Secondary value axis (4 / 2), combo (1 / 1)                                                                                                                                                    | Partly (a second value axis and extra groups exist)                         | not ablated                                    |
| Up/down bars, high-low lines (1 / 1)                                                                                                                                                           | Yes for stock                                                               | code reading                                   |
| Data table, date axis, multi-level categories, gradient fills, pattern fills, axis titles, log scale, display units, minor gridlines, error bars, of-pie, rounded corners, chart title overlay | No fixture uses them, so unmeasured                                         | none                                           |

Things seen on the rendered SVGs (Playwright screenshot of 8 charts through core): the line chart
with data labels shows none; a pie chart with labels shows none; marker styling and sizes are lost
(stacked-line-markers draws plain dots); its legend text is drawn much larger than the axis text and
collides with the category labels; transparent chart areas show the page behind them. Not compared
side by side with the pptx renderer: no automated pptx-vs-core image comparison exists.

## Ranked gaps

Ranking weighs fixtures that use the feature, how visibly the chart is wrong without it, and how
cheap the port is. "Pure" means no DOM and no three.js. "Retarget" is the cost of moving from
`PptxChartData` onto core's `ChartSpace` / summary / view model.

| #   | Feature                                                                                                                                                                                                          | pptx code (files / lines)                       | Pure?                                          | pptx-type files | Retarget                                                                                                                                                                       | Notes                                                                                         |
| --- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------- | ---------------------------------------------- | --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------- |
| 1   | Data labels: content cascade (point, series, group), separators, number format, percent, category, series name, custom text, font; placement (`ctr`, `inEnd`, `outEnd`, `inBase`, `t`, `b`, `l`, `r`, `bestFit`) | 8 / 1,243 (+ pie labels 2 / 533)                | Yes (1 file touches the DOM helper types only) | 6 of 8          | Medium. The cascade reads `PptxChartSeries.dataLabelOptions`; core's `ChartDataLabels` has the same fields. Placement needs plot geometry, which core's painters already hold. | Largest visible gap: 27 of 82 parts. The text composition is a pure function and ports first. |
| 2   | Axes: gridline presence and formatting, line and tick styling, deleted axes, axis titles, log scale, display units, crossing                                                                                     | 14 / 1,805                                      | Yes                                            | 12 of 14        | Medium: the neutral axis model (`ChartAxis`) carries these fields; core's painter ignores most.                                                                                | Core draws gridlines even when the part has none: a correctness bug, cheap to fix.            |
| 3   | Markers and series line styling (symbol, size, fill, line, smooth)                                                                                                                                               | 3 / 371                                         | Yes                                            | 2 of 3          | Easy                                                                                                                                                                           | 6 parts.                                                                                      |
| 4   | Legend: swatches, entry deletion, layout, text style                                                                                                                                                             | 5 / 531                                         | Yes                                            | 3 of 5          | Easy to medium; core has its own legend layout already                                                                                                                         | Legend font bug visible above.                                                                |
| 5   | Pie and doughnut: label placement and collision, best fit, explosion, first-slice angle, of-pie                                                                                                                  | 5 / 955                                         | Yes                                            | 4 of 5          | Easy: `pie-label-collision.ts`, `pie-best-fit.ts` already live in `src/core/chart/` and pptx re-exports them; core's pie painter does not call them                            | Depends on #1.                                                                                |
| 5b  | Fills: picture, pattern, gradient defs, per-point fills                                                                                                                                                          | 7 / 827                                         | Yes                                            | 5 of 7          | Medium (needs the image part resolver, which is host-supplied)                                                                                                                 | Gradients exist in core already (`chart-svg-appearance.ts`).                                  |
| 6   | Manual layout (plot area, labels, title, legend)                                                                                                                                                                 | 5 / 533                                         | Yes                                            | 5 of 5          | Easy: `chart/manual-layout.ts` is already shared                                                                                                                               | 10 parts.                                                                                     |
| 7   | Combo and secondary axes, stock combos                                                                                                                                                                           | 4 + 2 / 1,334                                   | Yes                                            | 6 of 6          | Medium to hard: core's summary holds extra groups but its axis scaling is single-pass                                                                                          | 1 combo part.                                                                                 |
| 8   | Data table                                                                                                                                                                                                       | 2 / 216                                         | Yes                                            | 2 of 2          | Easy                                                                                                                                                                           | No fixture; common in Excel charts.                                                           |
| 9   | Date axis (base unit, major and minor units, formats)                                                                                                                                                            | 2 / 435                                         | Yes                                            | 1 of 2          | Easy to medium                                                                                                                                                                 | No fixture; common in Excel.                                                                  |
| 10  | ChartEx: funnel, box and whisker, region map, pareto                                                                                                                                                             | 2 + 2 + 6 + 5 / ~1,800 (region map 1,084 of it) | Yes                                            | partly          | Medium; region map needs map data (large, keep out of core)                                                                                                                    | 3 of 7 ChartEx fixtures draw nothing in core.                                                 |
| 11  | Surface charts, flat and isometric                                                                                                                                                                               | 6 / 814                                         | Yes                                            | 5 of 6          | Medium                                                                                                                                                                         | 3 parts, now a labelled frame.                                                                |
| 12  | Trendline labels, error bars on category charts                                                                                                                                                                  | 6 / 680 (fitting already in core)               | Yes                                            | 5 of 6          | Easy                                                                                                                                                                           | Painting partly in core.                                                                      |
| 13  | Sparklines                                                                                                                                                                                                       | 1 / 191                                         | Yes                                            | 0               | Trivial                                                                                                                                                                        | No chart-part fixture (belongs to xlsx `x14:sparkline`).                                      |
| 14  | User shapes over a chart (`c:userShapes`)                                                                                                                                                                        | 6 / 964                                         | Yes                                            | 5 of 6          | Hard: needs the drawing-shape painter                                                                                                                                          | 1 part.                                                                                       |
| 15  | 3-D charts (three.js scenes, meshes, shading, interaction)                                                                                                                                                       | 40 / 5,816 (17 import three)                    | No                                             | 22 of 40        | Do not move                                                                                                                                                                    | Stays in `src/ui`. A flat 2-D approximation of 3-D in core is possible later.                 |
| 16  | Interaction, quick actions, editor, format panes                                                                                                                                                                 | 15 / 2,510 + 26 / 5,035                         | No (DOM, events)                               | 14              | Do not move                                                                                                                                                                    | UI only.                                                                                      |

## The architectural finding

The pptx renderer is not a set of drawing functions over a neutral model: about 80 percent of its
modules take `PptxChartData` (the pptx projection of the part, with colours resolved through the
theme) and produce a typed primitive tree (`SvgText`, `SvgPolygon`...) that three bindings map to
their templates. Core's renderer takes `ChartSummary`, builds `ChartViewModel` and prints SVG
strings. Moving "the renderer" therefore cannot be a file move: each feature is a port of its
algorithm onto core's view model (the pure pieces move verbatim, the pieces that read pptx fields are
rewritten against `ChartSpace`), with pptx delegating to the core function where the inputs are
neutral. That is how the slices below are done, one feature at a time, and why the `ooxml-ui` chart
view-model parity hashes stay untouched: pptx keeps its own painter and only reuses core's pure
helpers.

## What could not be verified

- No pptx-vs-core image comparison: the pptx renderer needs its element tree and a browser; only
  core's output was rendered and looked at.
- Detection is text-based, so a feature written with unusual prefixes would be missed; features with
  no fixture (data table, date axis, gradient, pattern, axis titles, error bars, of-pie) are
  unmeasured, not absent from the renderer.
- Ablation says "output changed", not "output is right".
- Line counts include comments (the pptx modules are heavily documented).

## Slices

Status of each port is recorded in `docs/agnostic-core-plan.md` (step 5, "Renderer unification")
and `PROVENANCE.md`.

1. **Data labels (landed 2026-10-17).** After the slice the harness measures core drawing labels on 27 of 27 parts that use them (0 of 27 before). Content, number format, `c:dLblPos` placement, per-point deletion and custom text, pie best fit. Still open inside gap 1: callout shapes and leader lines, manual label layout, "value from cells", rich-text runs, per-point fonts, labels on stock, waterfall, histogram and hierarchy charts.

### Axes and markers follow-up (2026-10-17)

The baseline measurements above precede these slices. Native axis metadata now reaches the
shared renderer through the summary, including the Excel/Visio workbook path. Axis IDs select
the primary and secondary scales; explicit bounds, log scales, reversed orientation, major/minor
units, number formats, deleted-axis labels/lines and gridline presence/paint are honored. The
range and tick implementations are extracted from PowerPoint, which imports the same core code.

PowerPoint's marker geometry also lives in strict core. The shared line, scatter and radar
painters honor native group defaults and series/point symbol, size, visibility, fill and outline.
The complete parse-to-SVG path is covered for all three families; the Excel native palette
round-trip checks also cover live and saved marker recoloring.

These are partial ports of rows 2 and 3: axis titles, display-unit captions, complete tick-mark
placement, date axes and line styling/smoothing remain. The baseline fixture counts above have
not been remeasured and are not evidence that those remaining features work.
