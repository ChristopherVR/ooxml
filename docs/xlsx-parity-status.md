# XLSX parity status

Updated 8 October 2026: mixed-title support at `4f3f87ea6`, followed by host
font measurements for chart titles/legends and natural title baselines.

**Full 1:1 Microsoft Excel parity has not been achieved.** The product is a
working spreadsheet editor with substantial supported behavior and an expanding
native comparison corpus. Compatibility is verified for specific cases, not
for every Excel feature, workbook or UI interaction. We have not established
lossless saving or whole-workbook pixel equivalence.

This page is the current summary. The [implementation review](xlsx-parity-review.md)
records each increment and its evidence; older sections describe the state at
their respective revisions. The [feature page](../viewers/xlsx/docs/features.md)
lists the broader product surface. [PROVENANCE](../PROVENANCE.md) records shared
logic reuse and extraction.

## Implemented and verified work

| Area                          | Current implementation                                                                                                                                                                                                                                                                                                                                                                                                                          | Important limits                                                                                                                                                                                   |
| ----------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Editing and clipboard         | Finite copies retain trailing blanks; compatible selections tile copied blocks. Directional fills copy content/formatting with translated formulas. Paste Special supports content modes, transpose, Skip Blanks, arithmetic, All Except Borders and column widths. Notes, validation, hyperlinks and supported conditional formats participate in the clipboard path. History, save/reload and structural reference behavior have regressions. | This is not complete Excel clipboard metadata or every cut/paste scenario. Multi-area paste and large-selection semantics have restrictions.                                                       |
| Conditional formatting        | Existing rules plus improvements to x14 data-bar preservation, copied rule identities, colors/borders, direction, axes, lengths, thresholds and context. Native comparisons exercise supported cases.                                                                                                                                                                                                                                           | Advanced rule authoring controls and exact data-bar raster appearance remain incomplete.                                                                                                           |
| Calculation                   | Existing formula engine, dependency recalculation and recorded Excel comparisons; follow-up work includes AMORLINC and DB/DDB compatibility cases.                                                                                                                                                                                                                                                                                              | Missing functions, coercion/array/date edge cases and external workbook resolution remain. A function catalogue is not proof of complete compatibility.                                            |
| Chart selection and structure | Pictures/charts can be selected, moved, resized and deleted. Charts support insertion, family/grouping changes, title and legend changes, history and save/reload.                                                                                                                                                                                                                                                                              | The supported SVG families are column, bar, line, area, pie, doughnut, scatter and radar. Other families and many native chart options remain unsupported.                                         |
| Chart formatting UI           | Shared Change Colors gallery, individual series fills, bar gap/overlap controls and docked fill panes. No fill, solid, transparency and supported linear/rectangular/circular/shape gradients have authoring paths, stop controls, previews and native presets. Chart area, plot area, title and legend reuse the same fill controls.                                                                                                           | Borders/effects, full axes/data-label editing, style/layout galleries and exact native ribbon/pane behavior remain incomplete. Gradient support varies by chart element/type.                      |
| Chart rendering and export    | Imported styles, supported gradients/shadows, axis visibility, built-in text defaults, chart-wide font inheritance and mixed title runs/lines/paragraphs. Supported fills and text formatting survive tested chart edits and regeneration.                                                                                                                                                                                                      | Text measurement, wrapping, spacing, automatic/manual layout and full chart raster fidelity remain approximate. Advanced text styling and mixed-format font authoring UI remain open.              |
| SmartArt                      | Imported cached drawings render through `office-ui-smartart`; shared theme colors, transforms and font choices improve native appearance. A native Basic Block List reference verifies three labels and measured text properties.                                                                                                                                                                                                               | Display-only. No insertion, text editing, layout reflow or complete layout/color/style galleries. Without a cached drawing, the viewer lists text in a placeholder. No whole-diagram pixel parity. |
| Shared architecture           | One XLSX core and one editor component, with six thin framework bindings. Shared XML/OPC, DrawingML/diagram, chart, geometry, color, units and Office UI controls are reused.                                                                                                                                                                                                                                                                   | Shared infrastructure does not by itself implement spreadsheet-specific behavior or establish parity.                                                                                              |

## Reuse from PowerPoint and other packages

- The chart picker uses `office-ui-gallery`, also used by PowerPoint's ribbon
  gallery. It shares visual tiles, selection, keyboard behavior and theme contracts.
- SmartArt uses the shared `office-ui-smartart` component and format-neutral
  diagram model/color/font resolution. XLSX supplies its workbook theme.
- Chart models, colors, gradients, fills, XML preservation and appearance logic
  live in shared core areas. Mixed titles reuse the DrawingML text reader and
  existing chart/theme resolver, rather than a viewer-specific text parser.
- Series and background formatting panes share fill bindings and controls;
  edits use the existing XLSX transaction, history and save pipelines.
- Modern OOXML logic stays in `ooxml-core`; legacy XLS/CFB codecs stay in `ole2`.
  Framework adapters do not contain duplicate spreadsheet engines.

## What the latest validation proves

At `4f3f87ea6`, the latest diagram/chart/XLSX sweep passed **8,534 ordinary
tests**, with **17 existing expected raster failures**. Those failures remain
explicit; this is not a clean pixel-parity result. **92 focused chart browser
checks** passed across vanilla, React, Vue, Angular, Svelte and Solid. Core/UI
typechecks, builds and package import/registration checks passed.

The subsequent host-measurement change reuses the grid's cached canvas measurer
for chart titles and legends, removing uniform-character width estimates from
their browser layout. The follow-up sweep passed **8,539 ordinary tests** plus
the same **17 expected failures**, and all **92 focused chart browser checks**
passed again. Browser checks compare painted rich-run advances and centering
before/after edits. Native placement, vertical metrics,
wrapping and axis-label measurement still require work; headless callers without
a supplied measurer retain the previous width estimates.

The next font-box increment passes **8,543 ordinary tests** with the same
**17 expected failures**, all **92 focused chart browser checks**, core/UI
typechecks, builds and package checks. The grid reuses its canvas context to supply font
ascent/descent for title heights and mixed-line baselines. Three independently
reopened Excel title-box references are retained in
`src/core/chart/excel-chart-title-geometry.json`. On the measured Windows fonts,
title heights are within 2 CSS pixels of these native references and a native
PNG comparison confirms improved second-line placement. This does not establish
general text/raster parity; native paragraph spacing, wrapping, padding and
overall chart geometry remain open.

Native evidence uses owned hidden Excel instances, principally Excel 16.0
build 20430. The current chart text corpus includes 96 built-in style
references, 24 chart-wide inheritance/direct-format references and three
mixed-title references. These are independently reopened measurements, with
chart XML retained in test-only fixtures.

Playwright MCP reviewed actual imports, controls and export flows. Recent UI
exports were reopened and resaved through Excel COM, then reopened again. The
mixed-title export retained the font size, name, bold, italic, underline and
color of all 23 measured title characters after a chart type change.

These results establish the tested behavior on that Excel build. They do not
prove every chart type, every workbook, every browser, Excel for Mac/web,
all Microsoft 365 versions or complete UI equivalence. Test counts are not
a parity percentage, and no defensible overall completion percentage exists yet.

## Major work still required

1. Native chart text measurement, wrapping, spacing, title/legend placement and
   overall chart geometry, followed by wider screenshot/raster comparisons.
2. Complete chart formatting/authoring UI, advanced families, axes, labels,
   effects and style/layout behavior, with native edit/export comparisons.
3. SmartArt insertion, text editing, reflow and shared layout/color/style UI.
4. Broader workbook save fidelity across structural edits, unsupported parts,
   drawings, tables, external links and metadata. No blanket lossless claim.
5. Grid/font/number-format/layout edge cases and native paginated printing,
   Page Layout and page-break preview.
6. Pivot refresh/rendering, slicers, timelines, sparklines, data connections
   and spreadsheet collaboration integration.
7. Wider formula and clipboard differential coverage. VBA execution, add-ins,
   Power Query and unsupported file formats also prevent whole-product parity.

## Completion criteria

Parity remains an active objective. Each feature needs evidence for import,
rendering, editing, undo/redo, save/reopen in Excel and the actual UI workflow,
using representative native fixtures and explicit unsupported cases. Visual
parity needs measured geometry and raster comparisons, not just matching XML
or SVG attributes. Full parity cannot be declared while any required feature
is missing, unverified or covered only by a narrower comparison.
