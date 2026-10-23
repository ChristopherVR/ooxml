# Roadmap

Where the project stands and what comes next. The repository layout and the way of working are described in the [repository README](../.github/README.md) and in [AGENTS.md](../AGENTS.md).

## Done

- **Shared areas:** `xml`, `opc`, `units`, `color`, `geometry`, `drawingml`, `diagram` (SmartArt), `chart` (model, parser, writer and SVG painter), `digest`, `crypto` (package encryption), `math`, `text`, `i18n`, `ribbon`, `collab` (Yjs real-time collaboration), `teams` and `automation`. `AGENTS.md` has the authoritative list.
- **Formats as areas:** `docx`, `xlsx`, `pptx` and `visio`, each with its own parser and writer (Visio edits are bounded, and legacy `.vsd` is preview only).
- **Products:** the PowerPoint, Word, Excel, Visio and OpenTeams viewers and the shared `ooxml-ui` elements live in this repository and are released from one pipeline.
- **Wave 6 (2026-10-23):** one binding kit in `ooxml-ui/bindings`, pptx headless controllers, and chart and SmartArt editing shared across products. Word and Excel insert and edit SmartArt; Word, Excel and Visio edit charts; pptx saves classic chart title, legend, axis and series edits and ChartEx title, legend, label and series edits through the neutral writers. Details are in [structure-plan.md](structure-plan.md) section 5 and [agnostic-core-plan.md](agnostic-core-plan.md).

## Next

- **Finish moving the PowerPoint viewer in.** pptx-viewer now lives in `viewers/pptx` with its history and release tags; its packages publish from this repository (once their npm trusted publishers point here), its docs are part of the Pages site, its demos and browser tests are at `demos/pptx/` and `e2e/pptx/`, and CI verifies it like the other viewers, including its browser suite in the `pptx-*` jobs. Still to do:
  - retire the sync workflow, the local-linking script and the fleet report, which only served separate repositories, once nothing outside consumes this one.
- **Finish DrawingML as a shared area:** the line and effect models (`a:ln` beyond width, fill, dash and cap; `a:effectLst`), the theme format scheme with `phClr` substitution, and the pptx text and colour parsers moved onto the neutral models (the neutral text model exists and is pinned to them by parity tests). The order is in [agnostic-core-plan.md](agnostic-core-plan.md) step 3.
- **One XML model:** the `pptx` area still uses its own XML object model and is compiled with relaxed TypeScript flags while it is migrated onto the shared `xml` area and tightened. New code is strict.
- **Write each binding's logic once:** the five PowerPoint bindings still port the editor to each framework (about 410k lines together). The presentation controller (timers, annotations, presenter window), the Fill inspector section and the pptx tab row on `office-ui-ribbon` are the next pieces of the headless controllers in `ooxml-ui/pptx`, with one store adapter per framework and native templates on top. See [structure-plan.md](structure-plan.md) section 5.
- **Chart and SmartArt editing, the remaining families:**
  - pptx chart data labels and fills saved through the neutral writer, then `PptxChartData` as an adapter over the neutral model, then pptx ChartEx parsing on the neutral model (agnostic plan C3 to C5);
  - Word and Excel ChartEx editing (C5);
  - SmartArt styles and colours, and the cached drawing writer (S4, S5);
  - the shared Format pane's remaining controls (effects, 3-D, text, trendlines, error bars, markers, axis titles) and Visio's use of it (U2).
    The steps are defined in [agnostic-core-plan.md](agnostic-core-plan.md) sections 3.5 to 3.7 and the order is in [structure-plan.md](structure-plan.md) section 5.3.
- **Less logic in the UI packages:** the Word editor and PowerPoint renderer code now live in `src/ui/src/docx` and `src/ui/src/pptx`; DOM-free pieces of them still move into `src/core` areas over time.

The structure, parity and documentation work planned across all products is in [structure-plan.md](structure-plan.md).
