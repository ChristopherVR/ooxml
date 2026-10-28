# Roadmap

Where the project stands and what comes next. The repository layout and the way of working are described in the [repository README](../.github/README.md) and in [AGENTS.md](../AGENTS.md).

The active implementation order and acceptance criteria for shared chart, SmartArt and DrawingML
parity are tracked in [Shared Office parity delivery](office-shared-parity-delivery.md). It separates
extraction from consumer integration, rendering fidelity and native Office validation.

Implementation and validation are tracked in [the further sharing delivery record](pptx-reuse-delivery-2026-10-28.md).

The [2026-10-28 further sharing review](pptx-reuse-review-2026-10-28.md) identifies the next
cross-product extraction candidates, their current source locations, consumer gaps and acceptance
slices. It includes a fresh chart-renderer measurement and distinguishes shared-product gains
from PowerPoint-only binding consolidation.

## Done

- **Shared areas:** `xml`, `opc`, `units`, `color`, `geometry`, `drawingml`, `diagram` (SmartArt), `chart` (model, parser, writer and SVG painter), `digest`, `crypto` (package encryption), `math`, `text`, `i18n`, `ribbon`, `collab` (Yjs real-time collaboration), `teams` and `automation`. `AGENTS.md` has the authoritative list.
- **Formats as areas:** `docx`, `xlsx`, `pptx` and `visio`, each with its own parser and writer (Visio edits are bounded, and legacy `.vsd` is preview only).
- **Products:** the PowerPoint, Word, Excel, Visio and OpenTeams viewers and the shared `ooxml-ui` elements live in this repository and are released from one pipeline.
- **Wave 6 (2026-10-23):** one binding kit in `ooxml-ui/bindings`, pptx headless controllers, and chart and SmartArt editing shared across products. Word and Excel insert and edit SmartArt; Word, Excel and Visio edit charts; pptx saves classic chart title, legend, axis and series edits and ChartEx title, legend, label and series edits through the neutral writers. Details are in [structure-plan.md](structure-plan.md) section 5 and [agnostic-core-plan.md](agnostic-core-plan.md).
- **Shared Office delivery (2026-10-27):** source-backed chart admission and saves, interpreted Word/Excel SmartArt previews and generated 2D caches, shared SmartArt style/colour galleries, DrawingML outlines/effects/format schemes, advanced shared chart controls, and Word/Excel ChartEx editing. PowerPoint's admitted classic and ChartEx save families and shared model projections have corpus, browser and native reopen evidence. The [delivery ledger](office-shared-parity-delivery.md) records tested cases and remaining boundaries.

## Next

- **Finish moving the PowerPoint viewer in.** pptx-viewer now lives in `viewers/pptx` with its history and release tags; its packages publish from this repository (once their npm trusted publishers point here), its docs are part of the Pages site, its demos and browser tests are at `demos/pptx/` and `e2e/pptx/`, and CI verifies it like the other viewers, including its browser suite in the `pptx-*` jobs. Still to do:
  - retire the sync workflow, the local-linking script and the fleet report, which only served separate repositories, once nothing outside consumes this one.
- **Finish the DrawingML consumer migration:** shared outlines, typed effect lists, theme format schemes with `phClr` substitution and pattern/picture fill codecs are implemented. Move remaining pptx text, colour and theme object-tree adapters onto those APIs with parity coverage; effect DAG interpretation and broader effect rendering remain open. The order is in [agnostic-core-plan.md](agnostic-core-plan.md) step 3.
- **One XML model:** the `pptx` area still uses its own XML object model and is compiled with relaxed TypeScript flags while it is migrated onto the shared `xml` area and tightened. New code is strict.
- **Write each binding's logic once:** the five PowerPoint bindings still port the editor to each framework (about 410k lines together). The presentation controller (timers, annotations, presenter window), the Fill inspector section and the pptx tab row on `office-ui-ribbon` are the next pieces of the headless controllers in `ooxml-ui/pptx`, with one store adapter per framework and native templates on top. See [structure-plan.md](structure-plan.md) section 5.
- **Chart and SmartArt editing, the remaining families:**
  - extend PowerPoint's remaining source adapters and unsupported chart families (C3 to C5), including wider 3D/XY combinations, geographic/Pareto creation and ChartEx custom-label authoring. The delivered neutral save families cover labels, paint, rich text, axes, analysis, structure and point formatting; current corpus and native acceptance are recorded in the delivery ledger;
  - extend Word/Excel ChartEx editing beyond the admitted local-workbook, family-option and common-formatting controls; retain explicit refusal for unsupported external sources and ambiguous hierarchy mappings (C5);
  - SmartArt's wider native style catalogue and full source-cache/3D writer (S4, S5). Word and Excel share three style intensities and five colour schemes, with source-preserving edits and native reopen checks; admitted generated 2D caches are now linked into their saved packages and identified explicitly as approximations;
  - extend shared chart controls where rendering and host admission permit: picture relationship allocation, broader secondary-axis combinations, multiple trendline/custom-error source editors and fuller native rendering comparisons (U2). Effects, 3D/text formatting, pattern fills, markers, trendlines, error bars and axis titles already use shared intents and controls.
    The steps are defined in [agnostic-core-plan.md](agnostic-core-plan.md) sections 3.5 to 3.7 and the order is in [structure-plan.md](structure-plan.md) section 5.3.
- **Less logic in the UI packages:** the Word editor and PowerPoint renderer code now live in `src/ui/src/docx` and `src/ui/src/pptx`; DOM-free pieces of them still move into `src/core` areas over time.

The structure, parity and documentation work planned across all products is in [structure-plan.md](structure-plan.md).
