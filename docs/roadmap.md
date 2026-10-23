# Roadmap

Where the project stands and what comes next. The repository layout and the way of working are described in the [repository README](../.github/README.md) and in [AGENTS.md](../AGENTS.md).

## Done

- **Shared areas:** `xml`, `opc`, `units`, `color`, `geometry`, `drawingml`, `diagram` (SmartArt), `chart` (model, parser, writer and SVG painter), `digest`, `crypto` (package encryption), `math`, `text`, `i18n`, `ribbon`, `collab` (Yjs real-time collaboration), `teams` and `automation`. `AGENTS.md` has the authoritative list.
- **Formats as areas:** `docx`, `xlsx`, `pptx` and `visio`, each with its own parser and writer (Visio edits are bounded, and legacy `.vsd` is preview only).
- **Products:** the PowerPoint, Word, Excel, Visio and OpenTeams viewers and the shared `ooxml-ui` elements live in this repository and are released from one pipeline.

## Next

- **Finish moving the PowerPoint viewer in.** pptx-viewer now lives in `viewers/pptx` with its history and release tags; its packages publish from this repository (once their npm trusted publishers point here), its docs are part of the Pages site, its demos and browser tests are at `demos/pptx/` and `e2e/pptx/`, and CI verifies it like the other viewers, including its browser suite in the `pptx-*` jobs. Still to do:
  - retire the sync workflow, the local-linking script and the fleet report, which only served separate repositories, once nothing outside consumes this one.
- **Finish DrawingML and charts as shared areas:** `drawingml` and `chart` exist and docx, xlsx and visio use them. What remains is the line and effect models, the theme format scheme, a paragraph and run model, and moving the pptx chart styles, ChartEx and writers onto the neutral chart model. The order is in [agnostic-core-plan.md](agnostic-core-plan.md).
- **One XML model:** the `pptx` area still uses its own XML object model and is compiled with relaxed TypeScript flags while it is migrated onto the shared `xml` area and tightened. New code is strict.
- **Write each binding's logic once:** the five PowerPoint bindings still port the editor to each framework (about 410k lines together). Logic moves into headless controllers in `ooxml-ui/pptx`, with one store adapter per framework and native templates on top; the Word, Excel, Visio and Teams wrappers share one binding kit in `ooxml-ui`. See [structure-plan.md](structure-plan.md) section 5.
- **Chart and SmartArt editing for every product:** editing operations, part writers and the editing UI (format pane, data sheet, insert pickers, design galleries, SmartArt text pane) move out of pptx into `chart`, `diagram`, `drawingml` and shared `ooxml-ui` elements, so Word, Excel and Visio edit charts and SmartArt the way PowerPoint does. The order is in [structure-plan.md](structure-plan.md) section 5.3.
- **Less logic in the UI packages:** the Word editor and PowerPoint renderer code now live in `src/ui/src/docx` and `src/ui/src/pptx`; DOM-free pieces of them still move into `src/core` areas over time.

The structure, parity and documentation work planned across all products is in [structure-plan.md](structure-plan.md).
