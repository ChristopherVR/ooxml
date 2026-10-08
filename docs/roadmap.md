# Roadmap

Where the project stands and what comes next. The repository layout and the way of working are described in the [repository README](../.github/README.md) and in [AGENTS.md](../AGENTS.md).

## Done

- **Shared areas:** `xml`, `opc`, `units`, `color`, `geometry`, `diagram` (SmartArt), `digest`, `crypto` (package encryption), `math` and `collab` (Yjs real-time collaboration).
- **Formats as areas:** `docx`, `xlsx`, `pptx` and `visio`, each with its own parser and writer (Visio edits are bounded, and legacy `.vsd` is preview only).
- **Products:** the PowerPoint, Word, Excel, Visio and OpenTeams viewers and the shared `ooxml-ui` elements live in this repository and are released from one pipeline.

## Next

- **Finish moving the PowerPoint viewer in.** pptx-viewer now lives in `viewers/pptx` with its history and release tags; its packages publish from this repository (once their npm trusted publishers point here), its docs are part of the Pages site, its demos and browser tests are at `demos/pptx/` and `e2e/pptx/`, and CI verifies it like the other viewers, including its browser suite in the `pptx-*` jobs. Still to do:
  - retire the sync workflow, the local-linking script and the fleet report, which only served separate repositories, once nothing outside consumes this one.
- **DrawingML and charts as shared areas:** move the DrawingML (fills, lines, effects, text, theme) and chart code out of `pptx` so every format uses one implementation. The order is in [agnostic-core-plan.md](agnostic-core-plan.md).
- **One XML model:** the `pptx` area still uses its own XML object model and is compiled with relaxed TypeScript flags while it is migrated onto the shared `xml` area and tightened. New code is strict.
- **Less logic in the viewers:** the Word editor's view code and the PowerPoint viewer's `shared` render logic still live in the viewers and move into areas over time.

The structure, parity and documentation work planned across all products is in [structure-plan.md](structure-plan.md).
