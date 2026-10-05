# Roadmap

Where the project stands and what comes next. The repository layout and the way of working are described in the [repository README](../.github/README.md) and in [AGENTS.md](../AGENTS.md).

## Done

- **Shared areas:** `xml`, `opc`, `units`, `color`, `geometry`, `diagram` (SmartArt), `digest`, `crypto` (package encryption), `math` and `collab` (Yjs real-time collaboration).
- **Formats as areas:** `docx`, `xlsx`, `pptx` and `visio`, each with its own parser and writer (Visio edits are bounded, and legacy `.vsd` is preview only).
- **Products:** the Word, Excel, Visio and OpenTeams viewers and the shared `ooxml-ui` elements live in this repository and are released from one pipeline.

## Next

- **Merge pptx-viewer into this repository.** The PowerPoint viewer is the one product still in its own repository. It will move in as `viewers/pptx` with its history and release tags, the way the other four viewers did. That also retires the scheduled sync workflow, the local-linking script and the fleet report, which exist only to keep separate repositories in step. It is not scheduled, and it needs some preparation first:
  - its CI (end-to-end tests sharded across five frameworks) has to be fitted into this repository's change-based CI planner, so a small change does not start every shard;
  - its release, changelog and Pages flows have to move into the root release planner and the combined site;
  - its package manager, lockfile, hooks and repository rules have to be reconciled with the root workspace.
- **DrawingML and charts as shared areas:** move the DrawingML (fills, lines, effects, text, theme) and chart code out of `pptx` so every format uses one implementation. The order is in [agnostic-core-plan.md](agnostic-core-plan.md).
- **One XML model:** the `pptx` area still uses its own XML object model and is compiled with relaxed TypeScript flags while it is migrated onto the shared `xml` area and tightened. New code is strict.
- **Less logic in the viewers:** the Word editor's view code and the PowerPoint viewer's `shared` render logic still live in the viewers and move into areas over time.
