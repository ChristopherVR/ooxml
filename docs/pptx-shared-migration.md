# PowerPoint shared implementation migration

PowerPoint follows the same ownership as Word: `src/core/pptx` owns document
operations and `src/ui/src/pptx` owns the shared renderer, web components,
localization, export views, browser loading lifecycle and assistant UI. The
viewer packages contain framework adapters and consume `ooxml-ui/pptx`.

## Public entries

- `ooxml-ui/pptx`: rendering and web-control APIs.
- `ooxml-ui/pptx/theme`, `/loader`, `/i18n`, `/ai`: focused entries.
- `ooxml-ui/pptx/<module>`: explicit renderer/export leaf entries needed by
  Angular; only built modules are public.
- `ooxml-core/pptx/editor/<module>`: DOM-free editing, loading and collaboration
  helpers, preserving existing function names through compatibility exports.
- `ooxml-core/pptx/automation/schemas`: the assistant/MCP tool contracts.
- `ooxml-core/pptx/editor/render/clone`: document and undo snapshot cloning.
- `ooxml-core/diagram`: cached drawing bounds shared across Office formats.

The private `pptx-viewer-shared` package is a compatibility facade. It owns no
renderer or document operations. Angular no longer vendors a source copy.

## Controllers and shared editing UI (Wave 6, 2026-10-23)

The five bindings no longer each hold their own copy of the editor's stateful logic. Undo and redo
history, recovery autosave and its recovery offer, the collaboration room lifecycle, the AI panel and
the canvas gestures are headless controllers in `src/ui/src/pptx/render/controller/` (and
`render/editor-gestures.ts`, `ai/panel-controller.ts`), each adapted by one store adapter per
framework. Keymap dispatch, the AI bridge, group press targets and slide show steps are pure
functions in the same shared layer. The binding conventions are in `viewers/pptx/AGENTS.md` (Rule 2).

Chart and SmartArt editing is shared too. The chart data sheet, chart Format pane, insert chart and
insert SmartArt dialogs, the SmartArt text pane and the chart Design galleries are `ooxml-ui`
elements (`office-ui-chart-data-sheet`, `office-ui-chart-format-pane`, `office-ui-chart-picker`,
`office-ui-smartart-picker`, `office-ui-smartart-text-pane`) that the pptx web components mount.
The per-binding chart grids, insert SmartArt dialogs and SmartArt text panes were removed. Chart edit
operations and the ChartEx model and writers live in `src/core/chart` (`edit`, `ex`, `workbook`), and
PowerPoint saves its classic and ChartEx title, legend, axis, label and series edits through them.

Angular's `pptx-angular-viewer/internals` no longer exports `InsertSmartArtDialogComponent`,
`SmartArtInsertEvent` or `SmartArtPreviewComponent`; their shared replacements are the elements above.

## Compatibility and validation

PowerPoint's migrated UI keeps its existing compiler flags in `src/ui/tsconfig.pptx.json`
(the core's relaxed pptx project is `src/core/tsconfig.pptx.json`); all other shared UI retains the
strict configuration.
Tightening these legacy flags is a separate correctness task, as with the
existing PowerPoint engine. Importing the main UI entry does not load PowerPoint.
The optional Three.js and AI SDK integrations remain optional peers.

Viewer maintenance scripts read the migrated UI source. Bun commands that
import these modules use `viewers/pptx/scripts/tsconfig.runtime.json` through
`--tsconfig-override`: compiler-only aliases target declarations and must not
be used for runtime resolution. The locale and customization package commands
apply this configuration automatically. Catalogue generators write into
`src/ui/src/pptx`, rather than recreating the viewer's old shared tree.

The original tests moved with their implementations. Tests spanning DOM and
model operations remain in UI; core unit tests stay DOM-free. Product tests run
with `bun run --cwd src/ui test:pptx`, against built package exports. Build core
first, then UI, then the framework bindings. The UI package smoke check imports
all public ESM entries, and core checks both ESM and CJS exports.

Extraction provenance is recorded in `PROVENANCE.md` and the complete module
inventory in `docs/pptx-migration-provenance.json`.

Validation at the time of the migration (2026-10-07; the counts have grown since): 10,475 migrated
UI tests, 674 extracted core unit tests, all five framework builds/typechecks, clean package
imports, and eight browser smoke checks per framework. The optional live AI tests remain opt-in.
