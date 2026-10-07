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

The private `pptx-viewer-shared` package is a compatibility facade. It owns no
renderer or document operations. Angular no longer vendors a source copy.

## Compatibility and validation

PowerPoint's migrated UI keeps its existing compiler flags in a separate
`tsconfig.pptx.json`; all other shared UI retains the strict configuration.
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

Validation: 10,475 migrated UI tests, 674 extracted core unit tests, all five
framework builds/typechecks, clean package imports, and eight browser smoke
checks per framework. The optional live AI tests remain opt-in.
