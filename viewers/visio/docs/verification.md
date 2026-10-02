# Verification record

Status: local development evidence, 2026-10-02. Full Microsoft Visio parity is not established.

Current focused checkpoint: 423 core Visio tests, 91 viewer tests, 43 framework/SSR tests and 18 documentation tests pass. Both core TypeScript projects, viewer/binding typechecks, production builds and worker/packed-consumer checks pass. A fresh full core suite and full-package build/import run is in progress; its outcome is recorded separately when complete.

## What has been exercised

- Strict TypeScript for the viewer and core; the existing relaxed PowerPoint project also remains type-correct.
- Core ZIP/XML/security, inheritance, geometry, text, layer and raster tests. The initial broad `ooxml` run passed 16,087 tests with 231 skipped; later isolated Visio additions require their own final focused rerun and are not represented by that older broad count.
- Viewer controller, resource lifecycle, source races, page-scoped selection, renderer safety and text-work budgets in a DOM test environment.
- Actual React, Vue, Angular, Svelte, Solid and vanilla lifecycle integration. The binding package has 38 DOM-environment tests and five separate no-DOM tests, including selected React/Vue hydration checks.
- Static documentation contracts: local/base-relative links, landmarks, labels, color contrast, responsive/reduced-motion rules and control-state tests.
- Production multipage build and parser worker bundle. The actual worker bundle parses a synthetic VSDX and returns structured errors in an isolated Node worker. This is not a browser/CSP test.
- The actual packed root artifact installs into a fresh consumer with the unreleased local core. ESM imports, declarations, a fresh Vite build and the packaged worker's valid/error paths pass. Framework adapters remain private source integrations, not standalone published artifacts.
- The core integration patch includes tracked changes and all new files. It applies to a clean checkout at the pinned revision. The setup script refuses mismatched revisions or unexpected source changes, uses a pinned npm lock and builds the Visio subpath in an isolated sibling setup.

## Real upstream corpus

The read-only research corpus contains 14 Microsoft Visio-authored drawings from LibreOffice/libvisio and five additional real drawings from Apache POI. The parser currently imports all 19: 22 pages, 706 normalized shapes and 627 geometry paths. Ten malformed security inputs are rejected. These observations apply to the inspected upstream revisions and do not establish universal format coverage.

Fixture bytes are kept outside this repository. Separate provenance records retain upstream URLs/revisions, SHA-256 values, license files and expected numeric/text assertions. No external document uploads were used.

### Visual diagnostics

The 14 libvisio drawings include embedded EMF previews. Those previews were converted locally through `emf-converter` and `@napi-rs/canvas`. The current viewer's SVG output was independently rasterized with native canvas font measurement. Uncropped page images and normalized content crops are retained separately, with generator/decoder versions and hashes.

This process detects gross omissions and transformation/color problems. It cannot certify browser rendering, exact fonts, full-page placement or native Visio fidelity. Several previews are stale or empty:

- `qs-box`: preview fill disagrees with the current drawing/upstream assertion
- `tdf136564-WhiteTextBackground`: embedded preview is blank
- `testfile6`: preview contains a connector/arrow/label absent from the current page/master XML
- `tdf154379-QuickStyleFillMatrix`: preview contains a larger process; the current file has a single document shape
- `github260`: embedded preview decodes empty/transparent, so it is not a visual oracle

These cases are excluded as authoritative fidelity oracles. A thumbnail mismatch is investigated against package XML before any code change. Solid theme-color assertions and actual gradient rendering are recorded separately. The post-fix SVGs contain true linear gradients. Saved corner rounding is implemented for closed axis-aligned rectangles and confirmed on five shapes in the 60973 source drawing. Other rounding cases remain explicitly unsupported.

## Not verified

- Real Chromium interaction, mobile layout or screenshots: the environment denied Chromium IPC sockets, including the approved retry, and the cloud browser blocked localhost. The Playwright suite is retained for a supported environment; its scenarios have not passed here.
- Microsoft Visio desktop open/render/edit/save/reopen testing
- A complete reference corpus with fixed Visio version, fonts, page settings and expected outputs
- Full accessibility testing with assistive technology
- All historical framework/Node/browser combinations
- Native VSDX writing, editing/history, formula recalculation, complete routing, legacy binary formats or full product parity
- GitHub CI, remote publication or Pages deployment: no remote repository has been created from this environment yet

## Re-run commands

Run `npm run check` after `npm run setup:core` and both root/binding installations. Run `npm run test:browser` only in a supported browser environment. `scripts/render-corpus.mjs INPUT_DIRECTORY OUTPUT_DIRECTORY` produces secondary-renderer diagnostic artifacts from an explicitly supplied local corpus.

Use the capability ledger for remaining functionality. Passing generated-fixture tests, showing a drawing, or absence of warnings is not a parity guarantee.
