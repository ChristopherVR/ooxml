# Verification record

Status: local development evidence, 2026-10-02. Full Microsoft Visio parity is not established.

Current focused checkpoint: 582 core Visio tests, 186 viewer tests, 60 framework/SSR tests and 18 documentation tests pass. Both core TypeScript projects, viewer/binding typechecks, production builds and worker/packed-consumer checks pass. The earlier full core run at the 423-test Visio checkpoint recorded 16,417 passes, 231 skips and one failure: the unchanged PowerPoint AES-256 test exceeded its existing 20-second timeout under aggregate load. Its entire crypto file passes all 72 tests when rerun alone. The full core build, all 16 non-CLI packed ESM/CJS exports and 47 script tests pass. The original aggregate run is not recorded as green.

## What has been exercised

- Strict TypeScript for the viewer and core; the existing relaxed PowerPoint project also remains type-correct.
- Core ZIP/XML/security, inheritance, geometry, text, layer and raster tests. The earlier broad core result and its unrelated timeout are reported above. Focused counts identify the frozen local checkpoint; later feature additions must be rerun before their own checkpoint.
- Viewer controller, resource lifecycle, source races, page-scoped selection, renderer safety and text-work budgets in a DOM test environment. Bounded document text search adds31 index/controller/UI regressions and six native-adapter cases, including reentrant navigation and unchanged-selection rendering work.
- Actual React, Vue, Angular, Svelte, Solid and vanilla lifecycle integration. The binding package has 55 DOM-environment tests and five separate no-DOM tests, including selected React/Vue hydration checks.
- Static documentation contracts: local/base-relative links, landmarks, labels, color contrast, responsive/reduced-motion rules and control-state tests.
- Production multipage build and parser worker bundle. The actual worker bundle parses a synthetic VSDX and returns structured errors in an isolated Node worker. This is not a browser/CSP test.
- The actual packed root artifact installs into a fresh consumer with the unreleased local core. ESM imports, declarations, a fresh Vite build and the packaged worker's valid/error paths pass. Framework adapters remain private source integrations, not standalone published artifacts.
- Bounded static SVG export passes shared API, lifecycle, XML-safety and amplification tests across all six handles. Original PNG/JPEG/GIF raster scenes pass dimension, crop, flip, alpha and shared-resource pixel assertions through librsvg 2.60.0/Cairo 1.18.4, with native canvas 1.0.10 generating and inspecting raster pixels. This is secondary-renderer evidence; the installed native canvas SVG decoder omits symbols/embedded raster images and is not used as the SVG oracle.
- The core integration patch includes tracked changes and all new files. It applies to a clean checkout at the pinned revision. The setup script refuses mismatched revisions or unexpected source changes, uses a pinned npm lock and builds the Visio subpath in an isolated sibling setup.

## Real upstream corpus

The read-only research corpus contains 14 Microsoft Visio-authored drawings from LibreOffice/libvisio and five additional real drawings from Apache POI. The parser currently imports all 19: 22 pages, 706 normalized shapes and 627 geometry paths. Ten malformed security inputs are rejected. These observations apply to the inspected upstream revisions and do not establish universal format coverage.

Fixture bytes are kept outside this repository. Separate provenance records retain upstream URLs/revisions, SHA-256 values, license files and expected numeric/text assertions. No external document uploads were used. The optional [corpus runner](corpus.md) reproduces the hash-pinned parse, scene-validation, color and rounded-rectangle assertions from explicitly supplied local upstream checkouts. Its optional `--svg` pass successfully exports all 22 pages, verifies XML, preserved diagnostics and local-only resources, and excludes controls/event attributes.

### Visual diagnostics

The 14 libvisio drawings include embedded EMF previews. Those previews were converted locally through `emf-converter` and `@napi-rs/canvas`. The current viewer's SVG output was independently rasterized with native canvas font measurement. Uncropped page images and normalized content crops are retained separately, with generator/decoder versions and hashes.

This process detects gross omissions and transformation/color problems. It cannot certify browser rendering, exact fonts, full-page placement or native Visio fidelity. Several previews are stale or empty:

- `qs-box`: preview fill disagrees with the current drawing/upstream assertion
- `tdf136564-WhiteTextBackground`: embedded preview is blank
- `testfile6`: preview contains a connector/arrow/label absent from the current page/master XML
- `tdf154379-QuickStyleFillMatrix`: preview contains a larger process; the current file has a single document shape
- `github260`: embedded preview decodes empty/transparent, so it is not a visual oracle

These cases are excluded as authoritative fidelity oracles. A thumbnail mismatch is investigated against package XML before any code change. Solid theme-color assertions and actual gradient rendering are recorded separately. The post-fix SVGs contain true linear gradients. Saved corner rounding is implemented for closed axis-aligned rectangles and confirmed on five shapes in the 60973 source drawing. Other rounding cases remain explicitly unsupported. Cached dash patterns2-23 now use bounded core-normalized stroke-width ratios, with inference diagnostics. Six actual pattern23 shapes in60973 retain their saved thin stroke widths without an arbitrary floor. Transparent/NoLine geometry no longer renders arrowheads. Numerical spacing still requires native reference comparison.

## Not verified

- Real Chromium interaction, mobile layout or screenshots: the environment denied Chromium IPC sockets, including the approved retry, and the cloud browser blocked localhost. The Playwright suite is retained for a supported environment; its scenarios have not passed here.
- Microsoft Visio desktop open/render/edit/save/reopen testing
- A complete reference corpus with fixed Visio version, fonts, page settings and expected outputs
- Full accessibility testing with assistive technology
- All historical framework/Node/browser combinations
- Native VSDX writing, editing/history, formula recalculation, complete routing, legacy binary formats or full product parity
- GitHub CI, remote publication or Pages deployment: no remote repository has been created from this environment yet

## Re-run commands

Run `npm run check` after `npm run setup:core` and both root/binding installations. Run `npm run test:browser` only in a supported browser environment. The current eight browser scenarios remain launch-blocked here. `scripts/render-corpus.mjs INPUT_DIRECTORY OUTPUT_DIRECTORY` produces secondary-renderer diagnostic artifacts from an explicitly supplied local corpus.

Use the capability ledger for remaining functionality. Passing generated-fixture tests, showing a drawing, or absence of warnings is not a parity guarantee.
