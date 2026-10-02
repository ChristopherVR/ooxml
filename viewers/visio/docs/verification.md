# Verification record

Status: local development evidence, 2026-10-02. Full Microsoft Visio parity is not established.

## Shared editing checkpoint

The current local checkpoint passes 1,096 core Visio tests (five optional corpus skips), 357 viewer tests, 74 framework DOM tests, five SSR tests, and 18 documentation tests. Both core TypeScript projects, viewer/binding typechecks, builds, actual production parser/edit workers, and the packed external consumer pass. Independent controller/history and worker reviews found no remaining blockers after regression fixes. The edit worker exercises edit/reparse, unchanged-byte no-op, unsupported-input rejection, and retained embedded-EMF behavior.

Shared editing is experimental plain local text only. It includes bounded source/history ownership, undo/redo, cancellation, and explicit VSDX copy download. Core rejects unsupported rich text, master-linked targets, signatures/macros, ambiguous targets, and serialization-sensitive input. Untouched package part payloads are preserved; edited XML/ZIP representations and recalculated native appearance are not claimed byte-identical. Native Microsoft Visio reopen remains unverified.

The previous published checkpoint passed [core CI](https://github.com/ChristopherVR/ooxml/actions/runs/37071473295) at core `7364222` and [viewer CI](https://github.com/ChristopherVR/visio-viewer/actions/runs/37071926285) at viewer `3e54de0`, including ten Chromium tests. Those tests cover bounded EMF import/rejection and pixel agreement between live vector rendering, SVG export, and detached print artifacts. They do not certify the newly added editing UI browser scenarios. The new desktop/mobile edit/download scenarios require remote CI because local Chromium fails before assertions with a socket permission error.

The current viewer uses released `emf-converter` 4.8.9 for a narrow admitted primitive subset in a disposable parser/edit worker. Saved custom gradients support horizontal angles only. Full Microsoft Visio parity is not established.

## Historical evidence

The sections below describe earlier checkpoints and investigation limits; they are not additional claims about the current release. Earlier disabled-conversion statements refer to the inspection-only implementation before the bounded worker activation described above.

The earlier full core run at the 423-test Visio checkpoint recorded 16,417 passes, 231 skips and one unchanged PowerPoint AES-256 timeout; its full crypto file passed separately. A later whole-core local checkpoint passed 16,989 tests with 235 skips. The current published core also passed its complete remote CI gate. A redundant later local aggregate was stopped after remote success and is not recorded as a completed local run.

## What has been exercised

- Strict TypeScript for the viewer and core; the existing relaxed PowerPoint project also remains type-correct.
- Core ZIP/XML/security, inheritance, geometry, text, layer and raster tests. The earlier broad core result and its unrelated timeout are reported above. Focused counts identify the frozen local checkpoint; later feature additions must be rerun before their own checkpoint.
- Viewer controller, resource lifecycle, source races, page-scoped selection, renderer safety and text-work budgets in a DOM test environment. Bounded document text search adds 31 index/controller/UI regressions and six native-adapter cases, including reentrant navigation and unchanged-selection rendering work.
- Actual React, Vue, Angular, Svelte, Solid and vanilla lifecycle integration. The binding package has 67 DOM-environment tests and five separate no-DOM tests, including selected React/Vue hydration checks.
- Static documentation contracts: local/base-relative links, landmarks, labels, color contrast, responsive/reduced-motion rules and control-state tests.
- Production multipage build and parser worker bundle. The actual worker bundle parses a synthetic VSDX and returns structured errors in an isolated Node worker. This is not a browser/CSP test.
- The actual packed root artifact installs into a fresh consumer with the unreleased local core. ESM imports, declarations, a fresh Vite build and the packaged worker's valid/error paths pass. Framework adapters remain private source integrations, not standalone published artifacts.
- Bounded static SVG export passes shared API, lifecycle, XML-safety and amplification tests across all six handles. Original PNG/JPEG/GIF raster scenes pass dimension, crop, flip, alpha and shared-resource pixel assertions through librsvg 2.60.0/Cairo 1.18.4, with native canvas 1.0.10 generating and inspecting raster pixels. This is secondary-renderer evidence; the installed native canvas SVG decoder omits symbols/embedded raster images and is not used as the SVG oracle.
- Immutable print snapshots pass bounded composition, source-mutation isolation, copied-resource, current-page handle and all-six-adapter tests. Independent adversarial review passed 138 focused checks, including depth boundaries, non-array collections, aggregate paragraphs/gradients, serializer reentry and copy-time resource insertion. The four-page 60973 drawing produces 275,842 SVG bytes from 423 composed shapes. Snapshots preserve saved display visibility; they do not invoke printing or implement printer settings.
- Shared page/background layer controls add 19 viewer regressions and six native-adapter API cases. They synchronize selection/search, preserve focus and leave saved SVG/print artifacts unchanged. Layer controls are bounded to 200 visible UI rows; immutable API overrides are capped at 25,000.
- Cached spline sequences and the shared positive-weight NURBS kernel use control-hull subdivision after bounded knot insertion. Analytic and adversarial tests include a degree-17 curve whose one-inch lobes escaped the previous fixed samples, plus strict predecessor-cache validation. This is stronger geometric evidence, not native Visio certification.
- Separate cached visibility reasons and a bounded core layer-visibility resolver pass 74 tests. The original display-hidden flags remain unchanged across the 19-file real corpus.
- Enhanced-metafile inspection is bounded per part and per document, with deduplicated pending work, an explicit asset queue and pre-inflation declared-size checks. It produces compatibility diagnostics only. The neutral-vector boundary independently validates SVG-tree and transported canonical inputs; live EMF conversion is disabled.
- The prospective path-only foreign-vector renderer passes generated clip/transform/viewport pixels through native canvas1.0.10's Skia SVG backend. A parallel librsvg2.60.0/Cairo1.18.4 probe fails nested clip-path intersections; this divergence is retained explicitly, including the optional `--require-librsvg` assertion. Browser and native Visio evidence is still absent.
- The core integration patch includes tracked changes and all new files. It applies to a clean checkout at the pinned revision. The setup script refuses mismatched revisions or unexpected source changes, uses a pinned npm lock and builds the Visio subpath in an isolated sibling setup.

## Real upstream corpus

The read-only research corpus contains 14 Microsoft Visio-authored drawings from LibreOffice/libvisio and five additional real drawings from Apache POI. The parser currently imports all 19: 22 pages, 706 normalized shapes and 627 geometry paths. Ten malformed security inputs are rejected. These observations apply to the inspected upstream revisions and do not establish universal format coverage.

Fixture bytes are kept outside this repository. Separate provenance records retain upstream URLs/revisions, SHA-256 values, license files and expected numeric/text assertions. No external document uploads were used. The optional [corpus runner](corpus.md) reproduces the hash-pinned parse, scene-validation, color and rounded-rectangle assertions from explicitly supplied local upstream checkouts. Its optional `--svg` pass successfully exports all 22 pages, verifies XML, preserved diagnostics and local-only resources, and excludes controls/event attributes.

### Visual diagnostics

The 14 libvisio drawings include embedded EMF previews. Those previews were converted locally through `emf-converter` and `@napi-rs/canvas`. The current viewer's SVG output was independently rasterized with native canvas font measurement. Uncropped page images and normalized content crops are retained separately, with generator/decoder versions and hashes.

An isolated local audit reproduced installed emf-converter 3.5.1 defects in odd-count driver strings, valid leaf-region replacement and MM_TEXT mapping extents. Even-count driver strings also ignore individual positions and the optional matrix. The original upstream Windows comparisons are not independently verified. The audit harness has a tiny fail-closed vector subset, which admits none of the five real embedded EMFs in the current corpus. Converter output remains secondary evidence; production EMF rendering stays disabled pending record-specific safety and fidelity checks. See [EMF adoption review](research/emf-adoption-review.md).

This process detects gross omissions and transformation/color problems. It cannot certify browser rendering, exact fonts, full-page placement or native Visio fidelity. Several previews are stale or empty:

- `qs-box`: preview fill disagrees with the current drawing/upstream assertion
- `tdf136564-WhiteTextBackground`: embedded preview is blank
- `testfile6`: preview contains a connector/arrow/label absent from the current page/master XML
- `tdf154379-QuickStyleFillMatrix`: preview contains a larger process; the current file has a single document shape
- `github260`: embedded preview decodes empty/transparent, so it is not a visual oracle

These cases are excluded as authoritative fidelity oracles. A thumbnail mismatch is investigated against package XML before any code change. Solid theme-color assertions and actual gradient rendering are recorded separately. The post-fix SVGs contain true linear gradients. Saved corner rounding is implemented for closed axis-aligned rectangles and confirmed on five shapes in the 60973 source drawing. Other rounding cases remain explicitly unsupported. Cached dash patterns 2-23 now use bounded core-normalized stroke-width ratios, with inference diagnostics. Six actual pattern 23 shapes in 60973 retain their saved thin stroke widths without an arbitrary floor. Transparent/NoLine geometry no longer renders arrowheads. Numerical spacing still requires native reference comparison.

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

## Initial converter package adapter checkpoint

The isolated bridge now targets emf-converter 4.8.7 with the reviewed local GDI fix
patch. The older 3.5.1 audit remains separate. Core `convertVisioMetafile` accepts
a trusted browser-package function and delegates all drawing conversion to that
package. It admits at most 512 records and 256 KiB with 2048-pixel dimensions and
only header/EOF, stock object selection, rectangle, ellipse, move and line records.
Every other record is rejected before package invocation. Input is privately
copied and reinspected; only sanitized neutral vectors cross the result boundary.

Seven isolated browser-distribution integration cases cover empty output, three
primitive drawings, stock paint selection and two rejected stream classes. These are generated structural
and adapter regressions, not native-render parity evidence. Production VSDX
conversion stays disabled. The adapter requires a disposable host worker with a
parent deadline; the test harness exercises that arrangement. It does not claim
a hard converter heap limit.

Browser regression execution is currently launch-blocked: the bundled Playwright
Chromium shell is missing; the installed Chromium alternative aborts before tests
with `socket() Operation not permitted`. All eight browser cases fail before their
assertions. This does not establish a product regression or a browser pass.

Initial checkpoint commands: `npm run check`, `npm run check:converter` (now
`npm run check:converter-source`),
`npm run test:converter-integration` and `npm run test:converter-setup`. The isolated
converter suite passes 3,743 tests with 117 skips (92 files pass, one file skips);
its TypeScript, Node/browser builds and browser-package consumer checks pass.
Ten bridge setup tests pass, including committed/freshly applied patch equivalence,
refusal of unrelated edits and semantic lock comparison. A fresh local baseline
clone plus the full patch independently reproduces the pinned source snapshot.
The integration runner verifies this snapshot and rebuilds the browser distribution
before its seven deadline-isolated cases, so stale output cannot count as evidence.

## Released converter adoption

The normal viewer setup and aggregate check now use published `emf-converter`
4.8.8 as an exact dev dependency. npm registry integrity is locked and the actual
browser artifact SHA-256 is verified. Its browser bundle is byte-identical to the
corrected local build above; all seven isolated integration cases pass with the
released package selected through its browser export. The source correction
bridge remains historical provenance, available only through explicit
`setup:converter-source` / `check:converter-source` commands. It is no longer a
normal build requirement. The older audit remains unchanged. Core publication
and the live-conversion gate are unaffected.

Release adoption validation also passes from a temporary harness with no sibling
converter source checkout. Only the installed package, locked release metadata,
core package and generated test inputs are required.
