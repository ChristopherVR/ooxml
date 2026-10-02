# Verification record

Status: local development evidence, 2026-10-02. Full Microsoft Visio parity is not established.

## PPTX-aligned interface checkpoint

The app and Pages sources now share a PPTX-referenced design language: compact Office-style workspace chrome, orange accents, Home/View tabs, page navigation, inspector cards, and coordinated light/dark themes. The landing and guide use a warm paper/rust palette, large headline, framework examples, and the existing lazy local demo. Public-beta and placeholder-package limitations remain explicit.

Local aggregate checks pass: 1,109 core tests (six optional corpus skips), 387 viewer tests (one optional corpus skip), 79 bindings tests, and 24 documentation tests. Formatting, strict typechecks, production builds, actual parser/edit workers, and the packed external consumer pass. Independent source review found no remaining blocking issue in the reviewed chrome/theme/navigation paths.

Desktop/mobile light/dark screenshot and interaction scenarios are committed for remote Chromium CI. Their visual result is pending at this checkpoint: local Chromium cannot launch. Screenshot artifacts contain synthetic demo drawings and documentation only. This redesign does not establish Microsoft Visio rendering or editing parity.

## Local rendering checkpoint

The combined rendering slice passed the complete viewer `npm run check` before
the upstream-only release/site rebase. Fresh exact-pinned core `2c6e66a` checks
then passed the same gates with: 1,113 core Visio tests (two optional EMF skips),
375 viewer tests (including the optional arrow corpus), 74 framework DOM tests, five SSR tests
and 18 documentation tests. The hash-pinned arrow corpus has 18 arrow tests; its real-document case uses a
30-second integration timeout after exceeding the default five seconds under
full-suite parallel load. A fresh exact-pinned setup passed, and all viewer gates
were rerun after that test-only timeout adjustment. The fresh parallel binding
runner was externally killed (exit 137); its 74 DOM and five SSR tests passed
when retried with two workers. Both core TypeScript projects, full core build and every packed
ESM/CJS entry import passed, as did viewer/binding types, production workers and
the packed external viewer consumer. The real corpus accepted all 19 drawings,
rejected ten malformed packages and exported 22 safe SVG pages. The new browser
arrow pixel scenario remains pending remote CI. Native Visio equivalence is not
established.

## Published shared editing checkpoint

The published shared editing checkpoint passed 1,096 core Visio tests (five optional corpus skips), 357 viewer tests, 74 framework DOM tests, five SSR tests, and 18 documentation tests. Both core TypeScript projects, viewer/binding typechecks, builds, actual production parser/edit workers, and the packed external consumer pass. Independent controller/history and worker reviews found no remaining blockers after regression fixes. The edit worker exercises edit/reparse, unchanged-byte no-op, unsupported-input rejection, and retained embedded-EMF behavior.

Shared editing is experimental plain local text only. It includes bounded source/history ownership, undo/redo, cancellation, and explicit VSDX copy download. Core rejects unsupported rich text, master-linked targets, signatures/macros, ambiguous targets, and serialization-sensitive input. Untouched package part payloads are preserved; edited XML/ZIP representations and recalculated native appearance are not claimed byte-identical. Native Microsoft Visio reopen remains unverified.

The published editing baseline passed [core CI](https://github.com/ChristopherVR/ooxml/actions/runs/37071473295) at core `7364222` and [viewer CI](https://github.com/ChristopherVR/visio-viewer/actions/runs/37073518667) at viewer `c43639f`, including 12 Chromium tests. Desktop literal-text editing, undo/redo, VSDX download/reopen and mobile Escape/focus/touch/reset checks passed remotely. Earlier vector import/export/print pixel evidence passed in [the ten-test checkpoint](https://github.com/ChristopherVR/visio-viewer/actions/runs/37071926285). The new code-5 arrow browser pixel test remains pending remote CI; local Chromium still fails before assertions with a socket permission error.

The current viewer uses released `emf-converter` 4.8.9 for a narrow admitted primitive subset in a disposable parser/edit worker. Saved custom gradients support horizontal angles only. Full Microsoft Visio parity is not established.

## Historical evidence

The sections below describe earlier checkpoints and investigation limits; they are not additional claims about the current release. Earlier disabled-conversion statements refer to the inspection-only implementation before the bounded worker activation described above.

The earlier full core run at the 423-test Visio checkpoint recorded 16,417 passes, 231 skips and one unchanged PowerPoint AES-256 timeout; its full crypto file passed separately. A later whole-core local checkpoint passed 16,989 tests with 235 skips. The current published core also passed its complete remote CI gate. A redundant later local aggregate was stopped after remote success and is not recorded as a completed local run.

## Earlier exercised checkpoints

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

These cases are excluded as authoritative fidelity oracles. A thumbnail mismatch is investigated against package XML before any code change. Solid theme-color assertions and actual gradient rendering are recorded separately. The post-fix SVGs contain true linear gradients. Saved corner rounding is implemented for closed axis-aligned rectangles and confirmed on five shapes in the 60973 source drawing. The new core slice additionally rounds five nonclamped open orthogonal connectors (14 corners); two short-segment connectors remain unchanged and diagnosed. General polygon and mixed/curved rounding remain unsupported. Cached dash patterns 2-23 now use bounded core-normalized stroke-width ratios, with inference diagnostics. Six actual pattern 23 shapes in 60973 retain their saved thin stroke widths without an arbitrary floor. Transparent/NoLine geometry no longer renders arrowheads. Numerical spacing still requires native reference comparison.

## Current limitations

- Local Chromium remains launch-blocked by IPC socket permissions. The 12 existing scenarios passed remotely at the published editing baseline; the newly added arrow pixel scenario has not yet passed remote CI.
- Microsoft Visio desktop open/render/edit/save/reopen testing
- A complete reference corpus with fixed Visio version, fonts, page settings and expected outputs
- Full accessibility testing with assistive technology
- All historical framework/Node/browser combinations
- Native Microsoft Visio validation of VSDX editing/history, formula recalculation, complete routing, legacy binary formats or full product parity
- Remote CI for this new rendering slice is pending. Published baseline CI is linked above; no Pages deployment is claimed.

## Re-run commands

Run `npm run check` after `npm run setup:core` and both root/binding installations. Run `npm run test:browser` only in a supported browser environment. The browser scenarios remain launch-blocked locally; published baseline remote evidence is linked above. `scripts/render-corpus.mjs INPUT_DIRECTORY OUTPUT_DIRECTORY` produces secondary-renderer diagnostic artifacts from an explicitly supplied local corpus.

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

## Code-5 concave arrow slice

The viewer now renders the saved code-5 concave arrowhead. Hash-pinned `60973.vsdx`
page ID 0 shapes 3/7/8/10 restore eight formerly omitted marker ends. Independent
analytic checks cover inward-base topology, bounded coordinates and outline,
endpoint anchoring, SVG orientation, opacity, short lines, NoLine and unsupported
codes. This is a source-confirmed symbol restoration with approximate dimensions;
JSDOM checks do not establish browser pixels or native Visio fidelity. See
[the evidence and reproduction commands](research/concave-arrow-evidence.md).

## Open orthogonal connector slice

The rebuilt core rounds five source-confirmed `60973.vsdx` connectors: page ID 4
shapes 814/830/835/857 and page ID 7 shape 293, totaling 14 circular corners.
Page ID 4 shape 825 and page ID 7 shape 149 remain sharp with explicit diagnostics
because the saved radius does not fit. Repeated vertices, reversals, diagonals,
curves and malformed coordinates are declined. No guessed short-segment clamping
is included. Analytic, source and secondary-renderer checks do not establish
native Microsoft Visio pixel equivalence. The companion core
`docs/visio-connector-rounding.md` records primary semantics and fixture hashes.
