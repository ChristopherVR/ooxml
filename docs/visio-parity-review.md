# Visio parity review, 2026-10-07

Reviewed baseline: `f6e9b730a` (ooxml-core 0.23.0, ooxml-ui 0.30.1).
The original working directory was detached at `7364222cc`, so this work uses
an isolated checkout of the current remote main. The current architecture places
VSDX semantics in `src/core/visio`, browser rendering and controls in
`src/ui/src/visio`, and six bindings, demos and browser tests in the viewer areas.
The capability ledger remains `viewers/visio/docs/parity.md`.

## Findings

The viewer is a substantial beta, with local VSDX loading, master/style inheritance,
cached geometry, rich-run text rendering, layers, search, shared ribbon controls,
page navigation, zoom, source-backed history and bounded editing/save operations.
These are implemented subsets. The visible ribbon includes disabled commands:
clipboard, rich-text formatting, style changes, connector tools and several shape
tools still require core commands. Six bindings share one controller and renderer.

The broad blockers are formula/master recalculation, live connector routing and
glue, accurate text measurement, complete line/fill/theme/foreign-object support,
general editing and native round-trip validation. Containers/swimlanes, stencils,
legacy formats, printing and collaboration also have substantial gaps. There is
no defensible overall parity percentage or full native visual-equivalence claim.
Historical verification counts in the ledger are not fresh results from this review.

## First implemented increment

Open orthogonal connector rounding previously rejected sections when the saved
radius did not fit. It also allowed full endpoint radii where native Visio reserves
half the segment. Ordered radius allocation now matches nine measured Visio 16.0
cases. The shared core correction reaches all six viewers through their existing
renderer, without adapter changes. Endpoint coordinates, transforms and package
payloads are untouched; extra emitted arcs retain the geometry-command budget.

The Apache POI 60973 regression now covers eighteen cached corners across seven
connectors. Shape 825 agrees with native exported radii. Shape 149 is recalculated
by Visio when opened and remains a live-routing gap. See
[connector evidence](visio-connector-rounding.md) for provenance and reproduction.

## Implementation order toward parity

| Order | Work                                                                                                             | Acceptance evidence                                                                                                   |
| ----- | ---------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| 1     | Expand the native comparison harness to fonts, paths, styles, transforms and themes                              | Versioned native VSDX and reference exports; per-feature geometry/pixel differences, with settings and fonts recorded |
| 2     | Finish cached viewing fidelity: arrowhead codes/sizing, gradients/patterns, text tabs/fields and foreign objects | Native-authored feature matrix plus genuine document regressions; explicit fallback diagnostics                       |
| 3     | Implement scoped ShapeSheet/master dependencies and connector glue/routing                                       | Resize/move native masters and connected shapes; compare dependent geometry after native reopen                       |
| 4     | Add general source-backed editing: style/rich text, selection transforms, pages, shapes and clipboard            | Atomic edits, undo/redo, untouched-part preservation and native save/reopen checks                                    |
| 5     | Add container/swimlane semantics, stencil/template workflows and export/print behavior                           | Representative native workflows and browser coverage across the six bindings                                          |

Treat native render parity, interaction parity and save parity as separate gates.
A matching synthetic render cannot establish package preservation, and a successful
library round-trip cannot establish that Visio reopens the result unchanged.
Each increment should update the capability ledger with its tested subset and
remaining limits. Full 1:1 parity remains a multi-stage target.

## Verification for this increment

- Visio core: 70 test files passed, 1,408 tests passed and 29 optional tests skipped.
  Native rounding and the hash-pinned 60489/60973 theme corpus were enabled;
  the other external corpus environments were not configured.
- Shared Visio UI: 40 test files passed, 433 tests passed and one optional skip.
  JSDOM reported its expected missing-canvas notices. Browser tests were not run.
- Visio documentation: 48 tests passed, including site build and capability ledger.
- Both core TypeScript projects passed. Formatting and diff whitespace checks passed.
- The nine-case native Visio 16.0 generator passed. Re-parsing its native VSDX
  agrees with exported circular radii, sweeps and tangent endpoints (within
  0.0002 inches for SVG coordinate rounding). The real-corpus regression passed.
- Visio ESM/CJS bundles and tsc declarations built. The existing `build:visio`
  shortcut fails in tsup's declaration worker because TypeScript 7 exposes no
  JavaScript `sys` compiler API. Verification disabled tsup's declaration step
  and used the repository's tsc declaration project separately. No dependency
  or build configuration was changed for that unrelated tooling issue.

Changes are local and unreleased. These checks cover this increment; they do not
establish complete native visual, editing or save parity.

## Native open-arrow increment

Open arrow styles 1 and 3 now use native measured proportions and line-width-dependent
sizes; diagonal tick style 9 is newly rendered. The geometry sizing rule lives in
core and shared UI supplies SVG markers for all six bindings. Seven sizes at three
line widths for each style match 63 native Visio 16.0 exports at unit drawing scale.
The compact numerical reference records native glyph coordinates and scale, without
shipping native documents or generated SVG markup. Reproduce it with
`scripts/record-visio-open-arrows.ps1` on Windows with Visio installed.

Native open markers use round joins/caps and have no endpoint setback. Regression
checks compare local glyph coordinates after the native y-down to model y-up
conversion, and cover endpoint anchoring, opacity, orientation and inert SVG export.
These are geometry checks, not a browser pixel equivalence claim. Filled arrow styles
2, 4 and 5 retain their explicit approximate-sizing warning. Non-unit drawing scales,
filled-marker setback, other codes, print behavior and full native pixel comparisons
remain unverified. Full Visio parity is still incomplete.

The subsequent increment adds native curved open style 7, moving open-glyph geometry
into core alongside sizing. The reference matrix now covers 84 native exports. Its
Bézier controls, local coordinate conversion and stroke paint are checked independently
of marker construction. The generator also saves a native-authored 84-shape VSDX for
optional parse-to-render coverage (`VISIO_NATIVE_OPEN_ARROWS_DIR`). Styles 6 and 8 and
filled-end setback remain unsupported or approximate; full parity remains unfinished.

Verification after the curved-arrow increment: 1,573 core tests passed (34 optional
skips), 521 shared UI tests passed (one optional skip), and core/UI typechecking passed.
The shared UI run included native-authored VSDX parsing and rendering for all 84 cases.
Native pixel rendering in browsers, filled sizing and non-unit scales were not tested.
