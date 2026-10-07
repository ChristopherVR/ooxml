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

## Native filled-arrow increment

Filled triangle and curved-base styles 2, 4, 5 and 6 now use measured native glyphs,
sizing and endpoint setbacks on single straight, round-capped connectors at unit
drawing scale, where the setbacks fit. Style 6 is newly supported in this subset.
Native fills have no outline. The original model path is preserved; the shared UI
renders a shortened stroke and positions the marker tip at the original endpoint.

The core helper reproduces the asymmetric begin setback and its small stem extension:
the begin setback is the end setback minus the larger of 0.005 inches and one percent
of the end setback. This rule matches all 84 two-ended reference exports. Another
84 one-ended exports verify the glyph coordinates and end setback. Reference stem
coordinates are checked to the precision of Visio's SVG export. Regenerate with
`scripts/record-visio-filled-arrows.ps1` (add `-BothEnds` for the second matrix), and
set `VISIO_NATIVE_FILLED_ARROWS_DIR` / `VISIO_NATIVE_FILLED_BOTH_DIR` to the output
folders to verify saved native VSDX parsing and rendering. Only numerical evidence
is committed; generated VSDX/SVG files stay local.

Overlapping arrows can make Visio reverse the stem and change setbacks. Curved,
closed and multi-segment paths, overlapping setbacks and non-round caps retain the
approximate fallback for styles 2, 4 and 5. Style 6 reports those unsupported cases
without rendering a guessed marker. Non-unit drawing scales, browser pixel parity,
other arrow codes, routing, text layout, editing and save fidelity remain incomplete.

Verification for this increment: 1,750 core tests passed (34 optional skips), 534
shared Visio UI tests passed, 48 documentation tests passed, and both core typecheck
projects plus shared UI typechecking passed. The UI run enabled the native open,
one-ended filled and two-ended filled VSDX matrices, plus the hash-pinned real
code-5 corpus. Browser pixels and general native save/reopen parity were not tested.

## Native short end-arrow increment

One-ended straight filled arrows now match Visio when the arrow is longer than the
line: the saved glyph keeps its size, its setback becomes zero, and the stem remains
at the original endpoints. At equal length, Visio emits a duplicate origin and a
small stem extension with a correspondingly reduced marker setback. Longer lines
retain the previously verified trim. These cases are resolved in core; the UI uses
the returned stem and marker anchors without changing the document model.

The new matrix covers 168 exports (four styles, three sizes, two stroke weights and
seven length ratios), including the exact boundary. Reproduce with
`scripts/record-visio-short-arrows.ps1`; `VISIO_NATIVE_SHORT_ARROWS_DIR` enables the
saved native VSDX parse/render check. Native SVG stem coordinates are compared at
export precision. Zero-length paths, two-ended overlaps, curved paths, drawing-scale
handling, browser pixels and full parity remain unverified or unsupported.

Verification: 1,919 Visio core tests passed (34 optional skips), 703 shared UI tests
passed, and both core typecheck projects plus shared UI typechecking passed. The UI
run included all native arrow matrices and the hash-pinned code-5 document corpus.

## Native drawing/page-scale increment

The parser now converts cached drawing geometry to physical page inches using
PageScale/DrawingScale, including page bounds, paths, transforms and text frames.
Stroke weights, font sizes, text margins and paragraph indents keep their physical
sizes. Pointer-driven rectangle insertion converts page inches back to drawing
inches; the source-backed edit API continues to accept drawing inches.

Six native-authored pages verify page and rectangle dimensions, filled-arrow stems,
rounded geometry, font size, margins and indents at different scale ratios. Reproduce
with `scripts/record-visio-page-scales.ps1`; `VISIO_NATIVE_PAGE_SCALES_DIR` enables
saved native VSDX parse/render checks against the committed numerical reference.
Synthetic save/reparse and controller undo/redo checks verify scaled insertion and
resize while preserving source scale metadata and untouched parts.

Native save/reopen editing acceptance remains incomplete: resizing the native
reference document fails with EDIT_UNKNOWN_DEPENDENCY because its ShapeSheet
contains dynamic dependencies that the edit index cannot prove safe. This failure
was retained as a remaining editing gap, without dropping formulas to bypass it.
Grouped and foreign placements have normalization code, but broader native reference
coverage and browser pixel comparisons remain outstanding. Full parity is unfinished.

Verification: 1,942 Visio core tests passed (30 optional skips), 705 shared Visio UI
tests passed, 48 documentation tests passed, and both core typecheck projects plus
shared UI typechecking passed. Native arrow and page-scale matrices and the pinned
real-document arrow/theme corpus were enabled. Browser pixel equivalence and native
editing save/reopen acceptance remain unproven.

## Native font-lookup edit increment

The scaled reference document's dependency rejection was traced to FONT("Arial").
Microsoft documents FONT as a font-name lookup, with identifiers that depend on the
system/document, rather than an implicit ShapeSheet cell lookup:
https://learn.microsoft.com/en-us/office/client-developer/visio/font-function.
The dependency analyzer now admits its explicit AST references while retaining
unsupported evaluation. An affected FONT formula still rejects recalculation, and
an INDIRECT argument still rejects dependency proof. Independent font formulas and
cached identifiers are preserved during geometry editing.

A native acceptance probe resized page ID 6, shape ID 2 to 8 by 4 drawing inches and
moved its pin to (6, 4). Visio 16 reopened the edited VSDX with those exact cached
values, FONT("Arial"), DrawingScale 2 and PageScale 1, exported the rectangle and
saved reopened-scales.vsdx. Core parsing reports physical size 4 by 2 inches and
local translation (1, 1). This closes the specific font-lookup rejection above;
general native editing/save fidelity and full parity remain unfinished.

Verification after the font fix: 1,946 Visio core tests passed (30 optional skips),
including the native scale edit test, and strict core typechecking passed. After
rebasing onto concurrent main changes, the relaxed PowerPoint typecheck fails on
unresolved ooxml-core/pptx self-imports in newly moved editor modules; no Visio error
was reported. The earlier scale increment's UI and documentation checks remain valid.

## Source-backed page insertion

The shared page bar now inserts a blank foreground page after the selected page,
through the existing worker, controller and package history. Core owns stable page
IDs, unique names, PageSheet copying, OPC relationships and content-type additions.
Existing page parts and other untouched package payloads remain byte-preserved.
The shared extended-property writer updates the existing Pages title/count category
while preserving other categories and unknown extensions. Undo clamps selection when
a removed page was the last page; redo restores the inserted page.

Native Visio 16 reopened an inserted page in the genuine scale corpus at index 5
of eight pages, with ID 99, literal name "Inserted & Page", zero shapes, cached size
8.5 by 11 drawing inches, DrawingScale 2 and PageScale 1. It saved the drawing again.
The expected physical dimensions are 4.25 by 5.5 inches. Core also accepts a subsequent
rectangle insertion on the new page. The acceptance script is
`scripts/verify-visio-page-insertion.ps1`. Set `VISIO_NATIVE_PAGE_SCALES_DIR` and
`VISIO_NATIVE_PAGE_INSERT_OUTPUT_DIR`, run `visio/edit-pages.test.ts`, then pass its
`core-inserted-page.vsdx` to that script with an output directory. Native binaries
and generated renderings remain local.

Insertion currently requires a separate transaction from shape edits. Page-number
and page-count-dependent formula caches are reported as unrecalculated. Renaming,
reordering, deleting pages, broader PageSheet/background semantics, and complete
native save/render fidelity remain unfinished. Full Visio parity is not established.

Verification: 1,958 Visio core tests passed (30 optional skips), 706 shared Visio UI
tests passed, and all six browser insertion/history/download/reopen cases passed.
The broader shared UI suite also passed 1,908 tests (six optional skips); its combined
script then found no PowerPoint tests matching the Visio filter. Documentation checks
passed all 48 tests. Strict core typechecking and ESM/CJS core plus ESM UI builds passed.
UI typechecking still reports the pre-existing Teams/PowerPoint declaration-resolution
errors from concurrent main changes, with no Visio errors. Native Visio acceptance
was repeated after the extended-property update; reparsing its resaved drawing confirms
page order, dimensions and the refreshed Pages property vectors.

## Source-backed page reordering

The All pages menu now opens Reorder Pages on the shared Office dialog, with a page
list, Move Up/Down, OK and Cancel. The draft becomes one atomic core transaction.
Core changes page metadata and the extended-property title order while keeping page
IDs, existing page payloads and relationships unchanged. Moving to the existing index
returns the original bytes with no changed parts. Insertion and ordering can share a
page transaction; shape edits still require a separate transaction.

The controller preserves the selected page by ID, emits its updated index to the
framework bindings, and checks shape selection against that relocated page. Undo and
redo preserve page identity. Draft cancellation leaves the source unchanged; replacing
the source closes stale drafts, and outdated errors cannot alter a new dialog draft.

A native Visio 16 reference moved page ID 6 to native index 1 via Page.Index. Its order
was 6, 0, 4, 5, 7, 8, 9. The core-produced drawing reopened in Visio with that exact
order, seven pages and Page-4 first, then saved successfully. Native outputs remain
local beside the scale corpus. Reproduce core output by setting
VISIO_NATIVE_PAGE_SCALES_DIR and VISIO_NATIVE_PAGE_ORDER_OUTPUT, then running
visio/edit-page-order.test.ts. The comparison uses the genuine scale-corpus document,
not a synthetic package as native acceptance evidence.

The six-binding browser workflow now also moves the inserted page, checks selected
page identity through undo/redo, and downloads/reopens the ordered drawing. Page-number
and page-count formula caches remain explicitly unrecalculated. Page rename/delete,
broader background semantics and complete native render/save fidelity remain open.

Verification: 1,967 Visio core tests passed (30 optional skips), 708 shared Visio UI
tests passed, and six browser workflows passed across every binding. All 48 documentation
checks passed. Strict core typechecking and core ESM/CJS plus shared UI ESM builds passed.
UI typechecking continues to fail only on the existing Teams/PowerPoint declaration
resolution errors. Native Visio reopen/save was verified independently of browser tests.

## Numeric page-dependent cache recalculation (2026-10-07)

Page insertion and reordering now refresh local numeric PAGENUMBER() and PAGECOUNT()
caches, including their supported static dependency closure and PageSheet cells.
Formula and unit attributes stay intact. Unchanged page contexts preserve their
payloads. Foreground numbering ignores background pages; background page numbers
are zero. PAGECOUNT counts foreground pages only, as specified by Microsoft's
[PAGENUMBER reference](https://learn.microsoft.com/en-us/office/client-developer/visio/pagenumber-function)
and [PAGECOUNT reference](https://learn.microsoft.com/en-us/office/client-developer/visio/pagecount-function).

The numeric evaluator requires explicit document context for these functions.
Affected cycles, unsupported functions, inherited or grouped dependencies and
master/style page-dependent formulas fail the transaction before saved bytes are
returned. Page changes without affected numeric page formulas retain their previous
preservation behavior. This is a bounded local numeric subset: string page functions,
inherited master materialization, general fields and foreground rendering of a
background page's fields remain open.

Native Visio 16.0 was exercised with two foreground pages and one background page.
An actual Page.Index reorder left the moved page's PAGENUMBER user-cell cache stale,
even after save/reopen and reassignment of the identical formula. Toggling to a literal
and back forced native recalculation. The core output matches that freshly evaluated
reference by stable page ID, including transitive User.Number+User.Count caches.
The core-produced output then opened and saved in Visio with every expected value.
An inserted foreground page was also accepted and saved: all existing foreground
and background shapes reported a count of three with correct page numbers and sums.
The change implements the documented numeric result, rather than reproducing this
native stale-cache behavior.

Reproduce with scripts/record-visio-page-formulas.ps1, set
VISIO_NATIVE_PAGE_FORMULAS_DIR to its output directory, run
visio/edit-page-formulas.test.ts, then run the script again with -CoreOutputPath
pointing to core-reordered.vsdx. Native binaries remain local; the compact metrics
are committed in `visio/__fixtures__/page-formulas-native.json`. To verify insertion
too, pass -CoreInsertedOutputPath pointing to core-inserted.vsdx.

Verification: 1,977 Visio core tests passed with the native numeric/arrow/scale
references enabled (30 optional skips); strict core typechecking and Visio ESM/CJS
builds passed. All 702 shared Visio UI tests passed (six optional native skips),
all six browser insertion/reordering/history/save workflows passed, and all 48
documentation checks passed.
