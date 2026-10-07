# Verification record

Status: local Windows development evidence, 2026-10-03. Full Microsoft Visio parity is not established.

## Native local group API rotation, 2026-10-08

Two owned Visio 16 captures record source and rotated parent/descendant cells,
independent XYToPage matrices and native SVGs:

- Custom-pivot group, 30 degrees: visio-group-rotation-670cfd645e7c4e618ad10247d89459b4.
- Two-level group, 210 degrees, drawing-to-page ratio 0.5:
  visio-group-rotation-9d5f9d76db764ae289bbf484b1e18bae.

The parent Angle assignment retains child local geometry and formulas while
changing descendant world positions. The existing parser, shared affine composer
and renderer already match those native matrices. The edit extends the existing
geometry transaction with a bounded local-tree proof and a parent-Angle exception
to the existing dependency evaluator. Affected descendant formulas still cause
atomic refusal. Masters, glue, foreign/1D descendants, inherited/error transform
caches and rotation protection are covered by refusal regressions. Other group
commands remain outside this proof.

The final post-rebase Visio core run passed 2,239 checks, with 145 optional checks skipped.
The final focused group run passed all 12 checks, including both native captures
and the additional inherited/error-cache regressions. Core/UI builds and root,
UI and viewer typechecks passed. Twelve browser workflows cover the public API,
worker, undo/redo, byte-exact history, public export and saved reload across six
bindings. Every authored parent/descendant SVG matrix is compared to native at
12 decimal places; actual screen matrices are separately checked at five decimal
places because Chromium rounds them. The initial ten-decimal screen check failed
on a coefficient difference around 4.4e-8, prompting the separate authored check.

Native measurement helpers are shared with the existing line recorder; its
post-extraction rotation capture completed successfully in
visio-line-movement-f107345af7b54d33be31431da0c8a439. Owned capture apps quit.
Native assignments are not literal native GUI gestures. Group controls/pointer/
menus, other group edits, broader formulas/display modes, exact paint and native
Visio reopening remain unverified. The post-rebase source import audit found
zero .js TypeScript imports and zero unresolved relative imports.

## Native static rotation formulas, 2026-10-08

Four owned Visio 16 captures supply rectangle and ellipse references:

- Absolute rotation from Width-dependent Angle to 90 degrees, custom pivots:
  visio-line-movement-3da4307f83444f7f8b6d66ec953b6e26.
- Exact current-value assignment from Width/Width*30deg, custom pivots:
  visio-line-movement-aa377d6a6aef40608a56431ffb6f4254.
- Right quarter-turn from Width-dependent Angle on a drawing-to-page ratio of 2:
  visio-line-movement-1ab9a6be262042129c4ce434e991b75c.
- Left quarter-turn from Width/Width*330deg:
  visio-line-movement-65f55f5b0dc349a08849ef0ba76fcbad.

The recorder's shared SourceAngleFormula setup and RotateToSourceAngle mode
preserve the native capture lifecycle. Both absolute assignment and per-shape
Selection.Rotate replace the unguarded formula. Assigning the exact current
native cached angle still removes F while retaining V and the pose. Owned
applications quit, and the original user instance remained intact. These are
native API recordings, not literal native pointer/menu interaction or reopening.

Rotation now shares edit-transform-formula.ts with flips. The bounded source
proof, cell writer, dependency recalculator, worker and history paths are reused.
Explicit current-value formula replacement creates one dirty page; another
same-value command after the formula is gone preserves the existing bytes.
Core checks compare formula removal, cached poses and unrelated package parts.
Quarter-turn regressions also compare independent native XYToPage matrices,
including physical scale normalization, rather than relying only on core reparse.

The full Visio suite passed 2,263 checks, with 109 optional checks skipped where
captures were not supplied. A focused quarter-turn run passed all 15 checks with
both new and five earlier native captures. Builds for Visio core and UI passed,
as did root, UI and configured viewer typechecks. The source import audit found
zero .js TypeScript imports and zero unresolved relative imports.

There are 60 distinct passing browser workflows across six bindings: 12 control
rotations, 12 pointer rotations, 24 menu quarter-turns and 12 current-value API
assignments. The first run passed the 48 pose-changing workflows but the 12
current-value tests exported during asynchronous undo. Waiting for completed
controller/history state fixed the test timing; the complete 24-case control/API
rerun passed. Current-value API checks prove generation/dirty/history changes,
byte-exact undo restoration, redo and public export/reload despite unchanged pose.
Pointer checks also cover cancellation and preview/source separation. Other
native pointer behavior, exact paint, inherited/unresolved formulas, groups,
masters/glue and native Visio reopen acceptance remain unverified.

## Native static flip formulas, 2026-10-08

Three owned Visio 16 captures exercise rectangle and ellipse source formulas:

- Width-dependent Angle, custom pivots: visio-line-movement-3fadfcaeafab4a9c986a7f4ff51df775.
- Width-dependent FlipY, drawing-to-page ratio 2: visio-line-movement-e122b2d52e6a40399b44b4e5d9701cd9.
- Width-dependent zero Angle: visio-line-movement-aa15358b12914a1585578a51369a365f.

The recorder accepts FlipAngleFormula and FlipFlagFormula through the same
owned capture and cleanup lifecycle. Selection.Flip replaces the nonzero Angle
and toggled flag formulas with constant values. It retains the source formula
when Angle already evaluates to zero. The original user instance was untouched;
these recordings do not prove native menu/pointer gestures or Office reopening.

The bounded evaluator extracted from edit-recalculate.ts is now reused by a
read-only source proof. It checks dependency availability, cycles, units and
cache agreement before allowing a static flip formula to be replaced. Existing
post-edit dependency recalculation and protections remain in use. Other geometry
commands retain their existing formula admission. Synthetic regressions cover
outgoing dependent cache updates with preserved formulas, inconsistent caches,
missing/cyclic references and formula redirection. Native comparisons also check
formula replacement/retention and byte preservation of unrelated package parts.

The full Visio core suite passed 2,264 checks, with 97 optional checks skipped.
After adding the zero-angle capture and explicit XML assertions, the focused
flip/recalculation run passed 45 checks, with 16 older optional capture checks
skipped. The two initial captures passed 24 browser scenarios; the zero-angle
capture passed another 12, across all six bindings. These actual menu workflows
compare saved poses, undo/redo and public export/reload. Core/UI builds and core,
UI and configured viewer typechecks passed.

Inherited transforms, groups/masters/glue, unknown or inconsistent formulas,
exact paint and Microsoft Visio reopen acceptance remain unverified.

## Fully blocked flip history and feedback, 2026-10-08

Two fresh owned native Visio 16 captures verify fully blocked transforms:

- Horizontal, guarded Angle and FlipX: visio-line-movement-6ec8fefd514941a188748a1eba7bfab3.
- Vertical, LockRotate and guarded FlipY: visio-line-movement-8d63c8d252544283a0c52ab9ce6a9d8a.

Both include rectangle and ellipse references, with a custom pivot or a
0.5 drawing-to-page ratio. Native assignments retain the protected cells and
poses. Capture applications closed; the original user instance stayed intact.
These captures do not represent native pointer/menu gestures or Office reopening.

All 24 browser scenarios passed across six framework bindings. Actual viewer
menu commands retain source bytes, document generation, clean state and history.
A move followed by undo creates a redo branch; another blocked flip retains it,
and redo still applies the move. Public source export and reload retain the
unchanged document. Feedback now reports "No changes were made." through the
existing controller generation and command announcement path.

The associated core run passed 76 checks covering all eight optional native flip
captures, rotation and master-move regressions. The final UI run passed 20 command/menu
checks. Core and UI builds, root and UI typechecks, and the configured Visio viewer
typecheck passed. Running the child e2e configuration directly from the root
cannot resolve its viewer-local vite/client type; the viewer typecheck includes
the same e2e sources and resolves those dependencies.

Dependent or inherited transforms, groups/master/glue, broader protection
combinations, exact paint and Microsoft Visio reopening remain unverified.

## Native endpoint cells after an explicit Width override, 2026-10-08

The existing line recorder accepts combined ResizeWidth=4 and MoveEndpoint
Begin/End. Two fresh owned InvisibleApp captures cover horizontal, 30-degree,
vertical and reversed lines:

- Begin: visio-line-movement-aff296c6167b41938cf55bb4ab28cd5a
- End: visio-line-movement-df4a2be6a1774ba08ce2976205014061

Both reside in the local temporary directory. Source is resized.vsdx; native
endpoint assignments are in endpoint.vsdx and endpoint-page.svg. The applications
quit after recording; the user's pre-existing Visio instance stayed untouched.
No native open/reopen attempt or user mouse interaction is represented here.

The native assignments retain Width=4 and its constant formula. Opposite raw
endpoint cells stay fixed; midpoint and derived angle follow the changed cells.
Displayed geometry remains four inches long, so its begin and end differ from
the raw cells by more than 0.1 inch in every case. Merely admitting static Width
in the current pointer command would therefore produce a different drop result
from its fixed-opposite-endpoint geometry contract. No admission was weakened
and no new endpoint edit support is claimed.

Two optional core regressions passed, enabled by VISIO_NATIVE_LINE_WIDTH_BEGIN_DIR
and VISIO_NATIVE_LINE_WIDTH_END_DIR. They check native cached Width/opposite
cells, parsed poses against XYToPage to 12 digits, visible-versus-cell endpoint
divergence and the retained source-backed refusal.

All twelve browser scenarios passed across six frameworks. Saved native poses
match DOM transforms to 12 digits; displayed begin/end positions match the
independent native SVG measurements to three decimals (export rounding).
Each scenario then loads the resized source, attempts all four endpoint commands,
checks their explicit refusals and preserves the complete page model and exported
source bytes exactly. Core strict/PowerPoint and viewer type checks passed.

This is viewing and refusal evidence, not native endpoint editing acceptance.
Native interactive endpoint behavior after Width overrides must be recorded
before implementing that interaction. Existing derived-Width endpoint editing,
paint and save/reload evidence remains valid at its documented scope.

## Pointer endpoint gradient paint after save/reload, 2026-10-08

The native begin/end paint captures below now also drive actual pointer gestures
in all six frameworks. Set VISIO_NATIVE_GRADIENT_ENDPOINT_GESTURE=1 alongside
VISIO_NATIVE_GRADIENT_RASTER_DIR to run this mode. Geometry and paint tests reuse
one dragLineEndpoint helper, including handle selection, drawing/page conversion,
mouse movement, preview and release checks. Shared gradient benchmark setup
handles API edits, Width-cell edits and pointer edits without separate painters
or copied pixel comparisons.

Each pointer case checks undo/redo and usable handles, saves through the public
VSDX API, compares the saved native geometry/paint model and reloads before live
and portable SVG raster comparisons. Pointer poses, widths and physical linear
gradient endpoints use the existing four-decimal gesture contract. Canonical
straight geometry reuses the core eligibility helper; stops, opacity,
interpolation and non-gradient style remain exact. API model comparisons retain
12 digits and exact geometry/style. Reference PNGs, registration, contour masks,
minimum pixel counts and all paint-error bounds remain unchanged.

All twelve pointer paint scenarios passed, producing 96 interior comparisons.
Both endpoint groups reached maximum channel error four. Begin/end maximum
means are 0.871514 and 0.884248 (rounded up); minimum pixel counts are 4,732 and
1,622. Observed pose differences are below 4.43e-7 and 1.15e-7 respectively.
The initial pointer diagnostic failed the API registration bound of 1e-9;
pointer registration now explicitly matches its existing four-decimal contract
(5e-5), while API registration retains 1e-9. No native paint tolerance was raised.

The shared-helper extraction passed 18 existing geometry workflows: both
endpoints across all frameworks and both endpoints at three scales in vanilla.
Six end-endpoint API paint workflows also passed with the original tighter
registration/model checks. The Width-cell paint setup is checked in vanilla
against a fresh native Width-four capture,
visio-gradient-raster-2f0b4945edf543cb88544a97fa474fb4, paired with the Width-two
source visio-gradient-raster-f0765a3a531a4a299f6a7484ceec5085.
Viewer type checks pass. This establishes the captured gestures and saved viewer
paint within current interior bounds, not exact pixels, native interactive
snapping, keyboard manipulation, other paint directions/angles or native Office
reopen acceptance.

## Native endpoint gradient paint after save/reload, 2026-10-08

The existing gradient recorder now accepts MoveEndpoint for native unglued
straight lines with derived Width/Angle and linear stroke paint. It assigns
the selected endpoint by (+0.75, -0.5) drawing inches, records native XYToPage
poses/extents and exports unchanged native PNG/SVG references. After saving
the edited native document, it restores the selected coordinates and saves a
matching source document with the same shape IDs and paint. The owned invisible
Visio instances restored raster settings and quit; the user's instance stayed
untouched.

Captures in the local temporary directory:

- Begin: visio-gradient-raster-f45c298ce47e48c2b432a495fbc3c985
- End: visio-gradient-raster-ce5bd91d257e48d6ae430b97741de8a2

Each contains four native cases: two/three stops, opaque/translucent paint and
a 45-degree saved gradient angle. Both endpoint assignments change line length
and orientation. The browser benchmark reuses its existing resize/edit stage
and shared painter; no product coordinate or paint implementation changed.
For endpoint captures it applies the shared command, verifies complete page
models through undo/redo, exports the public VSDX bytes, compares native saved
poses/widths to 12 digits and geometry/style exactly, and reloads the saved bytes
before performing the live and portable-SVG native PNG comparisons.

All twelve framework scenarios passed, producing 96 interior comparisons.
Begin/end maximum channel errors are both four; maximum mean errors are
0.871461 and 0.884248 (rounded up). Minimum compared pixels are 4,732 and 1,622;
pose differences are below 3e-16. The original maximum-seven and mean-one/1.5
linear gates and contour exclusions remain unchanged.

Two optional core regressions, enabled by
VISIO_NATIVE_GRADIENT_BEGIN_ENDPOINT_DIR and
VISIO_NATIVE_GRADIENT_END_ENDPOINT_DIR, passed all eight cases. They compare
native poses, geometry and paint, preserve the unrelated control and every
unchanged package part byte-for-byte, and confirm only page1.xml changed.
Core strict/PowerPoint type checks and viewer type checks passed. A fresh
unedited native line-paint baseline was also rerun in vanilla using
visio-gradient-raster-f0765a3a531a4a299f6a7484ceec5085. The first baseline attempt
stopped before import because its historical temporary capture was missing.

This proves the captured API edits and saved viewer output within existing
interior error bounds. Exact pixels/contours, pointer gesture paint, other
gradient directions/angles, glue/routing and native Office reopen acceptance
remain unverified.

## Scaled-page endpoint editing, 2026-10-08

Six fresh native Visio captures cover drawing-to-page ratios 0.5, 2 and 3,
both endpoint assignments and four orientations per capture: horizontal,
30-degree, vertical and reversed. The recorder accepts DrawingScale/PageScale
and records their native cached values alongside cells and XYToPage poses.
It also exports each resulting page as native SVG.

Optional capture directories, named by VISIO_NATIVE_LINE_SCALED_*_DIR:

- HALF_BEGIN: visio-line-movement-4066bcfeeec64a1c84be6b685edb710f
- HALF_END: visio-line-movement-16f211dfb5c646ca9ded30a7b1116ef9
- DOUBLE_BEGIN: visio-line-movement-6d135280ddbf4ef88051637e37e61886
- DOUBLE_END: visio-line-movement-418948da1efd48e7a5d455c8968db455
- TRIPLE_BEGIN: visio-line-movement-0e203678cc7e4613a0ab849f6b0f615c
- TRIPLE_END: visio-line-movement-699d528dcfcd47138abd1e4cfb269766

The existing shared page-point and page-to-drawing conversion passed without
production changes. Core tests compare native caches and scaled XYToPage poses
to 12 digits, check parsed scale metadata against native cached values and
preserve the pages.xml payload byte-for-byte. The full core Visio run passed
2,186 tests with 60 unrelated optional skips; both core TypeScript projects
and the viewer TypeScript check passed.

Across two browser runs, 72 distinct scaled workflows passed: 36 pointer and
36 API scenarios, six frameworks per endpoint/scale combination. The final run
passed 48 scenarios, including all 36 scaled pointer workflows with visible
handles after redo and the 12 ratio-three API workflows. Each scenario edits
four lines and checks history, saved copies and reload.

A shared browser helper measures endpoint positions directly from native SVG
path geometry and matrices, independently of the core parser. Native physical
paper dimensions match to 12 digits; saved pointer endpoints match to three
decimal places, accounting for native SVG export rounding. Existing pointer
pose/length checks retain four decimals and API checks retain 12 digits.
These measurements establish endpoint geometry for the captured cases, not
pixel/paint parity. Other scales/unit conventions, native snapping, keyboard
manipulation, gesture paint and native Office reopen acceptance remain open.

## Canvas endpoint dragging, 2026-10-08

Selected visible top-level straight connectors with source bytes now expose begin
and end pointer handles. The new core UI helper checks visual geometry eligibility;
it does not certify source formulas or protection. Connected lines are excluded.
Source-backed endpoint admission remains authoritative and can refuse a drag.
DOM controls live in the shared ooxml-ui/visio element, with no framework copies.

The gesture reuses pagePoint from the rectangle tool with unsnapped/unbounded
coordinates and the existing page-to-drawing conversion. Rectangle creation keeps
its snapping/bounds defaults. Handles sit in an SVG overlay above all shapes:
the native four-line fixture shares a begin point, and ordinary shape stacking
initially hid the selected line's hit target. SVG matrices provide the placement
and screen scale, keeping the circles five pixels in radius through zoom.

The preview is separate DOM; one core transaction commits on release. Escape,
pointer cancellation, lost capture, page/document/selection replacement, tool
changes and disposal cancel the preview. Clicks with no drag retain original
bytes/history. The full source-backed protection/formula checks stay in core.

All twelve pointer workflows passed: begin and end editing across six frameworks,
four native orientations each, preview, native pose/length comparisons, undo/redo,
download and reload. They use the prior Begin/End native COM captures recorded
below. Browser pointer-coordinate comparisons use four-decimal tolerance; native
API/cache comparisons still use 12 digits. This verifies gesture results against
native cell assignments, not native interactive snapping or handle appearance.
Cancellation/click byte preservation and the existing rectangle draw/delete/history
workflow also passed: 14 browser scenarios total. An initial run used stale UI
bundles; after rebuilding, the shared-point stacking failure exposed the need for
the overlay. Both were corrected before the final passing run.

The core Visio suite passed 2,180 active tests with 60 unrelated optional skips;
eleven shared UI handle/command tests and all core/shared UI/viewer typechecks pass.
Core bundles and shared UI build pass. Source imports remain extensionless.
All six demo builds, clean core package entry-point imports and packed viewer
ESM/declarations/external-consumer/parser/edit-worker checks pass.
Native snapping, keyboard endpoint manipulation, scaled-page gestures, gradient
paint after dragging, broader formulas, glue/routing and native Office reopening
remain open.

## Native endpoint-cell editing, 2026-10-08

The existing line recorder now takes MoveEndpoint=Begin or End, assigning the
selected endpoint a delta of (0.75, -0.5) drawing inches after translation.
It records endpointBefore/endpointAfter caches and formulas, XYToPage basis
points and endpoint.vsdx. Begin capture:
visio-line-movement-b6f97cba81e0467f8472f77b833b6d19. End capture:
visio-line-movement-c278e3a49a4640e1a2b90abe387d04a6. Both contain horizontal,
diagonal, vertical and reversed lines. Native assignment retains the derived
midpoint, length, angle and local-pin formulas; the other endpoint stays fixed.
No native canvas dragging or protection interaction was recorded.

The new move-line-endpoint command reuses local line admission, package scope,
protection/style resolution and bounded dependency/cache writeback. The duplicated
package/master changed-cell mappings now share geometryChangedCells. Worker
command snapshots retain endpoint identity without caller extras, and page-inch
conversion reuses the move coordinate path. Explicit native transform formulas
and fresh caches are required. Endpoint GUARD, affected locks, unproven formulas,
static Width overrides, flips, coincident endpoints and glue continue to refuse.
Length changes reuse the existing resize geometry dependency proof; a regression
rejects absolute endpoint coordinates that cannot follow Width. The saved result
also repeats shared line admission, guarding against invalid dependent geometry.

Core tests using VISIO_NATIVE_LINE_BEGIN_DIR and VISIO_NATIVE_LINE_END_DIR match
all eight native cell sets, saved models and six-coefficient poses. Native XML
rounds some caches differently from the numeric interpreter, so complete-model
numeric comparisons use 12-digit tolerances; nonnumeric fields remain exact.
The test loader accepts Windows PowerShell's UTF-8 BOM; the recorder now writes
UTF-8 without a BOM on both PowerShell runtimes.

All twelve browser API scenarios passed, covering both endpoints across six
frameworks, production edit workers, native pose comparisons, undo/redo, saved
model comparison and reload. There are no endpoint handles or dedicated controls
yet. The core Visio suite passed 2,173 active tests with 60 unrelated optional
skips. Both core typechecks, shared UI/viewer typechecks, core bundles, six demo
builds, all core package entry-point imports and packed viewer consumer/worker
checks passed. The source import audit found zero relative .js specifiers or
unresolved imports. Native Office reopening, canvas dragging, gradient paint
after endpoint changes, broader 1D formulas and routing/glue remain unverified.

## Native Width-cell line resizing, 2026-10-08

Native Visio Width.ResultIU=4 on horizontal, diagonal, vertical and reversed
DrawLine shapes replaces the endpoint-length Width formula with a cached width.
Endpoint cells, midpoint formulas and Angle remain unchanged; Width-dependent
local geometry and LocPinX update. This operation differs from endpoint dragging.
The ResizeWidth option in the existing line recorder captures the before/after
cells, resized.vsdx and native XYToPage basis points. The final capture is
visio-line-movement-19edef4a85de4b73849c2e8a47f103e5; the optional core test uses
VISIO_NATIVE_LINE_RESIZE_DIR. Core Width-cell resize matches all captured cells,
native saved models and six-coefficient poses.

The same line admission now serves movement and resizing, accepting native
derived Width or a proven static override. Cell/unit checks, protection,
absolute dimension-dependent geometry admission and dependency/cache writeback
are reused. Width's unguarded native length formula may be replaced by this
explicit command; midpoint and angle formulas remain intact. Requested Height
must stay zero, and the rotation pin stays fixed. Moving or resizing the saved
constant-width line remains admitted. Guards, locks, stale caches, incompatible
units, pose-changing dependencies and unsupported geometry continue to refuse.
The shared controls use the existing selected-shape lookup to permit zero Height
for a selected height-zero connector; rectangle creation retains positive sizes.

Six native Width-cell resize browser scenarios passed with worker execution,
controls, double-precision serialized SVG pose comparisons, undo/redo,
download/native-model comparisons and reload. The twelve movement/deletion
scenarios also passed. The pose helper initially used a CSS parser on SVG matrix
syntax, then SVGMatrix's float32 getters; it now checks emitted coefficients
without reducing the 12-digit comparison precision.

The native gradient recorder's LineShapeWidth option captures Width=4 stroke
paint. visio-gradient-raster-78cef432133c4999907774b97f15a7b7 contains four
two/three-stop opaque/translucent cases. The existing six-route raster test
loads the Width=2 source from
visio-line-1d-gradient-control-a20b437e4ca54d2e99a7a963f8db10f5 and applies real
Width-cell edits before comparison. VISIO_NATIVE_GRADIENT_RESIZE_SOURCE_DIR
selects this paired source while VISIO_NATIVE_GRADIENT_RASTER_DIR selects the
native Width=4 reference. All 48 live/export comparisons meet the unchanged
interior gates: maximum channel error four, maximum mean below 0.867, minimum
1,056 compared pixels and exact measured pose agreement. The twelve star and
rotated-polygon groups have no cases in this capture and are skipped.
Unverified-gradient-raster remains; exact pixels, contours, other resize paints,
endpoint dragging, broader resize formulas, glued routing and native Office
reopening remain open.

The core Visio suite passed 2,154 active tests with 60 unrelated optional skips.
Both core typecheck projects, shared UI and viewer typechecks passed. The shared
geometry/control tests passed, as did every clean core package entry-point
import. A TypeScript source import audit found zero relative .js specifiers and
zero unresolved relative imports.
The Visio core bundles, shared UI build and all six production demos build.
Packed viewer ESM, declarations, external consumer and production parser/edit
worker checks pass.

## Line deletion and endpoint protection, 2026-10-08

Unreferenced local 1D leaves now delete through the existing shared deletion
path: shape ownership admission, protection resolution, LockDelete, bounded
ShapeSheet reference analysis and Connect checks. Whole-leaf removal does not
rewrite a transform or geometry, so height-zero, point-sized or curved local
line geometry does not require translation admission. Master, group, foreign
and referenced or connected deletion remain excluded. Untouched package part
payloads remain byte-preserved.

The native line recorder's DeleteAfterMove option adds a control rectangle,
saves the source and translated package, deletes the four native DrawLine
shapes and saves deleted.vsdx. The local capture is
visio-line-movement-1ff35609e7df46ef934e2afbc5aa1bd0. Core deletion matches
the native surviving model exactly. The optional capture test uses
VISIO_NATIVE_LINE_DELETION_DIR. Regular tests cover zero-sized/curved leaves,
endpoint GUARD, LockDelete, inherited protection, formula references,
Connect records, atomic refusal and untouched part payloads.

Movement also now checks local LockBegin/LockEnd and resolves their style
ancestry using the existing shared protection helper. Microsoft's
[LockBegin reference](https://learn.microsoft.com/en-us/office/client-developer/visio/lockbegin-cell-protection-section)
and [LockEnd reference](https://learn.microsoft.com/en-us/office/client-developer/visio/lockend-cell-protection-section)
define these as fixed 1D endpoint locations. Static active locks, stale
protection caches, unproven dependency formulas and inherited active locks
refuse translation. Deletion does not overwrite the locked endpoints.
These lock regressions are source-backed checks against the documented cell
semantics, not native interactive UI recordings.

All twelve native-line-movement.spec.ts scenarios passed: movement plus
deletion on six frameworks, real worker execution, keyboard selection,
controls, undo/redo, download, saved-model comparisons and reload. The core
Visio suite passed 2,146 active tests with 60 unrelated optional skips; both
core typecheck projects and viewer types passed. Production viewer/demos builds,
packed viewer ESM/types/consumer/worker checks and every core package entry-point
import passed. Native Office reopening,
glued cascading deletion, broader 1D movement/resize and full parity remain
unverified or unsupported.

## Native line translation oracle, 2026-10-08

`scripts/record-visio-line-movement.ps1` creates four real one-dimensional
DrawLine shapes: horizontal, 30-degree diagonal, vertical and reversed
horizontal. It saves original.vsdx, translates each endpoint by (2.25, 1.5)
internal inches, then saves moved.vsdx and records native formulas and caches.
It owns and closes an InvisibleApp without changing raster settings.

The native formulas are `(BeginX+EndX)/2`, `(BeginY+EndY)/2`,
`SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)` and
`ATAN2(EndY-BeginY,EndX-BeginX)`. Translation preserves those formulas,
length, angle, local rotation pins and flips. Both saved packages agree with
the recorded caches. The existing shared formula interpreter matches the
native midpoint, length and angle for all four orientations before and after
translation, including dimensional square roots and ATAN2 argument order.

`native-line-transform.test.ts` passed five tests with
VISIO_NATIVE_LINE_MOVEMENT_DIR pointing to the local capture
visio-line-movement-a6330ef9743145c096db40a901d12302. The capture test is
optional; four measured formula regressions run without native Visio.
The shared editor now admits these unglued local straight-line translations.
It reuses cell access, protection checks, endpoint edit admission and the
existing bounded dependency graph and cache writeback. A per-operation proof
allows the target line's local cells to participate without relaxing group,
master, foreign, deleted, glued or inherited-cell checks. Midpoint and length
formulas remain unchanged; length, angle, local pins and flips must remain
fixed. Post-translation comparisons allow floating-point roundoff. Endpoint
GUARD/reference overwrites, stale caches, curved geometry and pose-changing
dependencies fail atomically. Width-cell resizing is now admitted as recorded
above; endpoint dragging and broader 1D movement formulas remain unsupported.
Unreferenced local line deletion is now admitted as recorded
above; glued or referenced deletion remains unsupported.

The additional native editing test matches all saved transform caches to the
native after-values. Fifteen regular editing regressions cover four orientations,
round-trip restoration, external endpoint references, locks, GUARD, no-op,
stale caches, geometry/unit refusal and pose-changing dependencies. Along-line,
perpendicular and rotated oblique native gradient captures also preserve paint
through direct movement. All six native-line-movement.spec.ts scenarios passed
using the real worker, geometry controls, keyboard selection, undo/redo,
downloaded package comparisons and reload. Core Visio tests passed 2,135 active
cases with 59 unrelated optional skips. Core/UI types, core/UI builds and viewer
types passed.

A new four-line source/candidate native reopen attempt used the fresh local
directory visio-line-reopen-8d2b94cc7e874316bd5b4f9d80126063. OpenEx(200) again
stalled while opening the untouched native original, before reaching the
candidate. The probe worker and invisible Visio process were stopped after
matching their recorded process IDs and start ticks; the user's instance stayed
open. No native repair-free acceptance or native resave evidence was obtained.
Native reopening of editor output remains unverified.

## Height-zero gradient strokes, 2026-10-08

Native DrawLine captures establish Width=2, Height=0 and local M 0 0 L 2 0
geometry. The former positive-width-and-height gradient admission rejected these
complete caches and produced solid line fallback. The existing shared parser now
admits nonnegative one-dimensional stroke sizes with one nonzero axis, retaining
the normalized cached angle until the physical stroke descriptor includes line
width. The existing endpoint helper, stop sampler, SVG painter, snapshots and
resource accounting handle them. Zero-area fills and point-sized strokes remain
rejected; nonlinear stroke directions remain diagnosed.

The recorder accepts LinearShape=line and draws with the native API. It records
the two-vertex outline, native extents and XYToPage pose using the same registration
and stroke-interior mask as rectangle strokes. Three genuine four-case captures
cover two/three stops and opaque/translucent paint at zero and 90 degrees, plus
225-degree paint on a line rotated by 30 degrees. Local directories are
visio-line-1d-gradient-control-a20b437e4ca54d2e99a7a963f8db10f5,
visio-line-1d-gradient-perpendicular-control-614d5a4fcfd44e9b92f67df26cb4aab6 and
visio-line-1d-gradient-oblique-control-b204ea91801040f9b45febb650333632.
Owned invisible applications restore raster settings and quit.

All six routes pass live and portable SVG comparisons: 144 comparisons against
unchanged native PNGs, with unchanged error bounds and independent pose checks.
Along-line paint reaches maximum channel error four and mean below 0.868 over
2,580 pixels per comparison. Rotated oblique paint reaches maximum five and mean
below 0.925 over 2,234 pixels. The perpendicular cases also satisfy the existing
maximum-seven and mean-below-one for opaque two-stop paint, or 1.5 otherwise,
bounds. These nonzero errors do not establish exact native pixel parity. Earlier
four-line setup captures are diagnostic; final captures also include the control
described below.

Direct movement of the native lines is still rejected by geometry admission.
The recorder adds an unrelated native 2D rectangle so preservation tests exercise
a genuine page-XML edit through move/save/core reparse without bypassing that
guard. Tests explicitly assert the line-movement rejection, move the 2D control
and verify all four gradient models remain intact in each final capture.
Generated regressions check cached angles and continued fill/point rejection.
This is paint viewing and source preservation evidence, not line editing parity
or native Office reopening. Routing/glue, arrows, other poses and exact contours
remain open; unverified-gradient-raster remains on every accepted gradient.

Core Visio tests pass 2,115 with 58 optional skips. Core/UI/viewer typechecks,
fast core/UI builds and the 62 relevant UI regressions pass.

## Physical oblique fill projection, 2026-10-08

Two independent native raster captures cover 45-degree fill and 225-degree fill
on a 30-degree rotated rectangle, each with two/three stops and opaque/translucent
paint. Local directories are
visio-fill-gradient-oblique-26c8157ed1c84e76963dffa1e62a2c9e and
visio-fill-gradient-oblique-rotated-af4ad7a04ec34d41bff6fbcb48170211.
The recorder owns its invisible application, restores raster settings and quits.

All six routes pass live and portable SVG comparisons against unchanged native
PNGs: 96 comparisons. Maximum channel error is four; means are below 0.773 and
0.785 respectively. Comparisons cover 34,816 pixels in the unrotated capture
and at least 8,350 in the rotated capture. Existing error bounds, native extents
and independent pose checks are unchanged. Exact native pixel parity is open.

Before the change, normalized native-SVG rotation reached maximum channel error
81 for the 45-degree opaque two-stop fill. Saved linear fills now call the existing
physical linearGradientEndpoints helper for oblique angles, as themed gradients
already do. Stops and source XML remain intact; the existing scaling, SVG paint,
snapshot and sigma/gamma paths consume physical endpoints. Generated tests check
analytic endpoints and parser-to-SVG coordinates. The shared optional native
regression preserves all four models from each new capture through move/save/core
reparse, alongside the existing 68-, 84- and 32-case captures. Native-SVG export
rotation is no longer the oblique saved-fill rendering oracle.

A smaller native save/reopen probe used the genuine four-shape source and an
edited candidate, in local directory
visio-small-reopen-0f9404f5-0eb5-4fe4-b1de-e0aa83a17788. It hit a 25-second timeout
while opening the original source, before opening or examining the candidate.
The probe's owned processes were cleaned up; the user's existing Visio instance
remained running. This establishes no native repair or candidate acceptance
result. Native Office save/reopen acceptance remains unverified.

Core Visio regressions passed 2,114 tests with 54 optional skips; UI passed 719
with seven skips. Core/UI/viewer typechecks and fast builds pass. All gradients
retain unverified-gradient-raster. Other shapes/angles, groups/flips, contour
pixels and exact colors remain open.

## Saved linear stroke paint, 2026-10-08

Additional native captures cover 45-degree paint and 225-degree paint on a
30-degree rotated rectangle. The local directories are
visio-line-gradient-oblique-4d9eef5cbed0437896e6123406f76233 and
visio-line-gradient-oblique-rotated-35eb4673d98f4407b0d56671c508f53e.
Each has four two/three-stop opaque/translucent cases. All six routes pass live
and portable SVG interior comparisons: 96 additional comparisons, maximum
channel error five, means below 0.965 and 0.971 respectively. Independent native
pose checks pass; source PNGs and existing error bounds are unchanged.

The original normalized SVG-style rotation reached maximum 94 and mean 17.52
on the 45-degree opaque two-stop stroke. Native raster paint uses physical
projection through the stroked rectangle. The descriptor now delegates to the
existing linearGradientEndpoints helper using dimensions plus physical line
width, and translates the result by half that width. Source caches keep their
normalized angle and stops; the renderer consumes a separate physical descriptor.
Analytic and parser-to-SVG regressions check the physical endpoints, scaled
dimensions, unchanged stop alpha and source preservation. Both native captures
retain all four gradient models through move/save/core reparse.

The fast build:visio command previously failed in tsup's legacy declaration
plugin with TypeScript 7. It now shares the full build's command runner and strict
declaration compiler, bundles both Visio entry points, and reuses the declaration
import rewriter. Full and fast builds, strict/PowerPoint and UI typechecks,
clean-package imports of every entry point and the NodeNext declaration test pass.
Other angles, nonrectangular stroke paint, groups/flips, contours, arrowheads,
nonlinear paint, editing controls, exact native pixels and native Office reopening
remain unverified. The raster diagnostic remains.

The native recorder shares stop setup between Fill section 249 and Line section
248 and accepts Paint and GradientAngle. Two genuine four-case rectangle captures
cover two/three stops and opaque/translucent paint at zero and 90 degrees. Their
local directories are visio-line-gradient-ba051c12890c4b4c94be53fc88bba162 and
visio-line-gradient-vertical-2d88777d33234f9a9eadf1774f6974a5.
Owned InvisibleApp instances restore raster settings and quit.

Unchanged native PNGs are registered using measured geometry extents, physical
line width and XYToPage pose. The benchmark excludes two pixels along stroke
edges and eight pixels near joins. Each comparison covers 6,480 pixels. Both
live and portable SVG pass on all six routes: 96 comparisons across the captures.
Horizontal maximum channel error is seven, mean below 1.11; vertical maximum is
five, mean below 1.39. Existing maximum-seven and mean-below-one for opaque
two-stop paint, or 1.5 otherwise, bounds remain unchanged. Exact parity is open.

The initial geometry-only span reached maximum 20 and mean 2.55. Native pixels
show that stroke paint includes the physical half-width outside the geometry.
The core descriptor now expands orthogonal endpoints after page scaling, keeping
line width physical and preserving source stops. The shared fill painter and
sigma/gamma sampler consume that descriptor.

An owned native SVG probe confirms that a green layer with 40 percent transparency
overrides red-to-blue line stops with solid #00ff00 and opacity 0.6. Another shows
that LineColorTrans does not multiply gradient stop alpha. Probe directories are
visio-line-layer-30532dc9eb1342f4964c21108d649c16 and
visio-line-opacity-2c0c19f48c9947c89934b6048c73214f. Core tests preserve all four
native gradients in each capture through move/save/reparse. Regression coverage
also checks independent fill/line caches, scaling, layer overrides, detached
snapshots and combined stop budgets. Core reparse is not native Office acceptance.

Arrowheads retain solid fallback paint and report unsupported-gradient-arrows.
Nonlinear directions report unsupported-saved-line-gradient. Other oblique spans,
groups/flips, contours, exact colors, editing UI and native Office reopening
remain unverified. Accepted gradients retain unverified-gradient-raster.

## Draft safety and bounded formula follow-up

The viewer now pins published core
[`2343de2d82d863b7d9f765c07b8484b97a96917a`](https://github.com/ChristopherVR/ooxml/commit/2343de2d82d863b7d9f765c07b8484b97a96917a).
Its [exact-commit CI](https://github.com/ChristopherVR/ooxml/actions/runs/37091852205)
passed all six test shards, types, builds and package checks. This bounded formula
increment adds MODULUS, AND/OR/NOT, conservative scalar bit operations and
square-inch intermediate arithmetic. It adds real-corpus formula/cache evidence
without broadening package admission or claiming native Visio fidelity.

Editing controls preserve unapplied text across geometry edits and history on the
same source and target. Replacing the source resets drafts even when page and shape
IDs match. Refused operations retain inputs and display the core error code;
obsolete cancellation errors cannot replace current status. Rectangle creation
accepts optional plain text. Browser regressions download and reopen saved copies
and verify unknown package payloads, applied geometry and text.

The combined Windows aggregate passes both core type projects, 1,277 Visio tests
with the original corpus enabled (six optional skips), 402 viewer tests (one
optional skip), 74 binding tests, five SSR tests, 27 docs tests, converter checks,
production builds, workers and packed consumers. Installed Chrome passes all 38
supported development-route scenarios and all 36 production-compatible scenarios,
including four geometry/save workflows. Two source-import-only rendering tests
run on the development route. Independent integration review is clear.
Fresh public sibling clones also reproduce the exact pin, portable locks and
Apache metadata: core setup, root/binding locked installs, both core types,
1,265 core tests (18 optional corpus skips without the corpus environment),
viewer types/build and binding types pass.

This integration preserves the published UI and Apache metadata from
[`24d9ff8308230f06df7745bbf3ea13f9a1aae2bf`](https://github.com/ChristopherVR/visio-viewer/commit/24d9ff8308230f06df7745bbf3ea13f9a1aae2bf).
Master-linked moves remain a separate unpublished candidate. Native Microsoft
Visio reopen and fidelity remain unverified.

## Published active-master independence follow-up

The local sibling core now resolves untouched active master templates, instance
overrides and selected style ancestry in separate sheet scopes. This is a bounded
independence proof, including Control row aliases, field-free text inputs and
direct single-reference SETATREF. It does not write master caches or enable
editing master-linked shapes. Missing mappings, fields, style collisions, cycles,
unknown dependencies and construction limits refuse the transaction.

The original hash-pinned POI `test_text_extraction.vsdx` rectangle (page 0, shape 1)
now passes separate move, resize and delete saves, combined move/resize and caller
mutation isolation. Tests preserve sibling subtrees and every untouched package
payload, including all master definitions and relationships. Existing corpus
candidates passing all three commands increase from three to four; rectangle
creation increases from six to seven of 19 accepted drawings.

POI `test.vsdx` remains refused because of explicit container membership references.
`60973.vsdx` remains refused because other pages contain ambiguous IDs and
unsupported dynamic formulas. Neither case has a supported recalculation proof.
Native Microsoft Visio reopen and rendering fidelity remain unverified.

Local verification passed both core type projects and the full viewer `check`:
1,265 core Visio tests with the original corpus enabled (six optional skips),
394 viewer tests (one optional skip), 74 binding tests, five SSR tests, 27 docs
tests, builds, workers and packed-consumer checks. Independent review passed
79 cases across four suites. All 25 installed Chrome scenarios passed against
the final rebuilt bundle. Changed Visio files pass formatting.

Core follow-up is published at
[`5590c8af893e81b8a3d9f3f8874c9434821b6abd`](https://github.com/ChristopherVR/ooxml/commit/5590c8af893e81b8a3d9f3f8874c9434821b6abd).
At this checkpoint the viewer pinned this exact commit with a clean portable lock
and no source patch.
Integration preserved the published shared chart, text, SVG and Word changes.
Both core type projects, 1,319 combined Visio/shared-API tests, full build,
packed-consumer checks and 49 release-script tests passed on the integrated core.
Native Visio fidelity remains unverified.

## Safe geometry bundle and Windows recovery

Development checkouts are siblings under
`C:\Users\Christopher\Documents\Codex\2026-10-03\task`: `ooxml` and `visio-viewer`.
Core geometry logic is published at
[`dc60e01ff97ae4db5e346c9942cd4e9052ba7b9d`](https://github.com/ChristopherVR/ooxml/commit/dc60e01ff97ae4db5e346c9942cd4e9052ba7b9d),
preserving concurrent upstream changes and their automated release. The viewer
pins that actual commit with its exact npm lock and no temporary source patch.
Patch-free setup and build reproduce in `ooxml-portable-verified`. Earlier
reviewed snapshots remain in `ooxml-scope-verified` and `ooxml-verified`; the old
baseline remains in `ooxml-pinned`. No unrelated Office source was edited.

The published core's [exact-commit CI](https://github.com/ChristopherVR/ooxml/actions/runs/37085119760)
passed both type projects, all six complete test shards, package builds/imports
and the final required status. Local verification also passed full core builds,
all entry-point imports and release tooling after preserving the upstream release.
The pinned npm lock includes optional native packages for every platform. It is
generated from clean manifests without an installed dependency tree or seed lock;
setup regressions check complete optional dependency entries. Linux-targeted
`npm ci` validation passes as well as the installed Windows route.

The shared UI and every framework handle expose create rectangle, move, resize
and safe delete through typed atomic `applyEdits`, isolated worker execution and
bounded history. Existing admitted top-level local 2D shapes are supported; the
shape need not originate in this viewer. Coordinates are rotation-pin drawing
inches, bottom-left origin, up-positive. Resizing holds the pin fixed. Bounded
numeric ShapeSheet dependency closure and cache writeback preserve F/U and untouched
package payloads. The core README documents every exclusion and next expansion.

Final verification on Christopher-PC:

- Complete viewer `npm run check` passed: formatting, setup regressions, both
  core TypeScript projects, 1,220 active core Visio tests, converter verification,
  viewer types, 394 viewer tests, 74 native binding lifecycle tests, five SSR tests,
  27 documentation tests, production builds/workers and the packed external consumer.
- Seven optional core corpus cases and one optional viewer arrow case were skipped
  in that aggregate run. Separate `VISIO_EDIT_CORPUS_DIR=../corpus` core execution
  passed 1,221 tests with six unrelated optional corpus skips.
- Installed Chrome via `CHROMIUM_PATH` passed all 25 browser scenarios, including
  the new geometry controls, editing an existing imported local shape, undo,
  downloaded text reload, desktop/mobile themes and layout.
- Hash-pinned parser/scene/SVG corpus passed: 19 accepted drawings, ten expected
  malformed refusals, 22 SVG pages, 706 shapes and 627 paths.
- Independent review added nine passing safety regressions. Findings concerning
  glued endpoint caches, non-page and transitive metadata dependencies, units,
  implicit text dimensions and contradictory requested coordinates were fixed.
  Shared controller/worker/history/adapter review found no remaining material issue.
- The scope/protection follow-up passed 45 independent review regressions across
  three suites, including inherited cache/default and surviving-reference checks.

The subsequent scope correction verifies Microsoft's `F="No Formula"` marker,
separates master/page IDs, checks default/selected style ancestry and overrides,
records implicit theme selectors, and resolves inherited protection only through
verified inactive scalar ancestors. Unknown dynamic and affected unsupported
dependencies continue to refuse. Independent review covers inherited constants,
created-shape defaults, other-page metadata, used masters and unsafe locks.

The 29-file geometry corpus now passes with six successful rectangle creations
(bgcolor, blue-box, color-boxes, dwg, fdo86664 and qs-box). Of six first-page local
candidates, three each pass separate move, resize and delete saves:

| Drawing                       | Page / shape | Move / resize / delete                 |
| ----------------------------- | ------------ | -------------------------------------- |
| libvisio blue-box.vsdx        | 0 / 75       | All pass                               |
| libvisio color-boxes.vsdx     | 0 / 68       | All pass                               |
| libvisio qs-box.vsdx          | 0 / 75       | All pass                               |
| POI 60973.vsdx                | 0 / 11       | Unknown used-master dependency refusal |
| POI test.vsdx                 | 0 / 1        | Unknown used-master dependency refusal |
| POI test_text_extraction.vsdx | 0 / 1        | Unknown used-master dependency refusal |

Every successful save reopens and preserves all untouched package payloads;
refusals preserve the caller's original bytes. The 13 creation refusals remain
explicit. Thirteen admitted files have no candidate in the first-page scan:
bgcolor, dwg, fdo86664, office_varient4, tdf136564-WhiteTextBackground,
tdf154379-QuickStyleFillMatrix, testfile1, testfile3, testfile4, testfile5, testfile6,
60489 and github260. This does not assert absence on every page. Ten malformed
cases retain two MISSING_PART and eight INVALID_ZIP failures. No native Visio
reopen/fidelity evidence exists; standard executable locations and COM
registration did not detect Visio.

Windows npm child invocation now runs the JavaScript CLI through Node without
shell interpolation. Setup and literal-argument tests pass. This repairs the local
route; earlier cloud ENOSPC remains a separate environment failure. No PC files
were deleted to repair that cloud error, and no credentials or local Codex memory
were modified.

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

Run `npm run check` after both root/binding installations. Run `npm run test:browser` only in a supported browser environment. The browser scenarios remain launch-blocked locally; published baseline remote evidence is linked above. `scripts/render-corpus.mjs INPUT_DIRECTORY OUTPUT_DIRECTORY` produces secondary-renderer diagnostic artifacts from an explicitly supplied local corpus.

The Playwright config accepts `CHROMIUM_PATH`; install Playwright's Chromium or point it at a system Chromium. In one development environment the IPC policy blocked launching Chromium and the cloud browser blocked localhost, so browser screenshots and visual verification were not claimed there; source/static audits, DOM tests and production builds were performed instead.

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

## Native layer paint and classic linear fills (2026-10-07)

The shared core now applies a single colored layer to geometry and character
paint without rewriting the source color cells. Shapes in multiple layers retain
their own colors, matching the native cases. Solid fills become white; classic
orthogonal fill patterns 25-30 retain native two- or three-stop gradients. Modern
cached stop colors yield to the single-layer paint context. Character alpha uses
Visio's observed byte quantization, and shared SVG rendering and immutable export
snapshots retain that alpha.

`scripts/record-visio-layer-colors.ps1` creates an owned invisible Visio 16 instance
and records 21 native pages as VSDX, SVG and PNG. Compact reference measurements
are committed in `layer-colors-native.json`; native binary files remain local.
Cases cover no color, multiple/mixed layers, zero/partial/full and fractional
transparency, source alpha, no fill/line, hatch, modern gradients and patterns
25-30. Hatch rendering remains unsupported and diagnosed. Groups, foreign paint
and text-background overrides remain unverified and are diagnosed when applicable.

The native corpus passes core parsing, move/save/reparse and shared UI rendering
comparisons. Native reopening of the edited package remains unverified: the owned
COM instance stalled when opening both the edited file and Visio's own original.
The stalled test instances were closed; the existing user application was untouched.
This is not evidence of native save/reopen parity.

Local checks: 1,994 core Visio tests passed (41 optional skips with only this native
corpus enabled), followed by seven focused tests after adding three regressions;
710 shared UI tests passed (six optional skips). Strict core and viewer typechecks,
79 framework binding tests and five SSR tests passed. The broader UI declaration
build still reports the existing seven Teams/PPTX dependency/type errors.
All six Chromium framework routes passed the 21-case live/exported SVG native
stroke-pixel comparison (252 RGBA comparisons). All six page insertion, history
and saved-package browser regressions also passed. These are sampled stroke
comparisons, not complete image or text-layout parity.

## Native hatch tiles and transparency (2026-10-07)

The core now normalizes all built-in bitmap fill patterns 2-24 into bounded
8-by-8 PNG tiles. The native coverage table comes from
`scripts/record-visio-fill-patterns.ps1`, with compact pixel evidence in
`fill-patterns-native.json`. Separate red/blue opaque and mixed-color/translucent
captures cover 46 Visio 16 pages. The implementation reuses shared color parsing
and clamping and the existing ole2 PNG encoder. The shared UI reuses its raster
URL, embedded-resource, disposal, validation and immutable snapshot machinery.
No viewer binding contains pattern logic.

Every RGBA byte in both captures matches. Eighteen Chromium scenarios pass:
12 opaque/translucent hatch scenarios across all six frameworks and six existing
layer-stroke scenarios. Hatch tests compare 96-by-96 interiors of every native
page with both live SVG (inlining its existing blob resources for rasterization)
and portable exported SVG: 552 exact interior comparisons, with no channel
threshold or tolerance. Flipping the complete pattern instead of its child image
resolved the initial one-byte interpolation differences.

Core move/save/reparse tests retain every normalized tile. The full local Visio
suites pass 1,999 core tests (41 optional skips) and 711 UI tests (six optional
skips), strict core/viewer types, 79 framework binding tests and five SSR tests.
The existing seven Teams/PPTX declaration errors remain outside this change.
The earlier seven-package registry-only consumer verification also completed
successfully, including packed ESM/declarations and production worker checks.

Native rotated/scaled hatch placement and edit/save/reopen still need evidence.
These interior comparisons establish this captured subset, not complete Visio
visual, editing or printing parity. The earlier layer-paint record's hatch
limitation is superseded for patterns 2-24 by this follow-up.
A further optional native-corpus test moves and saves the genuine 23-page source,
reparses it and verifies identical normalized tile bytes for every page. All three
focused pattern tests passed with that native corpus enabled. The saved acceptance
candidate is `core-fill-patterns.vsdx` in the local oracle directory; opening it in
Microsoft Visio remains an outstanding native acceptance step.

## Page-aligned hatch orientation and phase (2026-10-07)

Native rotation references exposed an incorrect behavior: hatch axes rotated with
shapes. Rendering now composes shape/group transforms through the same affine
helper used by bounded foreign vectors, counter-transforms the hatch axes and
retains the native bottom-based page-height tile origin. Hatch geometry uses the
native crisp-edge rendering flag. The shared affine extraction is recorded in
PROVENANCE.md. Drawing-scale normalization retains physical tile dimensions.

The native capture script now accepts angle and drawing/page scale. Four captures
(opaque, translucent, 90-degree rotation and drawing scale 2:1) cover 92 native
pages. All six framework routes passed 24 scenarios and 1,104 exact live/export
96-by-96 interior comparisons. Six optional oblique scenarios were skipped in
that run; they were separately enabled for vanilla to investigate the next gap.

At 30 degrees, counter-rotation and tile phase removed the large color/placement
mismatch, and crisp-edge rendering removed the large boundary mismatch. Exact
comparison still fails: 2,916 channels across 46 renders differ by one byte
(maximum difference 1). The zero-tolerance test is retained under
VISIO_NATIVE_FILL_PATTERNS_OBLIQUE_DIR; it is not counted as passing. Core parsing
reports unverified-hatch-angle for oblique angles. The remaining cause and native
nested group/flip behavior need stronger evidence before claiming pixel parity.

The latest native reopen probe initialized an owned invisible application with a
blank document first, then used normal events and OpenEx(128). It still stalled on
the native original, so opening the core-edited file was not reached. The probe's
own Windows Visio and PowerShell processes were terminated; the existing user
Visio instance was left running. Native acceptance remains unverified.

Before the final two focused regressions, local checks passed 2,000 core/shared
geometry tests (43 optional skips), 710 UI tests (seven optional skips) and strict
core types. The shared geometry ESM export smoke also passed. Existing seven
Teams/PPTX declaration errors remain outside this change.
Final core/shared geometry verification passed 2,004 tests with the native layer
and pattern source environments enabled (41 optional skips). Root lint passed
with the existing nine unrelated warnings. Viewer types are rechecked after
rebasing the separately published browser-test type-import fix.
Viewer typecheck passed after rebasing the public-package browser-test type import
fix. Binding verification passed 79 client tests, five SSR tests and Svelte checks.

## Native hatch reflections (2026-10-07)

The existing shared affine and hatch counter-transform logic was checked against
46 additional pages exported by Microsoft Visio 16: all 23 hatch patterns with
horizontal flips, then all 23 with vertical flips. The capture script now accepts
FlipX and FlipY and records both in evidence.json. Native sources were created in
an owned invisible Visio application and saved as VSDX, SVG and PNG references.

The horizontal capture is visio-fill-patterns-d8955fdfd1a8466ba78908011765ef96
in the local temporary directory; the vertical capture is
visio-fill-patterns-cc4ee2b703af44689ccc13d706266afb. Set
VISIO_NATIVE_FILL_PATTERNS_FLIP_X_DIR and VISIO_NATIVE_FILL_PATTERNS_FLIP_Y_DIR
to reproduce the optional fill-pattern.spec.ts comparisons.

All six framework routes passed for both captures: 12 scenarios, 552 exact
96-by-96 interior comparisons across live and portable exported SVG, including
all RGBA channels. Viewer typecheck passed. No new rendering implementation was
needed. Difference attachments now run before the aggregate assertion, so a
failing optional native comparison retains its per-pattern evidence.

This establishes only single-shape, axis-aligned reflection interiors. Nested
groups, reflected oblique shapes, full-page boundary fidelity and native reopening
of core-edited hatch documents remain unverified. The existing 30-degree native
comparison still fails by one byte and its zero-tolerance gate is unchanged.

## Native nested hatch groups (2026-10-07)

The native capture script can now build up to eight real group levels through
Page.CreateSelection and Selection.Group. Each level has an invisible sibling;
GroupAngle and GroupFlipX/GroupFlipY apply to the native parent. The evidence file
records settings and actual group IDs. The optional browser comparison now also
asserts that every page preserves the expected two-level hierarchy after worker
parsing, rather than only checking pixels.

Two captures used two 90-degree parent rotations, with and without a horizontal
reflection at each level. These cover 46 native pages in
visio-fill-patterns-4e96a0ec7e7f4cadb69f40361ebe0a32 and
visio-fill-patterns-412c142e096146f3ac419e78094ae7fa in the local temporary
directory. All six framework routes passed 12 scenarios and 552 exact live/export
96-by-96 interior comparisons. Existing shared transform composition handles
these cases without another implementation. Viewer typecheck passed.

The capture parameters follow Microsoft's Page.CreateSelection and
Selection.Group APIs:
https://learn.microsoft.com/en-us/office/vba/api/visio.page.createselection
https://learn.microsoft.com/en-us/office/vba/api/visio.selection.group

These references establish bounded quarter-turn group interiors only. Arbitrary
group rotations, parent resizing, clipping, full-page boundaries, grouped editing
and native reopening remain unverified. Native pixel equality at 30 degrees
continues to fail by one byte; its zero-tolerance test remains unchanged.

Two further captures rotated the child by 90 degrees as well, giving a total
270-degree orientation for the unreflected case and a quarter-turn child inside
the reflected hierarchy. These are
visio-fill-patterns-4076f84791094f11b7e2143ae131337d and
visio-fill-patterns-cc5b82f96b8c4e8485037eaca0b8aaba. Both passed all six
framework routes with hierarchy assertions and zero differing channels.
Across all four grouped captures, 92 native pages passed 24 scenarios and 1,104
exact live/export interior comparisons. Use VISIO_NATIVE_FILL_PATTERNS_GROUPED_DIR
and VISIO_NATIVE_FILL_PATTERNS_GROUP_FLIPPED_DIR for either corresponding pair.

## Native classic radial fills (2026-10-07)

Classic fill patterns 36-40 now normalize to radial gradients with native
object-bounding-box centers and radii (1.4 for the four corners, 0.73 for the
center). Centers use local y-up coordinates. Stops reuse the existing color,
transparency, validation and SVG stop paths. Linear gradient behavior is retained;
normalized radial geometry does not change with drawing/page scale. Live SVG,
portable SVG and immutable print snapshots share the same paint implementation.

The native capture script accepts FirstPattern and LastPattern while keeping the
original hatch defaults. Non-raster native patterns are retained as SVG/PNG/VSDX
references without assuming their pattern definitions contain a bitmap. The first
31-40 capture showed that patterns 31-35 use multiple linear-gradient triangles;
those remain explicitly unsupported.

Microsoft Visio 16 opaque and translucent 36-40 captures are respectively
visio-fill-patterns-0b9cfd92659d4f5dbddd2f16b5f8a97a and
visio-fill-patterns-2ee45429452b400f8e9ae224cffc96cd in the local temporary
directory. The alpha capture uses RGB(27,139,211) at 20 percent transparency and
RGB(231,61,83) at 50 percent transparency. Enable VISIO_NATIVE_RADIAL_FILLS_DIR
and VISIO_NATIVE_RADIAL_FILLS_ALPHA_DIR to reproduce browser comparisons.

All six framework routes passed both samples: 12 scenarios, 120 full-page
576-by-432 RGBA comparisons across live and portable SVG. Every channel matched
the native exported SVG raster exactly, with zero tolerance. This evidence is
native SVG browser rasterization, not a claim that browser output equals the
separately exported native PNG pipeline. Group rotation, nonrectangular geometry,
colored layers, modern radial ShapeSheet settings and native reopening of
core-edited radial documents still need authoritative acceptance evidence.

Core regressions cover all five centers, separate stop alpha, physical page-scale
normalization, move/save/reparse preservation and independent snapshot copies.
Scene validation rejects invalid radial radii; DOM tests check shared live/export
and immutable print output. Strict core types and viewer types passed. The Visio
core suite passed 2,029 tests with 43 optional skips. The UI suite initially had
one new test failure from omitting the print snapshot's required pageIndices;
that test was corrected and both focused paint tests passed. Seven existing
Teams/PPTX declaration errors remain outside this change. Viewer formatting also
identified the pre-existing parity.md table formatting issue; another session is
handling that isolated formatting fix.

The corrected complete UI suite passed 711 tests with seven optional skips. The
optional native radial regression passed for both captures, producing
core-fill-patterns.vsdx acceptance candidates and preserving all normalized paint
through move/save/reparse. That is parser/serializer preservation evidence; it
does not establish native Visio reopen acceptance. Root lint passed with the nine
existing unrelated warnings.

## Native classic region gradients (2026-10-07)

Classic fill patterns 31-35 now preserve native triangle regions and their
triangle-local linear gradients. The core normalizes two triangles for patterns
31-34 and four for pattern 35, including region-local endpoints, native rotation,
foreground/background colors and independent stop alpha. Existing linear endpoint
and color/opacity helpers are reused. The renderer shares its linear/radial stop
serialization; no framework-specific painter was added.

Native SVG uses top-down pattern coordinates. Flipping the pattern transform,
rather than the triangle content, is required by these captured pixel references.
The latter differed by one byte for opaque patterns and up to two for alpha;
the final transform matches both references exactly without a tolerance change.
Snapshots copy the region vertices and endpoints independently. Existing scene
validation moved into a reusable gradient helper, with the extraction recorded in
PROVENANCE.md. Export estimates and print work budgets count all rendered regions.

Microsoft Visio 16 captures in the local temporary directory are
visio-fill-patterns-6f5fdd0b579c4171ae17dd584c3ca7ef (opaque) and
visio-fill-patterns-74695b6af9d84dd59f80a4415f6ea7dc (alpha). The alpha values
are RGB(27,139,211) at 20 percent transparency and RGB(231,61,83) at 50 percent.
Enable VISIO_NATIVE_REGION_FILLS_DIR and VISIO_NATIVE_REGION_FILLS_ALPHA_DIR
for the optional browser and native-source preservation tests.

All six framework routes passed both captures: 12 scenarios and 120 exact
576-by-432 full-page RGBA live/export comparisons against native SVG rasterization.
The genuine native documents also passed core move/save/reparse preservation;
core-fill-patterns.vsdx candidates were written alongside their native sources.
This is not native Visio reopen acceptance, nor native PNG pipeline equivalence.
Grouped, rotated, nonrectangular and colored-layer region fills still need native
comparisons. Modern gradient settings and broader effects remain separate gaps.

Strict core types, viewer types and viewer formatting passed. Visio core passed
2,031 tests with 44 optional skips; UI passed 712 tests with seven optional skips.
Root lint passed with nine existing unrelated warnings. The snapshot regression
now narrows the gradient discriminant before checking linear endpoints, fixing
the test type errors introduced when radial paints were added. Its four focused tests passed. After generating the missing local PowerPoint
declarations, the full UI typecheck passed for both the strict and relaxed PPTX
projects. The seven missing-declaration errors were local build prerequisites.

## Native saved vertical gradients and inherited placeholders (2026-10-07)

Saved ShapeSheet linear gradients now accept all orthogonal cached angles,
including saved decimal rounding. The existing shared linear endpoint helper
normalizes 90 and 270 degrees in local y-up inches. Oblique saved angles, group
gradients and gradients that do not rotate with their shape remain diagnosed.

Native sources exposed an additional preservation/display issue: two explicit
stop rows inherit eight wholly themed placeholder rows from the Theme style.
Those unused tail rows are now excluded from active stop normalization. Empty
active rows and partly themed active rows remain rejected, as do malformed
colors, positions and transparency. Original XML rows remain unchanged on save.

The native setup in the layer capture was extracted into a shared PowerShell
helper; both capture scripts use it. A fresh 21-case layer run, in
visio-layer-colors-8f6e96259c324d1ebda122de6e084662 in the local temporary
directory, reproduced every original native paint record unchanged.
The extraction is recorded in PROVENANCE.md.

New modern gradient captures use legacy patterns 25-30 underneath explicit saved
stops, proving that the modern gradient overrides all six legacy codes. They are
visio-fill-patterns-8c74ec6e599846029ce5130cc14a194e (90 degrees, opaque) and
visio-fill-patterns-7f4b843fd2494e19b82707140ba063c7 (270 degrees, alpha) in
the local temporary directory. Enable VISIO_NATIVE_LINEAR_VERTICAL_DIR and
VISIO_NATIVE_LINEAR_REVERSE_DIR for native browser and preservation comparisons.
The core move/save/reparse test accepted both genuine sources and wrote
core-fill-patterns.vsdx candidates without changing normalized gradient paint.

A separate owned native reopen probe used OpenEx(456): hidden window, macros
disabled, no recent-document entry and no workspace restoration. It still stalled
on the native original, before reaching the core-edited file. Its owned Windows
Visio process 28404 and PowerShell process 38376 were verified and terminated;
the user's Visio process 43312 was left running. Native reopen acceptance remains
unverified. Flag meanings follow Microsoft's Documents.OpenEx documentation:
https://learn.microsoft.com/en-us/office/vba/api/visio.documents.openex

All six framework routes passed both saved-gradient captures: 12 scenarios and
144 exact full-page 576-by-432 RGBA live/export comparisons against native SVG
rasterization. Core regressions cover wrapped and rounded quarter-turn angles,
inherited wholly themed tails and malformed empty/partly themed active rows.
The Visio core suite passed 2,048 tests with 45 optional skips; UI passed 712 tests
with seven optional skips. Strict core types, the complete UI typecheck (strict
and relaxed PPTX projects), viewer types, viewer formatting and root lint passed.
Root lint retains its nine existing unrelated warnings. Native reopening and
native PNG pipeline equivalence remain unverified.

## Native saved oblique gradients (2026-10-07)

Native saved linear gradients rotate within the normalized shape bounding box.
Physical endpoint projection changes the angle on nonsquare shapes. The core
linear paint model now retains an optional bounding-box angle for oblique saved
caches; quarter turns retain their existing physical endpoints. The shared SVG
renderer reuses its gradient rotation and stop serializer. Drawing-scale
normalization leaves normalized endpoints intact, snapshot copying retains the
angle, and scene validation bounds the angle and endpoints.

Owned Visio 16 captures contain six pages each with explicit saved gradients
above legacy patterns 25-30: visio-fill-patterns-fa9ce73adad349318b980f6740ab70db
(30 degrees, opaque) and visio-fill-patterns-800f7e4db3694376914dffb3f645c0c5
(225 degrees, translucent), in the local temporary directory. Enable
VISIO_NATIVE_LINEAR_OBLIQUE_DIR and VISIO_NATIVE_LINEAR_OBLIQUE_ALPHA_DIR for
native comparisons. Both genuine source packages passed core move/save/reparse
preservation checks. Native reopening remains unverified.

All six routes passed both captures: 12 scenarios, 144 exact full-page
576-by-432 RGBA live/export comparisons against native SVG rasterization, with
zero differing channels. Core passed 2,052 tests with 45 optional skips; UI
passed 713 with seven optional skips. Strict core and complete UI types passed;
root lint retains nine existing unrelated warnings. Regression coverage includes
negative and wrapped oblique angles, decimal rounding, scaling, snapshot copying
and invalid normalized paint. Broader transformed gradients, modern radial
fills, effects and native PNG pipeline equivalence remain open parity work.

## Native saved radial gradient presets (2026-10-07)

Complete saved shape-rotating radial gradients now use the same normalized paint
model and SVG renderer as classic radial fills. Directions 1-7 retain the seven
native centers/radii; only the linear direction consumes FillGradientAngle.
Unsupported rectangular/path directions and group/nonrotating gradients remain
diagnosed. Classic fills now use the same core geometry helper, with unchanged
paint. Source imports remain extensionless.

Owned Visio 16 captures in the local temporary directory are
visio-fill-patterns-2a5f14c7ad994952a60b88c5760a09f7 (opaque) and
visio-fill-patterns-5410848eb1a0466a9f560793b35ec4dc (translucent), containing
all seven presets. Enable VISIO_NATIVE_SAVED_RADIAL_DIR and
VISIO_NATIVE_SAVED_RADIAL_ALPHA_DIR for the browser and preservation tests.
Both genuine packages passed move/save/reparse checks. All six framework routes
passed both captures: 12 scenarios and 168 exact full-page 576-by-432 RGBA
live/export comparisons against native SVG rasterization, with zero differing
channels. This does not establish equivalence to native PNG export.

The full Visio core suite passed 2,062 tests with 46 optional skips before the
native test bodies were consolidated. The UI suite passed 713 with seven skips;
strict core and complete UI types passed. After consolidation, the saved and
classic gradient suites passed 57 tests with six optional skips, including both
genuine radial captures. Native reopening remains a separate acceptance gate.

A fresh owned instance authored all seven translucent presets, saved and closed
its document, produced a core-edited candidate, then called OpenEx(200) on the
native original in that same instance. It stalled before reaching the candidate.
Its owned Visio process 33288 and PowerShell process 41760 (both started at
22:41:02) were verified and terminated. The user's process 43312 remained running.
Read-only enumeration of the owned process's top-level windows found PDFMakerWindow
alongside the Visio main window, with no separate file/repair dialog. This does
not establish the cause of the stall; no add-in or user settings were changed.

## Native saved rectangular gradient presets (2026-10-07)

Saved directions 8-12 now reuse the classic triangle-region gradient geometry,
including native 360-degree rotation spelling. The existing SVG renderer,
weighted gradient validation, snapshot copying and export/print budgets handle
the paint without a second implementation. Source imports remain extensionless.
Direction 13 (path-following), group gradients and gradients that do not rotate
with their shape remain diagnosed. A native rectangular direction-13 probe
resembled the centered rectangular preset, which does not establish arbitrary
path-gradient behavior and was not used to broaden support.

Owned Visio 16 captures contain all five presets in
visio-fill-patterns-9fb7f01f032b450a990f8155a45ac8f1 (opaque) and
visio-fill-patterns-4d6d37ce83af41498a5022aa7acf9ca6 (translucent), in the local
temporary directory. Enable VISIO_NATIVE_SAVED_REGIONS_DIR and
VISIO_NATIVE_SAVED_REGIONS_ALPHA_DIR. Both genuine packages passed core
move/save/reparse preservation checks. All six routes passed both captures:
12 scenarios and 120 exact full-page 576-by-432 RGBA live/export comparisons
against native SVG rasterization, with zero differing channels.

Core passed 2,067 tests with 51 optional skips; UI passed 713 with seven skips.
Strict core, complete UI types and root lint passed, retaining nine unrelated
lint warnings. Native capture PowerShell scripts passed syntax parsing. Native
reopening, broader transformed gradients, arbitrary path fills and native PNG
pipeline equivalence remain unverified.

## Native gradient raster evidence (2026-10-07)

Native SVG export does not establish native gradient color fidelity. A new
`scripts/record-visio-gradient-raster.ps1` probe captures both native SVG and PNG,
the genuine VSDX, application version and seven RGBA samples for each case.
Its 68 cases cover directions 0-12 on rectangles and direction 13 on rectangles,
ellipses, triangles and notched polygons, with two or three stops and opaque or
translucent colors. Shape exports are 288 by 144 pixels at 144 dpi. The owned
Visio instance restores its previous raster export settings before closing.

The completed local capture is
visio-gradient-raster-032b1a2ac8e44d0582f3838badac0448 in the temporary directory.
For an opaque red-to-blue horizontal two-stop gradient, native PNG samples at
x=36, 72, 144 and 216 (y=72) are respectively (250,0,63), (238,0,105),
(186,0,186) and (104,0,238). These differ substantially from the SVG's straight
interpolation. A GDI+ gamma-corrected sigma blend reproduces the first three
samples exactly and differs by one channel in the fourth; this is evidence for
further investigation, not a completed implementation. Three-stop native
red/green/blue samples instead follow ordinary straight color interpolation.

Native direction-13 SVG uses a rectangular fallback even for ellipses and
polygons; PNG follows a different contour. Direction 13 remains unsupported.
Two-stop normalized fills and classic gradients now report
unverified-gradient-raster. The earlier exact native-SVG comparisons remain
valid for exported SVG geometry and stop serialization, but do not prove native
PNG color or contour parity.

The polygon raster captures use native DrawPolyline. Earlier probes using
manually replaced geometry rows produced blank polygon PNGs and were discarded
as raster evidence. The shared explicit-row helper still serves the rounding
probe; all nine native rounding cases passed after extraction. A separate
read-only owned-instance check confirmed raster settings were restored to
source size and screen resolution. No user-owned Visio instance was changed.

Core passed 2,065 tests with 53 optional skips. The UI run passed 711 tests;
after updating the two saved-gradient assertions to expect the fidelity warning,
both passed, giving 713 passing tests and seven optional skips. Strict core types,
PowerShell syntax parsing, viewer formatting and root lint passed; lint retains
the same nine unrelated warnings. This checkpoint records a remaining gap and
does not change gradient interpolation or establish full Visio parity.

## Native opaque gradient interpolation (2026-10-07)

Saved opaque two-stop gradients with boundary stops now retain an explicit
sigma-gamma22 paint profile. The core sampler expands it into 256 color stops
using native GDI+ quantized sigma factors and gamma 2.2, reusing shared color
hex conversion. The ShapeSheet source rows remain two stops. Translucent,
non-boundary, theme and classic gradient interpolation retain the diagnostic;
the diagnostic also remains on the improved opaque profile because exact native
raster equality is not established. Three-stop fills do not use this curve.

Saved radial centers reuse the existing shared preset helper, while radii now
use the physical distance to the farthest corner. Native PNG uses a circle in
local physical coordinates, unlike the bounding-box ellipse exported by native
SVG. Scaling, scene snapshots, runtime validation, SVG/print byte budgets and
aggregate rendered-stop limits retain this distinction. Opaque region paint
also suppresses transparent seams on shared internal triangle edges.

The six-route gradient-raster.spec.ts benchmark imports the genuine 68-shape
package through the worker and measures directions 0-12 in live and portable
SVG paint: 156 native PNG comparisons. It compares the 272-by-128-pixel interior
of each 288-by-144 native shape export. All six routes passed the recorded
regression bound of seven channel levels and mean difference below one level.
These bounds measure improvement; they are not an exact-parity tolerance.
Outer-edge antialiasing, small phase/quantization differences and the centered
rectangular preset's residual difference remain open. The previous saved-gradient
native-SVG equality scenarios were retired because SVG is not a native paint
oracle; classic/hatch SVG compatibility checks remain explicit.

The mixed native package initially refused all edits because POLYLINE was
classified as a dynamic dependency. Its geometry-data function has explicit
arguments, so the existing dependency analyzer now records those arguments as
static. Recalculation remains unsupported and rejects affected expressions;
nested GETREF remains dynamic. All 68 native shapes then passed core
move/save/reparse gradient preservation. This does not establish native reopen
acceptance or arbitrary polyline editing.

Core passed 2,071 tests with 53 optional skips, including the native mixed
package. UI passed 713 tests with seven skips before the additional expanded-paint
preflight test; all 45 exporter tests then passed, giving 714 passing UI tests.
Strict core, complete UI and viewer types passed, as did the full UI build,
fresh Visio ESM/CJS bundles, PowerShell syntax and formatting. A Node public-ESM
smoke check edited/reparsed all 68 native shapes, retained 13 opaque profiles and
changed only visio/pages/page1.xml. Root lint retains nine unrelated warnings.

## Translucent and three-stop region seams (2026-10-08)

Every triangular region now uses crisp shared internal edges, including
translucent and three-stop fills. The enclosing shape still supplies its normal
outer contour. This reuses the existing renderer, stop serializer, color model
and snapshot/export pipeline; no second interpolation engine was introduced.
An experimental dense RGB/alpha ramp did not improve translucent linear paint
and was not adopted.

The native PNG benchmark now covers all 52 rectangle cases from the existing
68-case capture: directions 0-12, two/three stops and opaque/translucent colors.
All six framework routes passed both live and portable export comparisons:
624 comparisons of 272-by-128-pixel interiors within the 288-by-144 PNG exports.
The benchmark writes native-raster-differences.json per route, including each
case's name, stop count, alpha flag, maximum channel error and mean error.

Native export bounds include a saved line-width margin even when LinePattern
is zero. The benchmark now reads the native SVG viewBox only to register that
frame against the core shape placement. All expected colors come from the
unchanged native PNG. This corrected the centered opaque three-stop case's
maximum error from eight levels in the unregistered comparison to seven in the
registered core comparison. The original opaque two-stop maximum bound remains
seven, with mean below one; newly measured translucent and three-stop cases have
mean bounds below 1.5. These are regression bounds, not parity tolerances.

Maximum/mean errors observed on every route are: opaque two-stop 7/0.65,
opaque three-stop 7/1.36, translucent two-stop 6/1.26 and translucent three-stop
7/1.25 (means rounded up to two decimals). All normalized gradients now report
unverified-gradient-raster, since the three-stop corpus also demonstrates
nonzero native raster differences. Path-following fills, outer contour pixels,
broader colors/positions/transforms and native reopen acceptance remain open.
Classic rectangular native-SVG equality scenarios were retired: their exported
transparent seams are not a reference for the native PNG renderer. Classic
radial and hatch SVG compatibility checks remain separately identified.

Core passed 2,071 tests with 53 optional skips; UI passed 714 with seven skips.
Strict core, complete UI and viewer types, the UI build and viewer formatting
passed. Root lint retains nine unrelated warnings. This checkpoint fixes shared
region seams and strengthens native evidence; exact pixel parity is unproven.

## Native path-fill interiors (2026-10-08)

Saved FillGradientDir 13 now reuses the already evaluated single filled outline.
Straight closed bounds-filling polygons use the existing region gradient model
with normalized shape coordinates, bounding-box center and perpendicular
projections onto their outline edges. Shared flattening, winding, intersection
and projection helpers supply the geometry. Reentrant or degenerate faces are
skipped. Canonical axis-aligned bounds-filling ellipses retain analytic radial
paint with normalized radius 0.5. Curves beyond that ellipse, open/multiple
outlines, holes, self intersections and offset contours remain diagnosed rather
than inferred from Visio's rectangular SVG fallback.

The existing 68-case Visio 16 PNG capture now feeds every case into the shared
live and portable SVG renderers on all six framework routes: 816 comparisons.
The original 52-case rectangle benchmark and its bounds are retained. The 16
path cases exclude contour pixels: ellipse samples use the inner 80% radius;
polygon samples lie inside their captured outline and more than eight pixels
from every edge. The notched capture comes from native DrawPolyline, whose PNG
contains the actual fill. Native PNG colors remain unchanged; SVG supplies only
the export frame. Every output records its measured pixel count and errors.

All six routes passed. Maximum path-fill channel errors were 10, with worst mean
2.35 in the opaque three-stop triangle (rounded up). Regression bounds are 10
maximum and mean below 2.5; these are not parity tolerances. Native SVG fallback
experiments had much larger contour errors, but that exporter is not the native
paint oracle. Exact colors, contour antialiasing, arbitrary curves/holes,
broader concave outlines, rotations/groups and native save/reopen acceptance
remain open. All normalized gradients retain unverified-gradient-raster.

Core passed 2,076 tests with 53 optional skips, including genuine 68-case
move/save/reparse preservation and 17 opaque two-stop interpolation profiles.
UI passed 714 with seven skips. Core/UI/viewer types and UI builds passed.
The root lint run found an existing unused index parameter in XLSX drawings;
this Visio change does not modify that file. Path fan snapshot isolation,
normalized page scaling and generated-stop budget accounting have regressions.

## Extended native path corpus and unresolved star fidelity (2026-10-08)

The shared native shape helper now supplies exact polygon vertices to both
DrawPolyline and the capture metadata. record-visio-gradient-raster.ps1 accepts
PathShapes; its default 68-case capture is retained. The extended capture uses
rectangle, ellipse, triangle, notched, pentagon, chevron, ushape and star,
producing 84 cases in visio-gradient-raster-extended-e3f56a8aba4b4ec5a1feea169af04afe
in the local temporary directory. Native settings were restored and the owned
InvisibleApp quit; the user's pre-existing Visio process was untouched.

All six framework routes produced live and portable exports for every case:
1,008 interior comparisons. Exact capture vertices define the polygon masks,
with the same eight-pixel contour exclusion; ellipse masks are unchanged.
The 80 non-star cases meet the existing path bounds of maximum 10 and mean below
2.5, while directions 0-12 retain their tighter bounds. Every route's star
fidelity scenario fails the same unchanged maximum-10 gate. Playwright marks
these six scenarios as expected failures only after successful import, render,
artifact writing, result-count and minimum-pixel checks. Its summary counts
expected failures as passing tests; that does not mean native fidelity passed.

Star errors reach 43 levels with mean 3.72 (rounded up), in the opaque
three-stop live output. The exported counterpart reaches 42 and mean 3.53.
Opaque two-stop, translucent two-stop and translucent three-stop stars also
exceed the bound. No benchmark tolerance or production paint changed to hide
these failures. The earlier 68-case result remains valid at its original scope.

A separate diagnostic rasterized the native SVG geometry with the hypothesized
fan paint and swept center positions and pixel translations against unchanged
native PNGs. The best tested center remained the bounding-box center; a
half-pixel translation in each axis reduced the star's opaque three-stop
maximum to 12 and mean to 0.76. Triangle/pentagon errors also fell. Results are
in phase-probe.json beside the extended capture. This suggests sampling phase
contributes to the mismatch; it does not prove the native UI pipeline or justify
moving reference pixels or adding a fixed-resolution offset to production SVG.
Rotated/grouped comparisons and exact contour pixels still require evidence.

Both genuine native documents passed move/save/reparse preservation, including
all 84 extended shapes and 21 opaque two-stop profiles. The shared optional
regression covers both captures without copying the edit/preservation logic.
Core and viewer type checks passed. This checkpoint broadens the native oracle
and records a reproducible unresolved failure; it does not establish parity.

## Native rotated path fills and independent PNG frames (2026-10-08)

record-visio-gradient-raster.ps1 now accepts ShapeAngle and a direction range.
The final 30-degree capture is visio-gradient-raster-rotated-b043733eb59041948bd3f2e97c3a7293
in the local temporary directory: direction 13, eight outlines, two/three stops,
opaque/translucent paint, for 32 cases. The owned InvisibleApp restored raster
settings and quit. Additional setup captures were diagnostic only; empty or
misregistered frames were not used as fidelity evidence.

The recorder measures native page-space geometry extents with BoundingBox
(visBBoxDrawingCoords | visBBoxExtents), the saved physical line width, and the
local-to-page pose with three XYToPage calls. The browser uses DOMMatrix for the
same coordinate conversion and inverse ellipse masks. It checks all six parsed
pose coefficients against the independent native measurements, with error below
1e-9, before applying any expected-failure status. No transform algorithm or
production paint was copied or changed.

This experiment exposed an exporter difference: native rotated SVG frames
include more of the shape box than native PNG exports. PNGs fit the actual
geometry extents uniformly into the requested 288-by-144 canvas. The benchmark
now registers these captures using the measured geometry extents plus the saved
line-width margin, including hidden-line caches, and centers any unused canvas
space. PNG reference colors and pixels remain unchanged. Older captures without
these measurements retain their historical SVG-frame registration; their
small nonzero differences are not proof of exact pixel correspondence.

All six routes produced live and portable renders: 384 interior comparisons.
The 25 baseline cases meet the existing maximum-10/mean-below-2.5 path bounds.
Four star cases and three opaque polygon cases exceed the unchanged maximum-10
gate. Each route has one passing scenario and two explicit expected fidelity
failures, separated after successful import/render/artifact/count/minimum-pixel
and native-pose checks. Playwright reports 18 passing tests because it includes
expected failures; that is six passing scenarios and twelve expected failures,
not 18 native-fidelity successes.

Worst observed rotated errors are star opaque three-stop 35/2.44,
pentagon opaque two-stop 15/0.92, and triangle opaque three-stop 13/2.45
(maximum/mean, means rounded up). The pentagon opaque three-stop live output
also reaches 11. The passing group reaches maximum nine. Contour pixels remain
excluded by the eight-pixel polygon margin or inner-80% ellipse mask. Native
screens, other angles, flips/groups, exact colors and native reopen acceptance
remain open. These measurements verify one rotation scope, not broad parity.

Core optional regressions preserve all 32 rotated gradient models through
move/save/reparse, with eight opaque two-stop profiles. The same parameterized
regression also passed the original 68- and extended 84-case captures. Core
strict types passed. The source warning for unverified gradient raster remains.

## Native straight-line creation

The fresh `visio-line-movement-585c12348720422c8b81a92b6178b192` capture
uses the recorder's `-GridAligned` option and a newly owned invisible Visio
instance. Four native DrawLine cases cover horizontal, 45-degree diagonal,
vertical and reversed endpoints. The recorder saves original/moved SVG and
VSDX references and captures before/after cells and transforms.

Core clears the native lines, recreates their IDs and endpoints, and compares
geometry and effective styles exactly, with caches and transforms to twelve
decimal places. It also verifies creation followed by endpoint editing,
deletion, round-trip preservation, invalid/colliding input, detached command
snapshots, page scaling and ten-digit descendant ID allocation/exhaustion.
The focused suite passes eight tests, including the optional native check.

`native-line-create.spec.ts` passes all six framework routes: 24 real drags,
preview paint, native poses, undo/redo per line, public export, saved reload
and cancelled-drag source preservation. Saved endpoints match independently
measured native SVG endpoints to three decimal places. Horizontal SVG lines
have zero-height bounding boxes: checks inspect stroke visibility and actual
path length instead of using rectangle visibility as a paint proxy.

The earlier core Visio suite passed 2,193 checks with 64 optional skips; the
additional ID-allocation regression subsequently passed in the focused suite.
Core, shared UI and viewer typechecks pass. The existing rectangle toolbar
workflow passes after sharing its drawing lifecycle. Three vanilla native
endpoint/cancellation browser regressions also pass. Packed ESM, declarations,
consumer build and production parser/edit-worker smoke checks pass.

These checks prove the documented basic unglued-line scope. Native UI creation
gestures, native snapping and joining/glue, other default style documents,
scaled browser creation, exact paint pixels and Office reopen acceptance are
not proven. Native COM DrawLine is the geometry/style reference, not evidence
that our pointer behavior matches every native gesture.

## Scaled straight-line creation and paper-edge regression

Fresh GridAligned native DrawLine references use DrawingScale/PageScale
2/1 (`visio-line-movement-b6aea1eb0dd544f2ba23d485633a53eb`),
1/2 (`visio-line-movement-c3f006d574e340a0a59b71f48d2f8a0c`) and
1/3 (`visio-line-movement-0cd4f4e785474c309f0cdbb1bef1eb1c`).
The recorder converts its grid coordinates and translations into drawing
inches so all captures have the same physical line endpoints. Paper
dimensions retain native scale behavior and are independently checked
against exported SVG inch dimensions. Each owned invisible instance quit;
the existing user Visio instance was untouched.

The first comparison exposed a paper-edge clamp: the half-scale diagonal
was recreated at a different angle because its physical EndX exceeded the
paper width. Line gestures now use the shared page-point conversion without
that clamp, retaining native endpoint coordinates beyond the paper edge.
The native half-scale source retains its original dimensions and out-of-page
lines; the regression does not enlarge the page to avoid the mismatch.

Core has eleven passing checks with all four references enabled. Native
caches and COM transforms agree to twelve decimal places after scaling
translation components. Geometry and effective styles match the parsed
native references, and every part outside the edited page contents remains
byte-identical, including native page-scale metadata.

The browser suite covers 24 workflows and 96 actual line drags across
six frameworks and four ratios, with previews, per-line undo/redo, public
save, reload and cancellation. Native physical SVG endpoint comparisons
retain their three-decimal contract. Cancellation starts at an on-screen
line endpoint after fitting the page; arbitrary client locations on large
scaled pages are not assumed to be visible. The rectangle toolbar regression
also passes. Core/shared UI builds and all compiler projects pass; the
shared UI command/scaling suite passes thirteen checks with one optional skip.

Other scales and unit conventions, native mouse gesture semantics beyond
paper edges, outside-page paint visibility, native snapping/glue, exact
pixels and native Office reopen acceptance remain unverified.

## Rectangle document drawing defaults

Native Visio 16 captures `visio-line-movement-3da75dd25cc14af3bf3643997071521c`
and `visio-line-movement-08e67d1ab8e24e408e7ece914e731f03` use the existing
recorder's `-GridAligned -IncludeRectangle` options, with `-CustomDefaults`
for the latter. Three distinct native styles provide line RGB(24,96,168),
line weight 0.05 inch, fill RGB(240,176,80), and a separate text style.
The document's defaults use these IDs rather than DocumentSheet fallback IDs.
[Microsoft documents](https://learn.microsoft.com/en-us/office/vba/api/visio.document.defaultstyle)
that drawing methods and drawing tools use the document drawing defaults.

Before the fix, the native custom rectangle comparison failed with white
fill, black line and 0.01041666666666667-inch weight. Rectangle creation and
detached scope admission now share the existing line default-style helper.
The native line has local QuickStyle matrices of one; the native rectangle
inherits its matrices (100 in the custom capture). These are deliberately
not copied from line creation into rectangles.

Core native standard/custom geometry, style and twelve-decimal COM poses
match. Every part outside the edited page remains byte-identical. Native
custom-style line recreation also passes the existing line regression.
The full Visio suite passes 2,188 checks with 76 optional skips, followed by
four passing drawing-default checks including an additional regression for
the shared inherited-style dependency refusal. Core and viewer types and
the Visio ESM/CJS/declaration build pass.

Twelve browser workflows compare actual unselected fill/stroke colors and
opacity against native SVG, with real rectangle drags, undo/redo, public save
and reload in all six frameworks. The comparison clears selection before
measuring paint, because selection intentionally substitutes the accent
stroke. The shared native SVG helper supports native rect primitives as
well as paths. The existing vanilla line creation comparison also passes.

Other default styles and documents, text paint, gradients/effects, exact
pixels, native UI gesture semantics and native Office reopen acceptance
remain unverified. Each fresh owned invisible Visio instance quit; the
user's existing instance remained untouched.

## Native ellipse creation, resize and movement

The recorder's `-GridAligned -IncludeRectangle -IncludeEllipse` captures
`visio-line-movement-89d6ed4b25e74a448081428d6f98f06d` (standard defaults)
and `visio-line-movement-1fbaebe10f6144649668cd0e169b5fc2` (distinct custom
defaults) contain native rectangle and ellipse references. The ellipse starts
with Width 2, Height 1, PinX 6.5 and PinY 5.5. Native edits set Width 3,
Height 2, PinX 7 and PinY 6 and save/export `ellipse-edited.vsdx` and SVG.
[DrawOval](https://learn.microsoft.com/en-us/office/vba/api/visio.page.drawoval)
uses the native ellipse drawing method; the
[documented shortcut](https://support.microsoft.com/en-us/accessibility/visio/keyboard-shortcuts-for-visio)
is Ctrl+9. COM authoring is not proof of every native pointer gesture.

`create-ellipse` shares rectangle transform/default-style construction and
command snapshot/scaling validation. The native Geometry1 Ellipse row uses
X/C = Width*0.5, Y/B = Height*0.5, A = Width*1 and D = Height*1, with drawing
length units on A-D. Construction and resize proof share these control factors.
The existing dependency graph and recalculator update these formulas; altered
axes or missing native cache/formula proof are refused.

Five focused core checks pass, including creation/move/resize/delete with
round-trip part preservation, snapshots/scaling/invalid input, changed-axis
refusal and two native style-family comparisons. Native geometry/styles
match exactly and COM poses agree to twelve decimals before/after edits.
The full Visio suite passes 2,194 checks, with 76 optional skips.
Core/shared UI builds and types pass. UI command/scaling tests pass fourteen
checks with one optional native scale skip. The clean local core/UI package smoke test
also passes all 96 SSR entry imports and registration of 45 DOM tags.

`native-shape-create.spec.ts` shares the former rectangle workflow. Twenty-four
standard/custom workflows cover both shapes across six frameworks, including
twelve ellipse drags and resize/move control sequences, per-step history and
public save/reload. Native SVG physical geometry extents agree to three
decimals; unselected fill/stroke colors and opacity match computed native SVG
styles. The independent SVG extent registration uses each primitive's own
geometry box, not the page viewBox. A line creation regression and the rectangle
toolbar workflow also pass: 26 passing browser checks in total.

Core move/resize comparison targets the native ellipse; other native line
references in the recorder also move as part of its original probe. It is not
a whole-document equality claim. Other control axes and formulas, rotated or
scaled browser ellipse authoring, gradients/effects, text paint, contour and
exact pixels, native snapping/gesture semantics and native Office reopen
acceptance remain unverified. Fresh owned invisible instances quit; the
existing user Visio instance remained untouched.

## Native scaled rectangle and ellipse authoring

Fresh Visio 16 captures extend the shared shape workflow at these physical
drawing-to-page ratios:

- 0.5: `visio-line-movement-72974162251b475eac5d420e9da753b9`, DrawingScale 2,
  PageScale 1, standard defaults.
- 2: `visio-line-movement-ff0b777d5f5e4324a1270aa05e7d5084`, DrawingScale 0.5,
  PageScale 1, distinct custom defaults.
- 3: `visio-line-movement-1e61ca1a725c4c24893388aabd4b0c4f`, DrawingScale 1,
  PageScale 3, standard defaults.

The recorder keeps the same physical original rectangle/ellipse extents and
ellipse resize/move target, while native drawing-unit caches differ. The core
comparison uses the public drawing-inch API. Browser drags use physical page
inches through existing `visioPageEditToDrawing`; the experimental geometry
controls retain their documented drawing-inch contract. Test pointer coordinates
and native COM translations explicitly account for the ratio. No second unit
conversion or gesture implementation was added.

The ratio 0.5 capture places both boxes beyond the reduced paper dimensions.
The former box-tool clamping collapsed the drag to zero size, unlike native
DrawRectangle/DrawOval. All three tools now share unbounded page-point conversion
for start, preview and release, preserving the native off-paper coordinates.
This verifies retained geometry, not native off-paper visibility or snapping.

Eight core checks pass, including all five native default/scale captures.
Created and edited ellipse geometry/styles match exactly, physical dimensions
and native COM poses agree to twelve decimals, and every part except the edited
page XML is byte-for-byte unchanged. Native SVG page dimensions and primitive
extents provide independent physical measurements. The browser workflow covers
creation, ellipse resize/move, per-step undo/redo, exported VSDX and reload across
six bindings for both shapes; unselected colors/opacity match native SVG.
All 60 browser workflows pass, including 36 at the three new scales. Core/UI
builds and types, 20 UI command/geometry checks (one optional skip), viewer
types and the full viewer check command also pass. Viewer package checks use
released dependencies; the native workflows exercise the rebuilt local core/UI.

Other scales, rotated authoring, snapping/gesture semantics, contour/exact
pixels and native Office reopen acceptance remain unverified. Each fresh owned
invisible Visio instance quit, and the user's existing instance was untouched.

## Native local 2D shape rotation

The shared recorder's `-RotationDegrees` switch assigns native Angle.ResultIU
to both rectangle and ellipse after the original move/resize capture. It saves
`rotated.vsdx` and `rotated-page.svg`, recording angle caches and native XYToPage
poses. Sources use `ellipse-edited.vsdx`; comparison targets one shape at a time.

- `visio-line-movement-041ef57111794624b34bf9b3e7a5b48d`: 30 degrees, ratio 1.
- `visio-line-movement-64f4394fe5a34742837f7efa4a5645eb`: -45 degrees, ratio 2,
  distinct custom defaults.
- `visio-line-movement-afe4954faefa45f488ddc70e57cdd3e7`: 90 degrees, ratio 0.5.
- `visio-line-movement-51a9a06dfd5a4f3a99a0480a09823f8e`: 210 degrees, ratio 3.

Microsoft documents the parent-relative
[Angle cell](https://learn.microsoft.com/en-us/office/client-developer/visio/angle-cell-shape-transform-section)
and 2D UI rotation protection through
[LockRotate](https://learn.microsoft.com/en-us/office/client-developer/visio/lockrotate-cell-protection-section).
COM assignment is evidence of the saved transform, not native pointer-handle
or keyboard gesture equivalence.

`rotate-shape` uses the existing command snapshot, local leaf admission,
protection-style resolution, cell writer, dependency graph and recalculator.
The command holds dimensions and rotation pin fixed, writes an explicit RAD
angle and refuses active local/inherited LockRotate, protected angle formulas,
unknown affected dependencies, glued shapes, groups and 1D shapes. A no-op
angle preserves original bytes. Page scaling does not scale angular values.
The UI builds its new angle input and button with typed DOM creation and
shares the existing draft, selection, worker and transaction lifecycle.

Eighteen focused core checks pass, including eight native shape/scale cases,
angular unit normalization, snapshots, dependency recalculation, no-op bytes,
part preservation and unsupported-scope refusal. The broad Visio checkpoint
passes 2,210 checks with 80 optional skips. Eight geometry-control UI checks
pass. Native poses agree to twelve decimals; local geometry and effective
styles match exactly. Browser comparisons also measure transformed primitive
bounding boxes to three decimals and computed solid colors/opacity against
native SVG; these bounding boxes do not prove ellipse contour or exact pixels.
All 48 browser rotation workflows pass across six bindings, four native angles
and both shape kinds, including undo/redo and exported VSDX reload.
Core/UI builds and type checks pass, as does the full viewer check command.
Viewer consumer-package checks use released dependencies; native rotation
workflows exercise rebuilt local core/UI code.

Rotation handles, native shortcut semantics, grouped/master/glued and line
rotation, formulas that intentionally move the pin, other angles/scales,
text and gradient paint, exact pixels and native Office reopen acceptance
remain unverified. All fresh owned invisible instances quit; the existing
user Visio instance remained untouched.

## Native pointer rotation about the saved pin

The shared native recorder now saves `rotation-source.vsdx` before assigning
Angle and supports `-OffCentrePin`. Capture
`visio-line-movement-ebd3ce4c0c3643a89ae8b25f783d9e73` sets LocPinX to one quarter
of Width and LocPinY to three quarters of Height, then rotates both local
rectangle and ellipse by -30 degrees. The four previous captures provide
30, -45, 90 and 210 degree targets at ratios 1, 2, 0.5 and 3.

The parser retains source PinX/PinY/Angle through its existing transform reader
and page scaling. The browser never derives the rotation pin from a bounding
box centre. PowerPoint's DOM-free anchored rotation and viewport handle
placement move into shared core geometry, keeping PowerPoint's normalized
angle contract and adding opt-in signed turns. Visio's existing endpoint
capture/cancellation lifecycle is shared by both handle kinds. A viewport SVG
keeps an off-paper handle interactive without changing paper clipping, and
scroll/resize repositions the handle. Preview does not mutate source/history;
release submits one existing rotate-shape transaction.

Browser comparisons use real pointer arcs with an off-centre grip, cancel the
first gesture through Escape, then exercise rotation, undo/redo, export and
reload. The saved angle matches delivered pointer coordinates to twelve
decimals. Chromium's pointer coordinates introduce small native-target angular
error, bounded below 0.000002 radians; transform bounds account for that error
and shape size. Geometry and effective styles match native targets exactly.
This does not prove that native Visio uses the same gesture, snapping, handle
placement or preview paint. Group/master/glued/1D rotation, exact contour
pixels, text/gradient paint and Microsoft Visio reopen remain open.

All 60 pointer-rotation workflows pass across the six bindings, five native
captures and both shape kinds, including a viewport resize before each drag.
The extracted event lifecycle also passes 49 native endpoint workflows.
Six scroll-boundary checks retain the selected handle and a stable canvas extent.
Thirty-six focused core checks pass (one optional capture skipped), including
source-pin scaling and nonfinite metadata rejection; 62 PowerPoint compatibility
checks and two handle ownership checks pass. The broad Visio checkpoint passes
2,218 checks with 95 optional skips. Core/UI builds, strict/PowerPoint types,
all packed core entry imports and the full Visio viewer check pass. Published
viewer-package checks still consume released dependencies; browser workflows
exercise rebuilt local core/UI. The existing user Visio instance was untouched.

## Shared live rotation preview

Pointer rotation now renders a temporary pose through renderPage, using the
same geometry, text, fill/line paint and live resource management as the canvas.
Its affine pose composes the existing native pin-relative transform and shared
affine multiplication. The original selected SVG is temporarily hidden at its
existing paint position; the preview has no selection semantics or pointer
hit targets. Cancellation and release restore visibility and dispose resources.
Focus stays in the viewport so hiding a focused SVG shape cannot send Escape
outside the gesture owner. Document geometry/history are unchanged until release.

Ten native rectangle/ellipse source-target pairs agree to twelve decimals,
including four page scales, signed angles beyond 180 degrees and the custom
local pivot capture. Two additional core checks cover flips, signed turns,
no-op pose identity and invalid metadata. Browser tests compare preview and
committed coefficients, source bytes after Escape and during dragging, and
save/reload behavior. These comparisons establish the temporary pose and
transaction lifecycle; native gesture increments, handle appearance, text and
gradient pixels, preview raster paint and Office reopen are still unverified.

All 60 strengthened native pointer workflows and six scroll-boundary checks
pass across all bindings. Preview coefficients agree with saved/reloaded poses
to twelve decimals. Original source bytes remain identical immediately after
Escape and before release. Twelve focused core preview checks and 24 renderer/
lifecycle checks pass; replaced/cancelled raster URLs are revoked while the
original canvas URL remains owned by the canvas. The broad core checkpoint
passes 2,220 checks with 87 optional skips. Core/UI types and builds pass.
