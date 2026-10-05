# Viewer API

The framework-neutral API behind every viewer package. Each framework package
(`visio-react-viewer`, `visio-vue-viewer`, `visio-angular-viewer`,
`visio-svelte-viewer`, `visio-solid-viewer`, `visio-vanilla-viewer`) forwards
the same handle methods and events; see `packages/bindings/README.md` for the
native props, events and lifecycle of each adapter. Rendering and editing are
beta features; the [capability ledger](parity.md) lists what is missing.

## Mount and load

```ts
import { mountViewer } from 'visio-vanilla-viewer';
const viewer = mountViewer(container, {
	zoom: 1,
	events: { 'document-error': (error) => console.error(error) },
});
await viewer.load(file); // Blob, Uint8Array or ArrayBuffer
viewer.fit();
viewer.update({ pageIndex: 1 });
const snapshot = viewer.exportSvg(); // SVG string, dimensions, byteLength and diagnostics
const prepared = viewer.createPrintSnapshot(); // Frozen current-page artifacts, no printing
viewer.destroy(); // idempotent; later commands reject
```

Importing the module is SSR-safe; mounting and element registration are
browser-only. `ViewerController` is DOM-free, so another renderer can subscribe
to the same state/events. `renderPage` is the SVG renderer port. See
[architecture](architecture.md) for the module ownership.

## Find diagram text

The shared search controls find literal text across visible shapes and pages.
Programmatic hosts use `controller.setSearchQuery(query)`, `nextSearchResult()`,
`previousSearchResult()` and `selectSearchResult(index)`; immutable
query/results are available in `controller.state.search`. Querying leaves
selection alone until explicit navigation. A new document clears the search, and
newer reentrant host actions supersede old result navigation.

Search returns one result per matching shape, with a bounded text preview. It
uses ECMAScript lowercase comparison, not regex or language-aware collation.
Limits are 256 query characters, 500 matching shapes, 32,768 indexed characters
per shape and 2 million indexed characters overall. The status explicitly
identifies partial results. Hidden subtrees and suppressed group text are
excluded; background pages are indexed once as separate pages. Per-occurrence
highlighting and Shape Data search are not implemented.

## Display layers

Use the shared Layers controls or any mounted/native handle's
`setLayerVisibility(pageId, layerId, visible)` to override display visibility.
`visible: null` restores one saved flag; `resetLayerVisibility(pageId?)` resets
one source page or the entire document. Source-page IDs distinguish foreground
and background layers. Frozen overrides are available in
`controller.state.layerVisibilityOverrides`.

Layer changes synchronize rendering, selection and search without changing the
document. Guides, NoShow and incompletely described legacy hidden shapes cannot
be revealed by a layer override. The accessible panel shows at most 200 unique
layers and announces truncation; the API allows at most 25,000 document-scoped
overrides. SVG and print artifacts continue to use saved visibility. Layer
colors, editing and saved print policy are not implemented.

## Static SVG export

`exportPageSvg(model, pageIndex, { maxBytes })` and each mounted/native handle's
`exportSvg()` return a current-page snapshot without changing the document,
selection or viewport. Export includes supported backgrounds, embedded raster
resources and visible compatibility notes in SVG description/metadata. SVG export
has no remote assets or automatic downloads. The demo downloads only after an
explicit button press.

The UTF-8 ceiling is 16 MiB; callers can lower it. Conservative preflight checks
can reject content whose eventual serialization would be smaller. Fonts are not
embedded and text/layout remain approximate. Physical page dimensions are
preserved; downstream rasterizers must cap their output dimensions and pixel area
before allocating an image surface. PDF, bitmap export and print layout remain
unimplemented.

Optional Linux secondary pixel tests run with
`node scripts/test-svg-rasterization.mjs` after building, using system Python 3,
librsvg/Cairo and the core's native canvas dependency. These checks are separate
from browser or Microsoft Visio validation.

## Prepare drawing-page artifacts

`createPrintSnapshot(model, { pageIndices: [0, 2] })` preserves explicit page
order and returns deeply frozen, separate SVG records with drawing dimensions and
all export diagnostics. Each mounted/native handle's `createPrintSnapshot()`
captures its current page. It rejects document replacement during preparation
without changing view state. There are no frames, dialogs, downloads or printer
commands.

The snapshot uses saved display appearance. Layer Print, NonPrinting and saved
printer settings are not applied; drawing dimensions are not printer paper or
print scale. The result includes these limitations. Guards bound selected pages,
aggregate bytes, repeated backgrounds, raster dimensions, text/path work and
repeated whole-model validation. Callers may lower limits. Separate SVGs must
remain isolated because their internal resource IDs can repeat. This is
preparation for a future print workflow, not native Visio printing support.

## Experimental local plain-text editing

After `load(bytesOrBlob)`, select a local shape and use the shared text controls,
or call `replacePlainText(pageId, shapeId, text)` on any mounted/native handle.
`undo()` and `redo()` use bounded document history; `cancelEdit()` cancels
pending work. All six adapters forward the same methods and `document-change`
event, whose detail is `{ document, dirty, kind: 'edit' | 'undo' | 'redo' }`.
Inspect `controller.state.edit` for source availability, busy/dirty status,
undo/redo availability, history truncation, errors and diagnostics.

Only imported source-backed VSDX documents can be edited or exported as VSDX.
Binary VSD v11 imports use the released ole2 codec through ooxml-core: supported
explicit page/shape transforms, move/line geometry and plain text are previewed
with diagnosed fallback styles. Original legacy bytes remain privately retained;
editing, undo/redo and VSDX export are disabled for legacy sources. Earlier binary
versions, groups/masters, complex geometry, stencils, templates and macro-enabled
formats remain unsupported. Assigning a model with `document` does not supply
editable package bytes. Core validation rejects master-linked shapes, rich text,
fields, signed packages and macro-enabled content. XML/package edits run in an
isolated worker; framework wrappers do not implement format logic. Plain-text
edits do not recalculate text-dependent formula caches.

`exportVsdx()` returns `{ bytes, dirty, diagnostics }` for an explicit downloaded
copy. It does not overwrite the source file, upload it or claim native round-trip
fidelity. The shared UI downloads only after the user's explicit action. History
is bounded and may discard older undo states, reported by `historyTruncated`.
General drawing, style editing, rich-text editing and native Visio reopen
verification remain unsupported.

## Experimental geometry editing

The shared geometry controls expose create rectangle, move, resize and safe
delete. Every native/mounted framework handle also forwards
`applyEdits(edits: readonly VisioEdit[])` for an atomic batch through the same
worker, history and cancellation path.

```ts
await handle.applyEdits([
	{
		type: 'create-rectangle',
		pageId: '0',
		shapeId: '9',
		x: 4,
		y: 5,
		width: 3,
		height: 2,
		text: 'New shape',
	},
	{ type: 'move-shape', pageId: '0', shapeId: '1', x: 6, y: 7 },
]);
```

Coordinates are rotation-pin positions in drawing inches, bottom-left origin,
up-positive. Resizing holds the pin fixed. IDs are explicit. Existing admitted
local top-level 2D shapes can be edited, including shapes imported from other
producers. Core evaluates supported affected numeric ShapeSheet dependencies and
preserves untouched ZIP payloads. The viewer never mutates XML itself.

This is a narrow admitted subset. Masters/groups/foreign shapes,
connectors/glue, unsafe protection/redirection, unsupported affected formulas,
ambiguous package dependencies and referenced deletion fail with a visible
error. Relative line geometry scales; absolute line geometry needs a supported
dimension dependency. See the core
[`src/core/visio/README.md`](https://github.com/ChristopherVR/ooxml/blob/main/src/visio/README.md)
for the command contract, numeric limits, exclusion reasons and next expansions.
The [capability ledger](parity.md) records the current admitted set, which now
also includes bounded primitive master-instance moves with explicit local pins.

Generated/imported simple-shape tests and a local Chrome workflow pass. The 19
accepted public corpus diagrams currently admit zero geometry edits because
non-page dependency scope cannot yet be proved independent. Thirteen also lack an
eligible local shape. Practical editing coverage for that corpus remains blocked;
next work is a scoped package/master/theme dependency graph and master-instance
editing, followed by glued endpoint routing. Native Visio reopen/fidelity is
unverified.
