# Visio Viewer

A public beta, local-first Visio viewing project. One headless controller, one SVG renderer and one custom element, with thin React, Vue, Angular, Svelte, Solid and vanilla adapters. Format logic belongs to the sibling `ooxml` repository's new `ooxml-core/visio` area.

**This is an early implementation, not Microsoft Visio parity.** The shared viewer has experimental source-backed plain-text editing, bounded undo/redo and explicit VSDX-copy export. Native Visio reopening remains unverified. The [capability ledger](docs/parity.md) separates implemented code, tested evidence and missing functionality. Real upstream drawings and embedded previews have been inspected with a secondary renderer, but no controlled Microsoft Visio full-page comparison has passed.

## Local setup

Node.js 22.12 or newer is required. The Visio area is not published in `ooxml-core` yet, so this repository uses a pinned sibling checkout.

```sh
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
npm run check
npm run dev
```

The setup script clones the public core at the pinned revision in `integration/core-revision.txt`, installs dependencies from a pinned lock and builds only the Visio subpath. The current pin includes the published geometry commands and requires no integration patch. It refuses mismatched revisions, staged changes, different source edits or a different npm lock in an existing sibling checkout without overwriting them. The setup still supports an explicit temporary patch for development. Core changes must be reviewed and released separately before replacing the local dependency with a published version.

Open the printed local URL for the documentation landing page or `/demo/` for the viewer workspace. `.vsdx` input stays in your browser; there are no uploads, telemetry, external fonts or document URL fetches.

## Commands

- `npm run typecheck`: strict viewer and demo TypeScript
- `npm test`: controller, renderer, scene safety and binding tests
- `npm run check:bindings`: actual native framework lifecycle/type checks
- `npm run check:core`: both core TypeScript projects and focused Visio parser tests
- `npm run docs:sync`: regenerate the Markdown capability table from its canonical HTML table
- `npm run build`: reusable ESM/declarations and static multipage site in `site-dist/`
- `npm run check`: all non-browser checks and production build
- `npm run test:browser`: Playwright Chromium workflows (requires a supported browser environment)
- `npm run test:corpus -- LIBVISIO_CHECKOUT POI_CHECKOUT`: optional hash-pinned real/security fixture regressions; see [corpus setup](docs/corpus.md)

The Playwright config accepts `CHROMIUM_PATH`. Install Playwright's Chromium or point it at a system Chromium. The development environment's IPC policy blocked launching Chromium, and the cloud browser blocked localhost, so browser screenshots and visual verification are not claimed. Source/static audits, DOM tests and production builds were performed instead. Consult `docs/verification.md` for current evidence.

## Framework-neutral API

```ts
import { mountViewer } from '@christophervr/visio-viewer';
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

Importing the module is SSR-safe; mounting and element registration are browser-only. `ViewerController` is DOM-free, so another renderer can subscribe to the same state/events. `renderPage` is the SVG renderer port. See [architecture](docs/architecture.md) and `packages/bindings/README.md` for framework entry points.

## Find diagram text

The shared search controls find literal text across visible shapes and pages. Programmatic hosts use `controller.setSearchQuery(query)`, `nextSearchResult()`, `previousSearchResult()` and `selectSearchResult(index)`; immutable query/results are available in `controller.state.search`. Querying leaves selection alone until explicit navigation. A new document clears the search, and newer reentrant host actions supersede old result navigation.

Search returns one result per matching shape, with a bounded text preview. It uses ECMAScript lowercase comparison, not regex or language-aware collation. Limits are 256 query characters, 500 matching shapes, 32,768 indexed characters per shape and 2 million indexed characters overall. The status explicitly identifies partial results. Hidden subtrees and suppressed group text are excluded; background pages are indexed once as separate pages. Per-occurrence highlighting and Shape Data search are not implemented.

## Display layers

Use the shared Layers controls or any mounted/native handle's `setLayerVisibility(pageId, layerId, visible)` to override display visibility. `visible: null` restores one saved flag; `resetLayerVisibility(pageId?)` resets one source page or the entire document. Source-page IDs distinguish foreground and background layers. Frozen overrides are available in `controller.state.layerVisibilityOverrides`.

Layer changes synchronize rendering, selection and search without changing the document. Guides, NoShow and incompletely described legacy hidden shapes cannot be revealed by a layer override. The accessible panel shows at most 200 unique layers and announces truncation; the API allows at most 25,000 document-scoped overrides. SVG and print artifacts continue to use saved visibility. Layer colors, editing and saved print policy are not implemented.

## Static SVG export

`exportPageSvg(model, pageIndex, { maxBytes })` and each mounted/native handle's `exportSvg()` return a current-page snapshot without changing the document, selection or viewport. Export includes supported backgrounds, embedded raster resources and visible compatibility notes in SVG description/metadata. SVG export has no remote assets or automatic downloads. The demo downloads only after an explicit button press.

The UTF-8 ceiling is 16 MiB; callers can lower it. Conservative preflight checks can reject content whose eventual serialization would be smaller. Fonts are not embedded and text/layout remain approximate. Physical page dimensions are preserved; downstream rasterizers must cap their output dimensions and pixel area before allocating an image surface. PDF, bitmap export and print layout remain unimplemented.

Optional Linux secondary pixel tests run with `node scripts/test-svg-rasterization.mjs` after building, using system Python 3, librsvg/Cairo and the core's native canvas dependency. These checks are separate from browser or Microsoft Visio validation.

## Prepare drawing-page artifacts

`createPrintSnapshot(model, { pageIndices: [0, 2] })` preserves explicit page order and returns deeply frozen, separate SVG records with drawing dimensions and all export diagnostics. Each mounted/native handle's `createPrintSnapshot()` captures its current page. It rejects document replacement during preparation without changing view state. There are no frames, dialogs, downloads or printer commands.

The snapshot uses saved display appearance. Layer Print, NonPrinting and saved printer settings are not applied; drawing dimensions are not printer paper or print scale. The result includes these limitations. Guards bound selected pages, aggregate bytes, repeated backgrounds, raster dimensions, text/path work and repeated whole-model validation. Callers may lower limits. Separate SVGs must remain isolated because their internal resource IDs can repeat. This is preparation for a future print workflow, not native Visio printing support.

## Safety and limits

The core validates OPC relationships, rejects DTD/entities and suspicious ZIP paths, and bounds archive input/count/declared expansion/actual streamed output and XML complexity. The viewer checks file size before Blob allocation and validates externally supplied scenes before drawing, including actual raster structure, detected MIME and intrinsic dimensions. Raster validation checks structure/checksums, not full entropy decoding; native decoder robustness remains relevant. External relationships are never fetched, document strings are text nodes, and unsupported features are reported.

The browser custom element uses a dedicated parsing worker with cancellation and a parent-side 15-second limit. The headless controller and environments without Worker support use cooperative core limits. SVG/text rendering still runs on the main thread with separate scene, raster, metadata and text-work limits. The project is not yet security-certified or production-hardened.

## Metafile compatibility checks

Embedded enhanced metafiles now receive bounded record-level compatibility diagnostics. The checker rejects unsupported formats, styles, mapping states and resource amplification before any conversion. It does not execute embedded WMF data, fonts, images, scripts or external references. The disposable parser worker converts only a narrow admitted line/rectangle/ellipse and stock-object subset through the released converter package, with document budgets and a parent-owned hard deadline. Direct core parsing remains inspection-only unless a trusted converter is explicitly supplied in an isolated host. All five previously inspected real EMF media parts remain unsupported; structural admission alone does not promise rendering or Visio fidelity.

The headless core contains a tested neutral-vector sanitizer and transport validator. The parser worker uses the released converter for a bounded primitive EMF subset, which shares live, SVG export and immutable print rendering. The non-worker parser fallback remains converter-free. Generated path-only clipping tests agree with the local Skia SVG backend; the installed librsvg backend ignores nested clip-path intersections. Browser and native Visio comparisons remain required. See [the detailed adoption review](docs/research/emf-adoption-review.md).

## Publication

Install `visio-vanilla-viewer` or a framework package: `visio-react-viewer`, `visio-vue-viewer`, `visio-angular-viewer`, `visio-svelte-viewer` or `visio-solid-viewer`. Each ships the shared viewer, its adapter and both workers, and re-exports the document API. `visio-core` provides the headless `ooxml-core/visio` API. The hourly release workflow publishes implementation changes with npm provenance after package and browser checks. Try [the live demo](https://christophervr.github.io/visio-viewer/demo/). Rendering and editing remain beta features.

## Experimental local plain-text editing

After `load(bytesOrBlob)`, select a local shape and use the shared text controls, or call `replacePlainText(pageId, shapeId, text)` on any mounted/native handle. `undo()` and `redo()` use bounded document history; `cancelEdit()` cancels pending work. All six adapters forward the same methods and `document-change` event, whose detail is `{ document, dirty, kind: 'edit' | 'undo' | 'redo' }`. Inspect `controller.state.edit` for source availability, busy/dirty status, undo/redo availability, history truncation, errors and diagnostics.

Only imported source-backed documents can be edited or exported as VSDX. Assigning a model with `document` does not supply editable package bytes. Core validation rejects master-linked shapes, rich text, fields, signed packages and macro-enabled content. XML/package edits run in an isolated worker; framework wrappers do not implement format logic. Plain-text edits do not recalculate text-dependent formula caches.

`exportVsdx()` returns `{ bytes, dirty, diagnostics }` for an explicit downloaded copy. It does not overwrite the source file, upload it or claim native round-trip fidelity. The shared UI downloads only after the user's explicit action. History is bounded and may discard older undo states, reported by `historyTruncated`. General drawing, style editing, rich-text editing and native Visio reopen verification remain unsupported.

## Experimental geometry editing

The shared geometry controls expose create rectangle, move, resize and safe delete.
Every native/mounted framework handle also forwards `applyEdits(edits: readonly VisioEdit[])`
for an atomic batch through the same worker, history and cancellation path.

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

This is a narrow admitted subset. Masters/groups/foreign shapes, connectors/glue,
unsafe protection/redirection, unsupported affected formulas, ambiguous package
dependencies and referenced deletion fail with a visible error. Relative line
geometry scales; absolute line geometry needs a supported dimension dependency.
See the core `src/visio/README.md` for the command contract, numeric limits,
exclusion reasons and next expansions.

Generated/imported simple-shape tests and a local Chrome workflow pass. The 19
accepted public corpus diagrams currently admit zero geometry edits because
non-page dependency scope cannot yet be proved independent. Thirteen also lack an
eligible local shape. Practical editing coverage for that corpus remains blocked;
next work is a scoped package/master/theme dependency graph and master-instance
editing, followed by glued endpoint routing. Native Visio reopen/fidelity is unverified.
