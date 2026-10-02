# Visio Viewer

A private, local-first Visio viewing project. One headless controller, one SVG renderer and one custom element, with thin React, Vue, Angular, Svelte, Solid and vanilla adapters. Format logic belongs to the sibling `ooxml` repository's new `ooxml-core/visio` area.

**This is an early implementation, not Microsoft Visio parity.** It cannot edit or save native drawings. The [capability ledger](docs/parity.md) separates implemented code, tested evidence and missing functionality. Real upstream drawings and embedded previews have been inspected with a secondary renderer, but no controlled Microsoft Visio full-page comparison has passed.

## Local setup

Node.js 22.12 or newer is required. The Visio area is not published in `ooxml-core` yet, so this private repository uses a sibling checkout.

```sh
node scripts/setup-core.mjs
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
npm run check
npm run dev
```

The setup script clones the public core at the pinned revision in `integration/core-revision.txt`, applies the included local integration patch, installs dependencies from a pinned lock and builds only the Visio subpath. It refuses mismatched revisions, staged changes, different source edits or a different npm lock in an existing sibling checkout without overwriting them. The patch is a temporary development bridge, not a second canonical format implementation. Core changes must be reviewed and released separately before replacing the local dependency with a published version.

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
viewer.destroy(); // idempotent; later commands reject
```

Importing the module is SSR-safe; mounting and element registration are browser-only. `ViewerController` is DOM-free, so another renderer can subscribe to the same state/events. `renderPage` is the SVG renderer port. See [architecture](docs/architecture.md) and `packages/bindings/README.md` for framework entry points.

## Static SVG export

`exportPageSvg(model, pageIndex, { maxBytes })` and each mounted/native handle's `exportSvg()` return a current-page snapshot without changing the document, selection or viewport. Export includes supported backgrounds, embedded raster resources and visible compatibility notes in SVG description/metadata. There are no remote assets, automatic downloads or editable VSDX output. The demo downloads only after an explicit button press.

The UTF-8 ceiling is 16 MiB; callers can lower it. Conservative preflight checks can reject content whose eventual serialization would be smaller. Fonts are not embedded and text/layout remain approximate. Physical page dimensions are preserved; downstream rasterizers must cap their output dimensions and pixel area before allocating an image surface. PDF, bitmap export and print layout remain unimplemented.

Optional Linux secondary pixel tests run with `node scripts/test-svg-rasterization.mjs` after building, using system Python 3, librsvg/Cairo and the core's native canvas dependency. These checks are separate from browser or Microsoft Visio validation.

## Safety and limits

The core validates OPC relationships, rejects DTD/entities and suspicious ZIP paths, and bounds archive input/count/declared expansion/actual streamed output and XML complexity. The viewer checks file size before Blob allocation and validates externally supplied scenes before drawing, including actual raster structure, detected MIME and intrinsic dimensions. Raster validation checks structure/checksums, not full entropy decoding; native decoder robustness remains relevant. External relationships are never fetched, document strings are text nodes, and unsupported features are reported.

The browser custom element uses a dedicated parsing worker with cancellation and a parent-side 15-second limit. The headless controller and environments without Worker support use cooperative core limits. SVG/text rendering still runs on the main thread with separate scene, raster, metadata and text-work limits. The project is not yet security-certified or production-hardened.

## Publication

No npm publication, GitHub Pages deployment or public hosting is configured. A private GitHub repository must be created through an authorized signed-in connection before pushing these files. The local `ooxml` changes have not been pushed or released.
