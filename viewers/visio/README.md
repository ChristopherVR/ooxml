# Visio Viewer

A private, local-first Visio viewing project. One headless controller, one SVG renderer and one custom element, with thin React, Vue, Angular, Svelte, Solid and vanilla adapters. Format logic belongs to the sibling `ooxml` repository's new `ooxml-core/visio` area.

**This is an early implementation, not Microsoft Visio parity.** It cannot edit or save drawings. The [capability ledger](docs/parity.md) separates implemented code, tested evidence and missing functionality. No Microsoft Visio reference-render comparison has been performed.

## Local setup

Node.js 22.12 or newer is required. The Visio area is not published in `ooxml-core` yet, so this private repository uses a sibling checkout.

```sh
node scripts/setup-core.mjs
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
npm run check
npm run dev
```

The setup script clones the public core at the pinned revision in `integration/core-revision.txt`, applies the included local integration patch, installs dependencies and builds only the Visio subpath. It refuses to overwrite an existing sibling checkout that lacks Visio. The patch is a temporary development bridge, not a second canonical format implementation. Core changes must be reviewed and released separately before replacing the local dependency with a published version.

Open the printed local URL for the documentation landing page or `/demo/` for the viewer workspace. `.vsdx` input stays in your browser; there are no uploads, telemetry, external fonts or document URL fetches.

## Commands

- `npm run typecheck`: strict viewer and demo TypeScript
- `npm test`: controller, renderer, scene safety and binding tests
- `npm run check:bindings`: actual native framework lifecycle/type checks
- `npm run build`: reusable ESM/declarations and static multipage site in `site-dist/`
- `npm run check`: all non-browser checks and production build
- `npm run test:browser`: Playwright Chromium workflows (requires a supported browser environment)

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
viewer.destroy(); // idempotent; later commands reject
```

Importing the module is SSR-safe; mounting and element registration are browser-only. `ViewerController` is DOM-free, so another renderer can subscribe to the same state/events. `renderPage` is the SVG renderer port. See [architecture](docs/architecture.md) and `packages/bindings/README.md` for framework entry points.

## Safety and limits

The core validates OPC relationships, rejects DTD/entities and suspicious ZIP paths, and bounds archive input/count/declared expansion/actual streamed output and XML complexity. The viewer checks file size before Blob allocation and validates externally supplied scenes before drawing. External relationships are never fetched, document strings are text nodes, and unsupported features are reported.

The browser custom element uses a dedicated parsing worker with cancellation and a parent-side 15-second limit. The headless controller and environments without Worker support use cooperative core limits. SVG/text rendering still runs on the main thread with separate scene, raster, metadata and text-work limits. The project is not yet security-certified or production-hardened.

## Publication

No npm publication, GitHub Pages deployment or public hosting is configured. A private GitHub repository must be created through an authorized signed-in connection before pushing these files. The local `ooxml` changes have not been pushed or released.
