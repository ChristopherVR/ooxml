# Architecture

## Ownership

- `src/core/visio` (the `ooxml-core/visio` entry): package intake, XML, model, inheritance, cached geometry, ShapeSheet evaluation, editing commands, collaboration mapping (`ooxml-core/visio/collab`) and structured diagnostics. No viewer UI belongs here. The DOM-free view contract also lives in the core, in `src/core/visio/ui`: `contract.ts` (properties and events), `document-history.ts` (bounded source-backed edit history and VSDX-copy export state, with XML/package mutation running through a dedicated worker), `document-text-search.ts` (bounded literal matching over normalized visible shape text, with immutable page-scoped results) the collaboration option and presence types, and the other DOM-free editing and view plans (move, resize, formatting, draw paths, guides).
- `src/ui/src/visio` (published as `ooxml-ui/visio`): the browser viewer. `controller.ts` owns document/view state, selection, events, latest-load-wins behavior and transactional editing orchestration; `viewer-element.ts` is the shared `<visio-viewer>` surface; `binding.ts` is the client-only mount/update/load/fit/exportSvg/destroy lifecycle (reentrant newer updates supersede the remaining older patch); `render-svg.ts` and `render-text.ts` are SVG presentation and rendering warnings; `export-svg.ts` is bounded, static current-page serialization through the same SVG renderer, with embedded raster resources and compatibility metadata; `print-snapshot.ts` produces side-effect-free selected-page artifacts with aggregate budgets and preserved diagnostics (it does not create print frames or implement printer policy); the building blocks are `building-blocks-*.ts`.
- `viewers/visio/src/index.ts`: a re-export of `ooxml-ui/visio`.
- `viewers/visio/packages/bindings/src` (private) and `packages/{react,vue,angular,svelte,solid,vanilla}`: framework lifecycle and event/prop forwarding, and the native building-block compositions. No per-framework parser, rendering or geometry fork.

## Boundaries

The engine currently uses cached ShapeSheet values. This is not comprehensive formula evaluation. Browser font metrics, wrapping, mixed paragraph spacing/indentation, RTL and bullets are implemented with bounded work; these do not establish exact Visio text fidelity.

Document XML/HTML is not injected into the DOM. SVG nodes are created through DOM APIs. No document upload, telemetry or remote-document fetch is provided. Package and relationship checks belong to the core. The browser component parses in a dedicated worker with cancellation and a 15-second parent timeout. Independent scene, metadata, text-work and decoded-raster budgets protect rendering. Image resources are shared within each render and revoked on replacement, disconnect or disposal. Font load listeners trigger text reflow and are removed on disconnect. These defenses are not a security certification.

Experimental source-backed plain-text replacement supports bounded undo/redo and explicit VSDX-copy export. Core rejects master-linked shapes, rich text, fields, signed packages and macro content. Model-only documents are read-only. Formula caches are not recalculated and native Visio reopen verification remains outstanding; this is not an edit/save round-trip guarantee. The handle methods, events and limits are in the [viewer API](api.md).

## Safety and limits

The core validates OPC relationships, rejects DTD/entities and suspicious ZIP paths, and bounds archive input/count/declared expansion/actual streamed output and XML complexity. The viewer checks file size before Blob allocation and validates externally supplied scenes before drawing, including actual raster structure, detected MIME and intrinsic dimensions. Raster validation checks structure/checksums, not full entropy decoding; native decoder robustness remains relevant. External relationships are never fetched, document strings are text nodes, and unsupported features are reported.

The browser custom element uses a dedicated parsing worker with cancellation and a parent-side 15-second limit. The headless controller and environments without Worker support use cooperative core limits. SVG/text rendering still runs on the main thread with separate scene, raster, metadata and text-work limits. The project is not yet security-certified or production-hardened.

## Embedded metafiles

Embedded enhanced metafiles receive bounded record-level compatibility diagnostics. The checker rejects unsupported formats, styles, mapping states and resource amplification before any conversion. It does not execute embedded WMF data, fonts, images, scripts or external references. The disposable parser worker converts only a narrow admitted line/rectangle/ellipse and stock-object subset through the converter package, with document budgets and a parent-owned hard deadline. Direct core parsing remains inspection-only unless a trusted converter is explicitly supplied in an isolated host. All five previously inspected real EMF media parts remain unsupported; structural admission alone does not promise rendering or Visio fidelity.

The headless core contains a tested neutral-vector sanitizer and transport validator. The parser worker uses the converter for a bounded primitive EMF subset, which shares live, SVG export and immutable print rendering. The non-worker parser fallback remains converter-free. Generated path-only clipping tests agree with the local Skia SVG backend; the installed librsvg backend ignores nested clip-path intersections. Browser and native Visio comparisons remain required. See [the detailed adoption review](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/docs/research/emf-adoption-review.md).

## Verification

Core unit and fixture tests establish their asserted parse/model cases. Viewer tests establish lifecycle and view behavior. Browser tests establish actual interaction. Genuine Visio-authored fixtures, known reference renders and measured comparisons remain necessary for fidelity claims. Framework runtime tests are separate from wrapper compilation.

## Website

Static multi-page Vite build: home, guide, capability ledger, architecture and a separate playground. Links are relative and assets local. No third-party fonts or scripts. No parser bundle on the landing page.

## Publication

Eight npm packages expose the headless Visio API, the shared viewer with six native framework adapters, and the MCP tools. They depend on `ooxml-core` and `ooxml-ui`, which live in this repository and are released with them (see [releasing](releasing.md)). The repository's Pages build deploys the documentation and demos under `/ooxml/visio/`.
