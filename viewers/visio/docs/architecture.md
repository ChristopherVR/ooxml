# Architecture

## Ownership

- `ooxml-core/visio` in sibling `ooxml`: package intake, XML, model, inheritance, cached geometry, structured diagnostics. No viewer UI belongs here.
- `src/controller.ts`: document/view state, selection, events, latest-load-wins behavior.
- `src/document-text-search.ts`: bounded literal matching over normalized visible shape text, with immutable page-scoped results. The shared controller owns navigation and reentrancy checks.
- `src/render-svg.ts` and `src/render-text.ts`: SVG presentation and rendering warnings.
- `src/export-svg.ts`: bounded, static current-page serialization through the same SVG renderer, with embedded raster resources and compatibility metadata.
- `src/print-snapshot.ts`: side-effect-free selected-page artifacts, aggregate budgets and preserved diagnostics. It does not create print frames or implement printer policy.
- `src/viewer-element.ts`: shared browser surface.
- `src/contract.ts` and `src/binding.ts`: properties, events, client-only mount/update/load/fit/exportSvg/destroy lifecycle. Reentrant newer updates supersede the remaining older patch.
- Framework wrappers: framework lifecycle and event/prop forwarding only. No per-framework parser, rendering or geometry fork.

## Boundaries

The engine currently uses cached ShapeSheet values. This is not comprehensive formula evaluation. Browser font metrics, wrapping, mixed paragraph spacing/indentation, RTL and bullets are implemented with bounded work; these do not establish exact Visio text fidelity.

Document XML/HTML is not injected into the DOM. SVG nodes are created through DOM APIs. No document upload, telemetry or remote-document fetch is provided. Package and relationship checks belong to the core. The browser component parses in a dedicated worker with cancellation and a 15-second parent timeout. Independent scene, metadata, text-work and decoded-raster budgets protect rendering. Image resources are shared within each render and revoked on replacement, disconnect or disposal. Font load listeners trigger text reflow and are removed on disconnect. These defenses are not a security certification.

No native Visio writer exists. The current viewer cannot provide edit/save round-trip guarantees.

## Verification

Core unit and fixture tests establish their asserted parse/model cases. Viewer tests establish lifecycle and view behavior. Browser tests establish actual interaction. Genuine Visio-authored fixtures, known reference renders and measured comparisons remain necessary for fidelity claims. Framework runtime tests are separate from wrapper compilation.

## Website

Static multi-page Vite build: home, guide, capability ledger, architecture and a separate playground. Links are relative and assets local. No third-party fonts or scripts. No parser bundle on the landing page.

The documentation design adapts the [pptx-viewer reference](https://github.com/ChristopherVR/pptx-viewer). The in-depth review is in `research/pptx-pages-design-review.md`.

## Publication

The package is private and unpublished. There is no deployment workflow. Publication and access controls require a separate authorized decision. A private repository is not proof of private website access.
