# Safe multipage browser-print pipeline review

Review date: 2026-10-02. Viewer checkpoint: `2239716`.

This design review describes implementation options. Printing is not yet implemented or browser-verified.

## Findings that affect implementation

1. **Inline concatenation has a real ID collision.** `src/render-resources.ts:10,31` resets `visio-export-image-1` for every render. A local JSDOM reproduction exported two demo pages, one with PNG bytes and the other JPEG bytes, then found duplicate `visio-export-image-1` IDs and two references to `#visio-export-image-1` after concatenation. This proves duplicate IDs, not a particular browser's resulting pixels. SVG IDs must be unique in their node tree. Use separate SVG image resources, or rewrite all IDs and references consistently. [MDN SVG id](https://developer.mozilla.org/en-US/docs/Web/SVG/Reference/Attribute/id)
2. **Current whole-document limits are insufficient for composed print jobs.** A shared background is rendered once for every foreground page. Document validation counts its shapes and raster resources once. `exportPageSvg` has an independent 16 MiB ceiling for each page. Allowing 256 valid pages can still accumulate enormous markup, decoding, background work, and page area.
3. **Display-hidden and print-hidden are not interchangeable.** Core `shapes.ts:234-235` merges layer visibility, guides, and `NoShow` into one `hidden` bit. `VisioLayer.printable` exists, but the renderer does not use it. A print renderer cannot simply clear `hidden` to reveal printable layers: that can reveal guides or intrinsically hidden content. Microsoft documents `Visible` and `Print` separately. [Visible cell](https://learn.microsoft.com/en-us/office/client-developer/visio/visible-cell-layers-section), [Print cell](https://learn.microsoft.com/en-us/office/client-developer/visio/print-cell-layers-section)
4. **Drawing size is not printer setup.** Export preserves drawing inches; the core page model does not preserve printer paper, print orientation, margins, tiling, or scale. Microsoft permits drawing and printer sizes/orientations to differ. Printing one exported drawing onto one CSS page is useful, but must not be described as reproducing saved Visio print settings. [Microsoft page sizing](https://support.microsoft.com/en-us/visio/change-the-drawing-page-or-printer-paper-size), [Print orientation](https://learn.microsoft.com/en-us/office/client-developer/visio/printpageorientation-cell-print-properties-section), [PrintScale](https://learn.microsoft.com/en-us/office/vba/api/visio.document.printscale)
5. **Browser completion cannot establish successful printing.** The HTML standard permits ignored or nonblocking print calls. `afterprint` also covers preview closure. Never expose a result named `printed: true` from these events. [HTML printing algorithm](https://html.spec.whatwg.org/multipage/timers-and-user-prompts.html#printing), [MDN afterprint](https://developer.mozilla.org/en-US/docs/Web/API/Window/afterprint_event)

## Recommended first contract

Keep three separable stages. The first can ship without any dialog capability.

### 1. Build an immutable snapshot

`createPrintSnapshot(model, options)` produces an ordered, readonly list of page records containing page index/ID/name, inches, SVG text, byte length, per-page diagnostics, and snapshot-level limitations. It neither downloads nor inserts frames nor prints. Build the snapshot synchronously from the chosen document state before awaiting browser work; copy results/options, not the caller's entire model. Do not mutate or freeze the source model. Never cache raster validation solely by mutable byte-array identity across calls.

- Require an explicit bounded page list or a documented default. Recommended defaults: current page for the viewer action; all foreground pages in document order for a separately named all-pages action.
- All-pages selection excludes background pages as independent output sheets. An explicitly selected background page may still be exported on its own. Resolve each selected page's background chain through core `getVisioPageLayers`, preserving background-first order and using the selected foreground page's dimensions.
- Reject empty selections, duplicate indices, invalid/noninteger/out-of-range indices, and over-limit selections. Do not silently drop, sort, truncate, or rescale requested pages.
- Carry compatibility notes from every export. Show significant omissions before offering printing. Keep notes out of the printed drawing unless explicitly requested; metadata inside an SVG image is not visible disclosure.
- Say that this initial mode reproduces supported saved drawing appearance. It ignores viewport zoom, pan, selection, search highlighting, and temporary viewer filters. If future current-view layer overrides are supported, make that an explicit option captured into the snapshot.
- Do not advertise saved print-layer semantics until core preserves enough independent visibility/printability information. Add that normalization in core, never by reparsing package XML in the viewer. Multiple-layer print membership, group inheritance, and hidden-but-printable fixtures need primary-spec/native-reference evidence; do not infer their exact logic from display visibility.

### 2. Enforce job-wide budgets before browser resources

Retain the existing 16 MiB per-page conservative export budget. Add an aggregate budget independent of the per-page limit. Suggested conservative starting application limits are **32 selected pages, 32 MiB total serialized SVG, 100,000 composed shape instances, 100,000 composed geometry instances, 64 million decoded raster pixels counted once per distinct resource per output page, and 256 million raster-instance pixels**. These are product guardrails, not requirements from a standard.

Also cap physical page width/height and aggregate page area before creating SVG image resources. A conservative initial rule is each axis at most 4096 CSS pixels at 96 CSS pixels/inch, each page at most 16 million CSS-pixel area, and the job at most 64 million CSS-pixel area. This intentionally rejects some large-format drawings. The error must state the limitation and suggest selecting fewer/smaller pages. Do not silently shrink them. These layout proxies reduce risk but do not prove a printer driver's memory use at arbitrary DPI. [CSS absolute units](https://www.w3.org/TR/css-values-3/#absolute-lengths)

Compute work across resolved page/background compositions, including repeated backgrounds. For isolated SVG images, assume each page may decode its own copy; do not rely on browser cross-Blob deduplication. Reserve wrapper/CSS/diagnostic overhead too if exporting a combined HTML artifact. If a standalone HTML artifact embeds base64 SVG, account for that extra 4/3 expansion before allocating it.

Perform cheap selection, dimensions, composition-work and estimated-byte preflight for the whole job before building DOM/encoding images. Validate actual byte totals as each page is produced. Fail atomically: discard accumulated outputs and create no print frame when any page fails. Any configurable limit may lower but not raise the hard ceiling. Keep validation/preflight logic shared rather than duplicating the exporter algorithm. Revalidating the entire image-heavy model twice per exported page is correct but potentially expensive; a refactor may pass a private call-scoped validated context, never a persistent untrusted validation bypass.

### 3. Prepare a disposable browser session

Prefer one `<img>` whose source is an application-created SVG Blob URL for each output sheet. Keep the vector source and embedded PNG/JPEG/GIF images; do not rasterize via an unbounded canvas. Separate image documents isolate SVG IDs and cannot inherit host page CSS. Image-context SVG commonly disallows scripts and external resources, with embedded data resources allowed. Do not substitute `<object>`, `<embed>`, or SVG documents in individual iframes, which do not provide the same restrictions. [MDN SVG image restrictions](https://developer.mozilla.org/en-US/docs/Web/SVG/Guides/SVG_as_an_image)

The containing print iframe should be a generated static shell with no untrusted HTML. Construct labels through `textContent`; use only validated numbers and generated names such as `sheet-0` in CSS. Do not interpolate page names, font names, IDs, hyperlink targets, document styles, filenames, or arbitrary URLs into HTML/style text. Only locally generated Blob URLs are assigned to image `src`. Keep the existing renderer's element creation, color sanitation, raster validation, and omission of links/scripts/foreignObject.

Recommended defense in depth for a same-origin frame: sandbox tokens `allow-same-origin allow-modals`, never `allow-scripts`; CSP `default-src 'none'; img-src blob: data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'`. The inline style exception is only for application-generated CSS. A missing `allow-modals` flag blocks `print()` and suppresses print events. Host CSP/Trusted Types may further restrict preparation; report that accurately rather than weakening host policy. [MDN iframe sandbox](https://developer.mozilla.org/en-US/docs/Web/HTML/Reference/Elements/iframe#sandbox)

Do not use `display:none` for a frame that must lay out printable content. Keep it offscreen, noninteractive and out of sequential focus while retaining a layout viewport; verify this in each browser. Load the shell, install handlers, add images, await each `decode()` (or load/error fallback), then wait for required font/layout readiness. Set an abortable preparation timeout so corrupt/blocked resources cannot hang forever. With isolated images, frame font readiness alone does not prove embedded SVG font metrics; maintain the existing unembedded-font diagnostic. [MDN image decode](https://developer.mozilla.org/en-US/docs/Web/API/HTMLImageElement/decode)

## Page geometry

Generate one block sheet per snapshot page. Use generated named page rules, for example `@page sheet-0 { size: 8.5in 11in; margin: 0 }`, and apply `page: sheet-0` to the matching block. Set exact sheet width/height, block image display, border-box sizing, no margins/padding/border, and explicit clipping to the drawing page bounds. Force breaks between sheets, not after the last sheet; keep each sheet together. Do not use viewport units or viewer zoom. Two explicit lengths already determine orientation; do not add an unrelated landscape rotation.

Named pages and `size` describe intended output, but browser/printer settings may override paper, scale, margins, colors, headers, or footers. Browser support and mixed-size output require actual browser evidence. `page-orientation` rotates after layout and is not a substitute for the intended width/height. CSS print color adjustment, if added, is only a request. [W3C page size and named pages](https://www.w3.org/TR/css-page-3/), [MDN page size](https://developer.mozilla.org/en-US/docs/Web/CSS/Reference/At-rules/@page/size)

## Lifecycle and state

Use states `preparing`, `ready`, `requested`, `closed`, `failed`, `disposed`. Preparation must never invoke `print()`. A visible explicit user action on a ready session invokes printing once; disable or reject repeated clicks while that job is active. Tests replace the print method and never open real dialogs.

A session owns its frame, Blob URLs, image listeners, print listeners, timers, abort handlers and readiness promises. Disposal is idempotent. Register completion handlers before invoking print. Dispose on explicit cancel, preparation failure, stale-generation completion, viewer destruction, document replacement, and `afterprint`. Resolve all pending promises predictably and ignore late load/decode/event callbacks. Remove event listeners and clear timers even when frame creation or URL creation fails halfway through. Restore focus only when it is still appropriate for the initiating viewer.

Do not treat immediate return from `print()` as a guaranteed closed preview. For a browser that never emits `afterprint`, retain a bounded single session with explicit cancel/dispose and cleanup on document replacement/destroy; do not repeatedly create leaked frames. A fallback lifecycle policy needs real-browser verification, and cannot turn a timeout or focus event into a claim that printing succeeded. One frame, selected-page caps and explicit disposal bound resources without silently deleting a preview after an arbitrary short interval.

Viewer state and ordinary state events must be identical before/after successful snapshot creation, rejected preparation, cancelled print, and teardown. Framework bindings should forward the same shared API; no framework-specific rendering or print state machines.

## Required regression coverage

1. **Selection:** current, ordered foreground pages, explicit background, empty/all-background selection, repeated/invalid indices, page deletion or document replacement while preparing.
2. **Composition:** nested backgrounds, missing/cyclic backgrounds, repeated shared background, differing background dimensions, foreground extent clipping, no extra background sheets.
3. **Isolation/security:** distinct PNG/JPEG pages with colliding original symbol IDs; gradients, markers, clips and use references; malicious names containing closing tags/CSS; unsafe fill/font strings; all URLs restricted to generated Blob/data/fragment references; no host CSS leakage or network fetches.
4. **Budgets:** exact boundary and one-over for pages, bytes, axes, area, composed shapes/geometry/raster pixels; many foregrounds sharing one large background; tiny bytes with huge dimensions; repeated raster uses; partial-export failure creates no frame/URLs.
5. **Physical layout:** portrait/landscape/square/mixed-size pages, fractional inches, one page per sheet, no blank final page, no viewport/zoom dependence, crop/flip/opacity retained. Unit tests inspect generated CSS; only browser-produced output establishes pagination.
6. **Readiness:** already-loaded images, delayed load, decode rejection, unsupported decode fallback, timeout, abort and late completions; no print invocation until ready.
7. **Cleanup:** fail each allocation step, cancel before/after ready, repeat clicks, destroy while preparing, new document while decoding, beforeprint/afterprint repetition, ignored/nonblocking print stub, one-time URL revocation, zero retained frames/listeners/timers.
8. **Semantics:** current mode explicitly diagnoses unsupported saved print settings; later core fixtures cover visible/unprintable, hidden/printable, guides, intrinsic hidden, mixed memberships and group inheritance without changing source scenes.
9. **Bindings/accessibility:** all existing framework handles forward the shared method, destroyed handles reject, viewer page/zoom/selection/events stay unchanged, ready/cancel controls remain keyboard-accessible.

Existing librsvg/Cairo pixel checks can validate individual exported SVGs and collision fixtures with bounded surfaces. They do not validate browser named pages, sandbox behavior, dialog lifecycle, printer output, or Microsoft Visio parity. The current environment cannot run the required browser path; mark those tests unrun until an authorized supported browser environment is available.
