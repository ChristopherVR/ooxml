# Bounded Visio EMF conversion

`parseVsdx` remains converter-free by default. Its optional `metafileConverter` callback is a trusted package injection, not a document-supplied function. Hosts must supply it only inside a disposable worker whose parent enforces a hard deadline and cancellation. A Promise timeout within the same worker cannot interrupt synchronous converter work.

The private viewer injects the released `emf-converter` browser package in its parser worker. The UI owns a 15-second deadline for the entire document and terminates that worker on timeout, cancellation, success, or failure. The non-worker fallback never injects a converter.

## Admission and retention

- Internal image relationships only; no network loading
- Structural EMF inspection before conversion; no converter invocation for rejected input
- Conversion allowlist: header, EOF, MoveTo, LineTo, ellipse, rectangle, and admitted stock SelectObject records
- No mapping changes, clipping records, paths, custom GDI objects, text, bitmap records, comments, EMF+, WMF, SVG, or OLE conversion in this slice
- At most eight unique conversion attempts and 1 MiB aggregate input per document; repeated references reuse the validated result
- At most 256 KiB, 512 records, 4096 inspected points and 2048 by 2048 pixels per converted asset
- Output sanitizer rejects unknown tags, attributes, URLs, markup, fonts, images and non-inert styles. The retained scene contains only validated numeric geometry, paint, and bounded clips
- Conversion output ceilings: 2048 nodes, 20,000 path operands, 10,000 commands and 256 KiB characters, including bounded expansion
- Retained output bounds do not constitute a peak converter heap limit. Browser workers provide termination, not a hard memory quota. Narrow input admission and the fixed converter options constrain this exposure

Conversion failure omits the asset with a diagnostic while preserving the rest of the drawing. Successful output reports the limited primitive subset explicitly; it does not establish Microsoft Visio visual fidelity. Cached image placement, opacity, shape transforms, master inheritance and local replacement remain core-owned.

Generated tests cover admission, failure, deduplication, limits, placement and inheritance. Real Microsoft Visio rendering comparison and broader record support remain future work.
