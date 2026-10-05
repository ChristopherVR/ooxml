# Code-5 concave arrow evidence

This bounded slice adds one previously omitted arrowhead code. It does not establish Microsoft Visio pixel equivalence or exact arrowhead sizing.

## Primary semantics and original implementation

The [MS-VSDX BeginArrow definition](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/8f7d58be-20e0-433b-8419-59add0650616) identifies code 5 as a triangular arrowhead whose base curves inward. The [BeginArrow cell reference](https://learn.microsoft.com/en-us/office/client-developer/visio/beginarrow-cell-line-format-section) associates it with the first vertex and delegates size to BeginArrowSize.

The viewer uses an original, bounded SVG glyph with straight sides and one quadratic inward-curved base. Its control hull and half-unit stroke outline fit inside the marker viewport. The tip is anchored at the line endpoint; SVG `auto-start-reverse` provides opposite start/end directions in the shape's transformed coordinates. Paint uses the guarded line color and saved line opacity. Geometry with NoLine or transparent line pattern produces no markers. Very short lines retain their saved endpoints; markers may overlap, with no invented shortening or collision avoidance.

The pre-existing size approximation remains in use, including size range 0–6 and user-space marker dimensions. Exact Visio proportions, line-weight-dependent sizing, short-line overlap policy and native-render equivalence remain unverified. Codes 6–45 and custom code 254 remain unsupported and diagnosed.

Independent corroboration: [libvisio's pinned VSDContentCollector.cpp](https://github.com/LibreOffice/libvisio/blob/49fb9d3a9d21d4374cad782925e48c577a41f5be/src/lib/VSDContentCollector.cpp#L220) describes a concave arrow for code 5. Its implementation is secondary evidence; no source glyph or fixture bytes were copied into this repository.

## Real source regression

Apache POI revision `732120980140d5ed64b482c470e0b625cdb1ab15`, [test-data/diagram/60973.vsdx](https://github.com/apache/poi/blob/732120980140d5ed64b482c470e0b625cdb1ab15/test-data/diagram/60973.vsdx), SHA-256 `c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03`.

`visio/pages/page1.xml`, page ID 0, shapes 3, 7, 8 and 10 contain direct `BeginArrow=5` and `EndArrow=5` caches. Their Geometry section IX 0 explicitly saves NoFill=1, NoLine=0 and NoShow=0. Each starts at (0,0) and ends at (width,0), with widths 0.590551181102362, 0.8244596039807381, 0.8191605115954582 and 1.456692913385826 inches respectively. Saved shape transforms place these horizontally, obliquely and vertically. The normalized style uses black paint, opacity 1, width 0.03125 inches and arrow sizes 2.

Before this slice, shared SVG export emitted zero markers for that page. The hash-pinned regression now asserts eight markers on four two-ended connectors, retained transforms and safe export. This is source-confirmed restoration of omitted symbols; it does not rely on the embedded preview.

## Reproduction and scope

- `npm test -- src/arrowheads.test.ts src/render-svg.test.ts`: generated structural and analytic checks, including concavity, closure, finite control coordinates, stroke bounds, endpoint anchors, orientation attributes, short strokes, opacity, NoLine, size bounds and unsupported-code diagnostics.
- `VISIO_ARROW_CORPUS_DIR=/path/to/apache-poi/fixtures npm test -- src/arrowheads.test.ts`: additionally runs the hash-pinned real regression using an external `60973.vsdx`. Skips the real test when the environment variable is absent.
- After rebuilding, `npm run test:corpus -- /path/to/libvisio /path/to/poi --svg`: includes the eight-arrow export assertion alongside the complete pinned corpus checks.

JSDOM validates structure and analytical SVG contracts. Actual browser marker rasterization and fixed-version Microsoft Visio comparison remain separate, unverified work.

## Secondary-renderer limitation and browser regression

A bounded post-change comparison found the retained 60973 preview is only 87×73 pixels and its converted SVG wraps a raster image. It cannot resolve the saved arrow's curvature or exact dimensions reliably. The existing `@napi-rs/canvas` SVG decoder omitted markers in the fresh output. A minimal standalone marker probe also produced no marker pixels with `orient="auto-start-reverse"`, `auto` or `0`, so this decoder's omission is not evidence of a viewer regression.

`tests/concave-arrows.spec.ts` exercises actual Chromium SVG rasterization for live serialization and shared export. Analytic sample points require red pixels inside both arrows, white pixels in both inward-base gaps and white pixels behind the tips. These assertions distinguish absent markers, a filled triangle and reversed direction without claiming native Visio dimensions. Local execution stops before assertions because the Playwright Chromium executable is unavailable; remote browser CI must establish its result.
