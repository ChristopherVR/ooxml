# Native Office visual comparison

The shared renderers have a reproducible comparison against installed Microsoft PowerPoint.
This is a manual Windows acceptance tool; CI uses deterministic synthetic unit and browser
fixtures. It neither downloads Office nor depends on a remote rendering service.

## Capture contract

A manifest names a local presentation, one-based slide/shape indices, a shared SVG, and the
output width and height. Each case may override the root presentation. The capture script
exports the full native slide at the required density and records the authored shape frame.
The comparison crops that frame without translating, registering or content-cropping the image.
Both renderers therefore use the same viewport. Cropped overflow is not measured.

Native metadata includes source SHA-256, Office version, executable version and shape dimensions.
The shared capture records the Chromium version and SVG SHA-256; comparison refuses stale source hashes. HTTP(S) loads are blocked; fixture images must
be local data URIs. Fonts come from the local machine. Close PowerPoint before running capture;
the tool creates and closes its own COM instance. The presentation opens read-only with alerts
suppressed. Successful visual capture is not evidence that Office opened without repair.

Metrics compare RGBA composited over white: normalized RGB mean absolute error, fraction of
pixels with a channel difference over 16, and foreground-mask intersection over union (any
channel below 245). Blank images and mismatched dimensions are invalid. Cases with no declared
threshold are **measured**, never passed. Difference images amplify error four times for inspection.
Thresholds belong to a named fixture/control family, not a universal Office-parity assertion.

## Run

Run commands from the indicated package so Bun resolves declared dependencies. Replace the
example output directories with local absolute paths. No generated provider cache is committed.

```powershell
# From src/core: generate shared SVGs and a manifest for the committed chart gallery.
bun scripts/native-visual/charts.ts D:/Temp/ooxml-native/charts
# From the repository root: export native slide images and capture metadata.
powershell -NoProfile -File src/core/scripts/native-visual/capture.ps1 -Manifest D:/Temp/ooxml-native/charts/manifest.json
# From e2e/pptx: rasterize shared SVGs in Chromium at 1x.
bun native-visual/rasterize.ts D:/Temp/ooxml-native/charts/manifest.json
# From src/core: compare and write comparison.json plus three-panel PNGs.
bun scripts/native-visual/compare.ts D:/Temp/ooxml-native/charts/manifest.json
```

The chart generator uses the document theme, direct chart formatting and linked classic-chart
style/color parts. It lists unsupported families explicitly rather than silently omitting them.
The committed gallery contains 19 admitted charts and one unsupported surface chart. The gallery
is a diagnostic matrix, not a pixel-equivalence test with relaxed pass thresholds.

Additional generators:

- `src/ui/scripts/picture-native-fixture.ts`: ordered effects, tile/offset/crop/rotation and a
  native baked artistic source when supplied. It writes the same manifest case structure.
- `src/core/scripts/chart-region-map-native-fixture.ts`: a local native geography cache paired
  with a shared map SVG. Provider geometry is not checked into this repository.
- `src/core/scripts/generate-picture-reflow-native.ts`: picture-layout before/after decks and
  neutral drawings; `src/ui/scripts/native-smartart.ts` paints the drawings using the actual
  Word/Excel shared SVG renderer.

Keep generated images and source packages beside their manifests. Only the bounded evidence
summary is committed. Source-model/edit/round-trip tests and real binding browser tests remain
necessary: a close screenshot cannot establish data preservation or editing correctness.
