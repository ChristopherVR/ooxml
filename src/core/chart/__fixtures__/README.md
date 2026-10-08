# Native Excel gradient raster reference

`native-gradient-raster.json` was recorded by
`scripts/record-xlsx-chart-gradient-raster.ps1` with Excel 16.0 build 20430.
The recorder owns a hidden Excel instance, creates independent square/wide
charts, saves copies, reopens them, activates each chart and exports its PNG.
The JSON retains the saved DrawingML fill, native angle, raster dimensions and
25 background sample pixels for each of 20 cases. Sample coordinates avoid
chart chrome such as gridlines. The colors are opaque red and white, at the
0%/100% endpoints, with `a:lin/@scaled="1"`.

The core regression compares normalized vector projection and paint-stop
interpolation against native pixels, allowing two RGB levels for quantization.
The browser regression paints the actual chart DOM gradient through SVG/canvas
in all six XLSX bindings and checks the same pixels. This verifies this measured
paint profile, not chart typography/layout or other gradient profiles. Generated
workbooks and PNGs remain outside the package; no Office dependency is required
for these regressions.

`native-gradient-linear-profiles.json` adds six profiles (120 independently
saved/reopened charts, 3,000 pixels): transparent endpoint pairs, opaque interior
pairs, opaque three-stop fills, translucent three-stop fills, crossed pairs and
coincident pairs. Generate each profile with the recorder's `-Profile` option
into its own output directory. The additional JSON includes raw PNG alpha;
comparisons use premultiplied color channels and test alpha separately. Opaque
samples retain the two-level tolerance; translucent premultiplied samples allow
three levels, with alpha within two levels.

The additional capture deliberately retains unresolved differences. Sixteen
translucent three-stop core cases are explicit expected failures, not passing
parity evidence (3.22 to 3.67 premultiplied levels against a limit of 3; see
the note in `gradient-raster.test.ts`). Browser coverage has a strict passing
sweep plus a separate expected-failure repro in every binding. That marker must
be removed once the renderer matches the native pixels; the tolerances must not
be raised to hide the remaining discrepancies. Samples are evaluated at pixel
centres. The wide 135-degree coincident-edge case passes since coincident
linear stops open into the measured 1/256 native ramp
(`gradient-coincident-stops.ts`).

## Rectangular path profiles

`native-gradient-path-profiles.json` contains 24 additional Excel 16.0 build
20430 captures and 600 RGBA points. Reproduce each output directory with
`scripts/record-xlsx-chart-gradient-raster.ps1 -Profile <profile>`, selecting
`path-center`, `path-corner`, `path-center-transparent` or
`path-corner-transparent`. Center styles use `TwoColorGradient(7, variant)`
and corner styles use `TwoColorGradient(5, variant)`. In these file names,
`angle-1` through `angle-4` identify the COM variant; each case's `angle` getter
is not used to derive path geometry. Saved `path="rect"` and `fillToRect`
values are the geometry evidence.

Opaque endpoint pairs match the shared sigma/gamma curve. Translucent pairs
use direct color/alpha interpolation and an independent opacity mask.
Comparisons retain the existing two-level opaque and three-level translucent
premultiplied tolerances, with alpha within two levels. These profiles do not
prove circle/shape paths, every target rectangle or all multi-stop path fills.
