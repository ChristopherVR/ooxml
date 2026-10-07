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
