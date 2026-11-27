# Native arrowhead glyphs, codes 1 to 45

`arrowheads-native.json` is read from the SVG that Visio 16 exports (`Page.Export`) for a page
drawn by `src/core/scripts/visio-arrowheads-native.ps1` in an invisible instance: one 2 inch
line per EndArrow code 1 to 45, EndArrowSize 2, LineWeight 1 pt. For each code it records the
export's marker glyph (`lendN`, in marker units of 3.52 pt, that is `visioOpenArrowExtent(2, 1/72)`),
whether the export fills it (`stroke:none`) or strokes it, and how far the export trims the 144 pt
line (the setback). Codes 8 and 35 to 38 reuse the markers of 5 and 31 to 34 in the export.

Only output was measured; no Visio binary was inspected. Run with
`powershell -File src/core/scripts/visio-arrowheads-native.ps1 -Out <folder>` and read
`arrows.svg`.
