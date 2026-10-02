# Bounded open connector rounding

## Primary semantics

Microsoft's [Rounding cell reference](https://learn.microsoft.com/en-us/office/client-developer/visio/rounding-cell-line-format-section)
and [MS-VSDX 2.4.4.298](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/378bcccc-bf84-46e8-a63a-0f9881f84ac2)
define a saved radius in internal inches, applied at contiguous path segments.
For perpendicular line segments with sufficient room, a tangent circular quarter
arc has that radius and tangent distances equal to that radius. The implementation
uses this geometric construction without changing cached endpoints or transforms.

Microsoft's [Cutting Corners](https://techcommunity.microsoft.com/blog/microsoft_365blog/cutting-corners/237250)
explains that close vertices reduce rounding and coincident vertices can suppress
it. It does not establish an exact short-segment radius allocation algorithm.
Consequently this extension does **not** merge repeated vertices or guess clamping.
Native Visio output for the short-segment cases remains required before expanding
this support. The pre-existing rectangle clamp is unchanged by this extension.

## Admission and preservation

Each visible Geometry section is considered independently. It must contain one
MoveTo/RelMoveTo followed only by LineTo/RelLineTo, have finite valid cached
coordinates, remain open, and have strictly axis-aligned nonzero segments. Forward
collinear vertices are preserved; reversals and duplicate vertices are declined.
Every right-angle corner consumes one radius of each adjacent source segment.
No segment may be shorter than the sum of its two endpoint corner demands.
Subprecision radii and serialized tangent collapse are declined.

Unsupported sections retain their original path and rounding diagnostic. Curves,
multiple subpaths, diagonal segments and other closed polygons are not rewritten.
NoFill/NoLine/NoShow behavior, source transforms, endpoint cells, master inheritance,
and relative-coordinate normalization remain on the existing paths. Each extra
arc is charged against the package geometry-command limit. Work and memory are
linear in the already budgeted source rows; there is no quadratic corner search.

## Hash-pinned real corpus

External Apache POI revision `732120980140d5ed64b482c470e0b625cdb1ab15`,
`test-data/diagram/60973.vsdx`, SHA-256
`c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03`.
No fixture bytes are redistributed. Run the optional test with
`VISIO_THEME_CORPUS_DIR` pointing to the external fixtures directory.

Masters 28 and 36 (master26.xml and master31.xml, root shape 5) save Rounding
0.07874015748031496 inches (2 mm). Page ID 4 shapes 814, 830, 835, 857 and page ID 7
shape 293 now contain 14 circular corner arcs. The corpus test checks all five
shapes, source endpoints, inherited radius, fill/stroke, and diagnostics. Shape 830
has an independently computed two-arc tangent regression: the intermediate segment
0.15890284960748251 inches exceeds twice the saved radius 0.15748031496062992.
Page ID 4 shape 825 and page ID 7 shape 149 have insufficient space and stay sharp
with an explicit unsupported-corner-rounding diagnostic. They are not counted as
rendering improvements. Other straight or near-duplicate connectors remain outside
the improvement count.

Generated analytic tests check tangent orthogonality, radius, turn sweep,
reflection, translation, traversal reversal, exact fit, endpoint preservation,
relative coordinates, malformed/nonfinite/extreme values, multiple sections,
curves, command budgets and unchanged shape metadata. These are geometric and
source-normalization evidence, not native Visio pixel equivalence. SVG/raster
smoke checks, when performed downstream, are secondary renderer evidence only.
