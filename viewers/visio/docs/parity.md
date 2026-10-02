# Visio capability ledger

Full Microsoft Visio parity is a target, not a current claim. Snapshot: 2026-10-02.

Implementation status is based on source inspection. Test filenames identify evidence; current pass counts and limitations are recorded in verification.md. Generated fixtures and the original demo scene do not establish visual parity with Microsoft Visio.

| Capability                     | Target                                        | Current state                                       | Evidence to inspect                                                     | Still missing                                                                        |
| ------------------------------ | --------------------------------------------- | --------------------------------------------------- | ----------------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| VSDX package intake            | Safe local OPC/ZIP reading                    | Initial implementation                              | Core package.ts; package.test.ts                                        | Broader producer corpus, independent security review                                 |
| Format discovery and pages     | Relationships, page order, dimensions         | Initial implementation                              | Core parser.ts; parser.test.ts                                          | More genuine document fixtures                                                       |
| Background pages               | Resolve and draw inherited background pages   | Partial                                             | Core getVisioPageLayers; viewer render-svg.ts                           | Visio reference comparisons; page-scale interactions                                 |
| Master inheritance             | Master and instance cell/section merging      | Partial                                             | Core shapes.ts and sheet.ts                                             | Formula-driven resizing, full master semantics                                       |
| Group transforms               | Nested coordinate systems, rotation, flips    | Partial                                             | Core geometry.ts; geometry.test.ts                                      | Complex group corpus and reference renders                                           |
| Basic paths                    | Move, line, relative quadratic/cubic segments | Partial                                             | Core geometryPaths in geometry.ts                                       | Comprehensive geometry conformance corpus                                            |
| Ellipses and arcs              | Native cached arc and ellipse geometry        | Partial                                             | Core geometry.ts; geometry.test.ts                                      | Degenerate cases and reference images                                                |
| NURBS, splines, polylines      | All Visio Geometry row forms                  | Not implemented                                     | Unsupported row diagnostics in geometry.ts                              | Evaluation and rendering of remaining row types                                      |
| Line and fill styles           | Full styles, patterns, gradients and opacity  | Partial: solids and normalized linear gradients     | Core theme/gradient tests; ten real-corpus gradients; SVG paint tests   | Radial/hatch fills, exact dash/cap behavior, complex effects                         |
| Arrowheads                     | All connector end variants                    | Partial: codes 1-4                                  | Viewer arrowheads.ts; SVG markers                                       | Exact sizing, remaining codes, reference corpus                                      |
| Connectors and glue            | Connectivity, routing, dynamic glue           | Partial read                                        | Core page connection model; cached geometry                             | Live routing, glue constraints, editing and reroute                                  |
| ShapeSheet formulas            | Recalculation and dependencies                | Cached values only                                  | Core sheet.ts diagnostics                                               | Formula evaluator, dependencies, functions, recalc                                   |
| Text                           | Runs, margins, orientation and layout         | Partial                                             | Core style.ts; viewer render-text.ts                                    | Exact wrapping, font metrics, tabs, mixed paragraphs                                 |
| Themes and styles              | Theme colors, variants and inheritance        | Partial: bounded selected theme records             | Core theme resolver; exact upstream color assertions and gradient stops | Complete old-theme variants, effects and full reference equivalence                  |
| Images and foreign data        | Raster, vector, OLE and foreign objects       | Not rendered                                        | Foreign-object diagnostics                                              | Safe decoding, intrinsic size, transforms, preview                                   |
| Layers                         | Visibility, locks, colors and print settings  | Not implemented                                     | Layer membership warning in shapes.ts                                   | Layer model and all rendering overrides                                              |
| Shape data and hyperlinks      | Properties, data graphics and links           | Partial: cached data and guarded hyperlink metadata | 389 real-corpus data rows; generated URL tests; shared inspector tests  | Internal-target navigation, external subaddresses, frames, data graphics and editing |
| Containers and swimlanes       | Container relationships and layout            | Not implemented                                     | Cached child shapes may still render                                    | Semantic model, memberships, layout and editing                                      |
| Page navigation                | Select pages and report current page          | Initial implementation                              | ViewerController and viewer-element.ts                                  | Genuine multi-page usability corpus                                                  |
| Zoom and fit                   | Bounded zoom and viewport fitting             | Initial implementation                              | ViewerController; viewer-element.ts                                     | Pinch/pan behavior and large-diagram performance                                     |
| Shape selection                | Identify a clicked shape                      | Initial implementation                              | shape-select event; SVG shape identifiers                               | Keyboard shape tree, multi-selection, accessibility audit                            |
| Framework integration          | One renderer across host frameworks           | Shared lifecycle API                                | contract.ts; binding.ts; adapters                                       | Separate build/runtime verification per framework                                    |
| Editing and history            | Move, resize, style, text, undo/redo          | Not implemented                                     | Read-only viewer contract                                               | Command system, mutations, history, bindings parity                                  |
| Native save and round-trip     | Save editable VSDX preserving unknown parts   | Not implemented                                     | No writer API                                                           | Writer, relationships, unknown-part preservation, Visio reopen                       |
| Export and print               | SVG, images, PDF and print layout             | Not implemented                                     | No export/print API                                                     | Dedicated export pipeline and fidelity tests                                         |
| Legacy and other Visio formats | VSD, VDX, stencils, templates, macros         | Not supported                                       | Input contract currently targets VSDX                                   | Format-specific parsing, validation and test corpus                                  |
| Collaboration and comments     | Shared editing with durable state             | Not implemented                                     | No collaboration transport                                              | Editing model, conflict handling, authentication, privacy                            |
| Visio visual parity            | Measured equivalence on real documents        | Not established                                     | Generated fixtures and original sample only                             | Genuine corpus, reference renders, metrics, manual review                            |

| Corner rounding | Rounded contiguous geometry | Partial: closed axis-aligned rectangles | 29 generated cases; five source-confirmed 60973 shapes | General polygon, open and mixed/curved path rounding |

## Evidence required

1. Genuine Visio-authored documents and a representative compatibility corpus.
2. Versioned Visio reference renders with fonts/page settings.
3. Measured geometry, text, connector and style comparisons.
4. End-to-end edit, save and reopen tests in Microsoft Visio.
5. Unknown XML part and relationship preservation.
6. Cross-framework/browser interaction and accessibility tests.

Warnings are useful but not comprehensive. Absence of diagnostics is not a fidelity guarantee.

## References

- [Microsoft file format reference](https://learn.microsoft.com/en-us/office/client-developer/visio/visio-file-format-reference)
- [Introduction to VSDX](https://learn.microsoft.com/en-us/office/client-developer/visio/introduction-to-the-visio-file-formatvsdx)
- [MS-VSDX](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/50c23601-c943-4ff2-b4a1-02445f52daf0)
- [Engine repository](https://github.com/ChristopherVR/ooxml), local implementation under src/visio. Unpublished changes may not be on the remote.
- [Viewer design reference](https://github.com/ChristopherVR/pptx-viewer)
