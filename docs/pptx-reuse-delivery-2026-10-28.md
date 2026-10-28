# PowerPoint reuse implementation, 28 October 2026

Implementation of the ordered candidates in [the review](pptx-reuse-review-2026-10-28.md),
based on `b36e9472a`. Work is isolated from the shared checkout, with GPT-6.1-sol agents
owning chart, geometry/picture, and UI streams. This record tracks bounded capabilities,
not universal Office parity.

## Delivered boundaries

| Shared owner                      | Capability                                                                                                         | Actual consumers                                                                        |
| --------------------------------- | ------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------- |
| `core/chart`                      | Manual inner/outer plot layout, smooth series paths and authored strokes, point-label offsets and pie leader lines | Shared Word/Excel/Visio painter; PowerPoint adapters delegate reusable geometry         |
| `core/chart`                      | Calendar date stepping, date formatting and continuous date-axis placement                                         | Shared painter, Excel workbook date-system adapter, PowerPoint date helpers             |
| `core/drawingml`, `core/geometry` | Guide-driven custom geometry, path SVG projection, source-crop codec and crop rectangle math                       | Word custom shapes; shared SmartArt geometry; Excel image crop; PowerPoint crop adapter |
| `core/text`                       | Font substitution, PANOSE classification, style-name fallback and per-script font selection                        | PowerPoint span adapter; Word metric fallback and mixed-script Print Layout measurement |
| `ui/form/color-input`             | Target-aware live colour preview, restoration and commit transactions                                              | All five PowerPoint facades; Excel Format Cells fill preview in its draft               |
| `ui/fonts`                        | Catalogue grouping, availability, custom FontFace registration and owned registration scope                        | PowerPoint existing controls; Word font catalogue                                       |
| `ui/dialog`                       | Submit/dismiss session lifecycle with async completion handling                                                    | Word Zoom; Excel modal frame                                                            |
| `ui/recovery-scheduler`           | Timer cancellation and serialized recovery writes                                                                  | PowerPoint and Visio, with product snapshot/storage/encryption policies retained        |

Additional delivered boundaries:

- Shared chart data-table metrics and painting, funnel geometry, and box distributions/painter.
  Word, Excel and embedded chart hosts use the neutral painter. PowerPoint delegates geometry,
  statistics and table metrics while retaining its presentation-specific rendering descriptors.
- Ordered grayscale, luminance and alpha picture codecs and shared SVG filters. Excel reads,
  edits, paints and saves them; PowerPoint consumes the percentage/transfer helpers. Unknown
  operations are retained and reported rather than declared visually reproduced.
- Cached SmartArt pictures, internal image relationships and cropped/effected picture fill
  painting. Word and Excel provide their own media URL resolvers; external links are excluded
  from automatic media loading. Source drawing parts and relationship data survive save/reopen.
- Shared OPC hyperlink policy with explicit PowerPoint ftp/tel/file and relative navigation
  compatibility. Internal `ppaction` URLs remain presentation actions. Unsupported arbitrary
  custom schemes now follow the shared allowlist rather than the former permissive blocklist.

Typed text-body scene, extrusion, bevel and flat-text codecs also live in DrawingML.
PowerPoint delegates shared semantic reads and vector/preset helpers while retaining its
compatibility writers. Shared SmartArt reports 3D text as approximated; this is a model and
preservation capability, not a 3D rendering claim.

Excel colour previews change the dialog draft; its existing confirmation applies one workbook
undo transaction. Word font selection preserves literal text offsets and the existing metric
fallback policy. Generated fields, symbols and inline objects retain their existing mapping.
The browser font scope is additive; current PowerPoint font uploads retain page-session lifetime.

## Validation

- Strict core checks pass for all integrated feature slices. Each extraction includes focused
  unit tests and source-preservation/edit tests where appropriate.
- Six Word mixed-script Print Layout browser cases pass, including font selection, decoding
  supplementary-script offsets in core tests, and unchanged source ranges on the print surface.
- Six Excel live-fill preview/submit/undo/cancel browser cases and six Word Zoom cases pass.
- Forty PowerPoint recent-colour and embedded-font browser cases pass across all five bindings,
  including the shared core font adapter and font-part byte preservation on save.
- The 82-part chart corpus passes without exceptions. Manual-layout ablation changes from
  **0/10 to 10/10**, and funnel/box parts now render. The only unrendered ChartEx family in this
  corpus is region map. [Raw delivery measurement](evidence/chart-gap-delivery-2026-10-28.json).
  Date/smooth/table cases use focused XML fixtures because the historical corpus lacks them.
- Native Word 16 recognizes shape-extent guide variables separately from local path coordinates.
  Native Excel 16 normal open recognizes cropped pictures and grayscale/luminance values.
  These checks do not establish complete no-repair behavior or pixel fidelity for every effect.
- Native PowerPoint 16 capture confirmed omitted funnel values when labels are absent, and
  default box mean X markers. Both shared and PowerPoint paths now follow those defaults,
  including explicit overrides. Native colours/fonts/chrome remain comparison limits.

- Twelve SmartArt picture browser cases pass across six Word and six Excel bindings,
  including actual blob image decoding and save/reload. This caught and fixed Word's missing
  diagram-picture media admission.
- Twelve Word funnel/box editor and Print Layout browser cases pass across all six bindings.
- Six Word custom-shape resize/save browser cases pass across all bindings.

Integrated CI-planned verification, final chart browser cases and native evidence artifacts
are pending final integration. Source fixtures and SVG assertions establish supported
semantics; they do not alone establish pixel equivalence with Microsoft Office.

## Scope and remaining limits

The planned extraction sequence is implemented. Final integrated verification is still pending.
Region-map rendering, richer artistic effects, tiled/offset picture fills and full native visual
comparison remain open. Shared SmartArt picture painting currently uses cached drawing nodes;
this does not make every picture-bearing layout reflow equivalent to Office.

Product orchestration, framework lifecycle and clipboard payload formats remain product-owned.
The review's optional PowerPoint clipboard-controller consolidation is separate from the
cross-product extraction sequence.
