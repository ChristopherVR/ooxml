# PowerPoint viewer reuse audit

Reviewed on 2026-10-03. The viewer still holds reusable algorithms as well as UI.
Move document and numerical logic into this repository; keep DOM, framework
lifecycles and editor interaction in the viewers or `ooxml-ui`. A pure function
can still describe UI policy, so absence of DOM imports alone is not sufficient.

## Extracted in this round

| Shared area | Capability                                                | Current PowerPoint consumers                            |
| ----------- | --------------------------------------------------------- | ------------------------------------------------------- |
| `geometry`  | SVG path flattening, including curves and elliptical arcs | Shape merge outlines and SmartArt extrusion             |
| `chart`     | Linear/polynomial regression and R-squared                | Trendline calculations, including the older viewer path |
| `chart`     | Quartiles and repeated-category grouping                  | Box-and-whisker charts                                  |
| `chart`     | Blank-point handling and stacked series values            | Line/area chart models                                  |
| `text`      | Script categorization and run segmentation                | Multilingual text font selection                        |
| `color`     | Existing hex/channel/opacity primitives reused            | Fill, stroke and text styling                           |

The new areas are strict, DOM-free subpaths with no dependencies on an Office
format or viewer. PowerPoint compatibility entry points forward to them. Other
formats can consume these APIs directly; extraction does not itself add Word
chart rendering, equation editing or font-slot behavior. Source paths, commits
and changes are recorded in `../PROVENANCE.md`.

## Extracted on 2026-10-06

| Shared area | Capability                                                                                                   |
| ----------- | ------------------------------------------------------------------------------------------------------------ |
| `chart`     | Pie best-fit labels and outside-label collision, radar geometry, squarified treemap, trendline equation text |
| `color`     | Readable text colour over a fill, hex to unit RGB                                                            |
| `text`      | Tab leader characters and decimal-tab anchoring                                                              |
| `geometry`  | Snap guides and grid snapping, align and distribute                                                          |

The viewer copies of the automatic axis, histogram binning and theme swatch modules, which had been duplicated rather than moved, now forward to the core.

## Remaining candidates and prerequisites

On 2026-10-07, the shared renderer moved into `src/ui/src/pptx` and DOM-free
editing helpers into `src/core/pptx/editor`. Snapshot cloning now belongs to
the core editor entry as well. Cached drawing bounds are shared through
`diagram/computeDiagramDrawingBounds`, with a PowerPoint adapter preserving
the independent text-frame fallback behavior. These bounds use authored
rectangles, without rotated geometry or effect extents.

| Candidate in `src/ui/src/pptx/render`                     | Destination                                      | Required separation                                                                                           |
| --------------------------------------------------------- | ------------------------------------------------ | ------------------------------------------------------------------------------------------------------------- |
| `chart-date-format.ts` and chart number formatting        | Existing `xlsx/numfmt` plus shared chart adapter | Resolve Excel serial-60 behavior and format-code contracts before consolidation                               |
| Remaining SmartArt family layout approximations           | Neutral core diagram model/engine first          | Avoid treating the viewer's approximate family rendering as a complete layout engine                          |
| Font-picker catalogs and grouping                         | `ooxml-ui`                                       | Inject localization keys, theme/embedded fonts and product defaults                                           |
| Cross-product history contracts and `secure-random.ts`    | Neutral editor / crypto areas                    | PowerPoint history is already in core; agree cross-product snapshots and preserve strong RNG failure behavior |
| `path-gradient-rect.ts`, `chart-number-format-pattern.ts` | `drawingml` / `xlsx/numfmt`                      | Emits SVG markup today; number formats must reconcile with `xlsx/numfmt` first                                |

OMML equation converters are already in `math`. Preset shape evaluators and
boolean operations already live in `geometry`. SmartArt interpreter compatibility
modules already delegate to the PowerPoint core; their neutralization belongs
inside OOXML rather than another viewer extraction.

## Migration checks

Keep numerical outputs and viewer public APIs stable, record extraction provenance,
and test both core algorithms and their real viewer consumers. Verify each affected
framework demo and packed ESM/CommonJS consumers after consuming the published
core. Remove implementation copies from the viewer rather than leaving parallel
versions that can drift. Do not replace differing format policies solely because
two helpers have similar names.
