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

## Remaining candidates and prerequisites

| Candidate in `pptx-viewer/packages/shared/src/render` | Destination                                      | Required separation                                                                               |
| ----------------------------------------------------- | ------------------------------------------------ | ------------------------------------------------------------------------------------------------- |
| `chart-histogram-binning.ts`                          | `chart`                                          | Neutral histogram options; keep label formatting out of the chart view-model dependency           |
| `chart-axis-nice.ts`                                  | `chart`                                          | Retain the PowerPoint-derived automatic-axis policy; Excel uses a different interval policy today |
| `chart-date-format.ts` and chart number formatting    | Existing `xlsx/numfmt` plus shared chart adapter | Resolve Excel serial-60 behavior and format-code contracts before consolidation                   |
| `theme-color-swatches.ts`                             | `color` for luminance/scheme descriptors         | Product-specific theme-reference serialization and popup UI remain separate                       |
| `smartart-drawing-viewbox.ts`                         | `diagram`                                        | Adapt from the PowerPoint model to shared drawing bounds, including separate text frames          |
| Remaining SmartArt family layout approximations       | Neutral core diagram model/engine first          | Avoid treating the viewer's approximate family rendering as a complete layout engine              |
| Font-picker catalogs and grouping                     | `ooxml-ui`                                       | Inject localization keys, theme/embedded fonts and product defaults                               |

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
