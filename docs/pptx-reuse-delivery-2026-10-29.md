# Remaining PowerPoint reuse capabilities, 29 October 2026

Follow-up to [the 28 October delivery](pptx-reuse-delivery-2026-10-28.md), based on
`bf7be3e5c`. Implementation and native acceptance were performed in isolated worktrees with
three GPT-6.1-sol agents. This closes the listed implementation/adoption work; it does not
assert that every Office document now has identical pixels or supports every editing operation.

## Delivered

| Shared owner                                                 | Capability                                                                                                                  | Product adoption                                                                                      |
| ------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| `core/chart/region-map`                                      | Bounded clear/compressed native geography cache, compressed polygon decoder, projections, region matching and colour scales | Shared Word/Excel/Visio chart painter and PowerPoint compatibility renderer                           |
| `core/xlsx` embedded chart APIs                              | ChartEx admission, live source hydration and source-preserving edit transactions                                            | Visio embedded workbook data and formatting panes                                                     |
| `core/drawingml`, `core/geometry`, shared UI picture painter | Ordered standard/artistic metadata, source crop, tiled/mirrored/offset fills, fill rectangle and rotation behavior          | Word images/Print Layout, Excel drawings, shared SmartArt and PowerPoint picture adapters             |
| `core/diagram`                                               | Picture role associations, original media identity, node edits and layout reflow, generated picture caches/relationships    | Word, Excel and PowerPoint, including save/reopen and switching away from and back to picture layouts |
| `core/drawingml`, `core/digest`                              | Existing validated raster metadata and CRC32 extracted from Visio                                                           | Visio and synchronously sized PowerPoint SmartArt tile descriptors                                    |
| Native acceptance scripts                                    | Same-frame Office/Chromium captures, source/version metadata, quantitative differences and contact sheets                   | Shared chart, picture and SmartArt renderers                                                          |

ChartEx edits reuse the existing worksheet transaction and source-preserving writer. Unknown
unsupported content refuses unsafe edits, while no-op edits keep original bytes. Native map
workbooks with omitted chart caches hydrate their local source ranges through the existing
formula engine. Sparse PowerPoint map data retains compact arrays and source-index identities.

Picture operations retain their order. Nonlinear HSL/tint stages use shared pixel arithmetic;
preceding and following SVG effects are applied in that order. Native processed artistic blips
are used directly, with their original WDP relationship and extension metadata preserved, so a
cached effect is not applied twice. A committed Office-generated pencil fixture proves that path.

SmartArt pictures belong to presentation roles, which can omit `presOf`; assigning them to a
content node incorrectly paints the caption in native Office. The new association code preserves
that distinction, authored node styling above role palettes, media identity across layout changes,
original frame extent, crop/effects and native picture relationships. The admitted 2D layouts
reuse the shared engine and cache writer; complex layouts keep their existing explicit fallback.

## Native evidence

[The comparison workflow](native-visual-comparison.md) records Office and Chromium versions,
source hashes and unregistered frame metrics. A threshold result is separate from a diagnostic
measurement. The native run uses PowerPoint 16.0.20527.20088. [The committed evidence](evidence/shared-native-2026-10-29.json)
records **24/24 picture controls passing** at normalized RGB error <= 0.02 and foreground
intersection over union >= 0.8. Meaningful blur controls use radii 2 and 6. Chart gallery,
cached map and four SmartArt before/after captures are recorded as measured diagnostics,
not threshold passes.

Concrete discrepancies found and corrected include:

- Cropping incorrectly reduced the tile period; Office uses the uncropped bitmap dimensions.
- Funnel bars lacked the native default 33% gap; explicit zero/custom gaps remain respected.
- Map series solid fills were ignored, absent borders/labels were fabricated, and two legends
  were painted. Source paint and geographic labels now follow their distinct contracts.
- SmartArt role colours, host frame extent and native presentation-point picture identity were
  lost during regeneration. The saved native decks now show real nonuniform image content in
  the intended image frames before and after editing.
- Alpha inversion, nested alpha modulation, colour replacement and nonlinear HSL behavior
  differed from native picture output. Native controls now exercise those semantics.

The committed chart gallery is fully compared for its 19 admitted charts; surface is explicitly
listed as unsupported. These diagnostic images expose remaining chart layout/typography
variations and are not treated as passing pixel-equivalence tests. The native cached-map probe
exercises more than half a million decoded points; its provider geometry remains local and
synthetic polygons cover the decoder in CI.

## Interpretation boundaries

The implementations and consumers above are delivered. Native measurements do not imply
universal Office parity: uncached artistic previews and tint remain declared approximations. Creating or changing an
artistic effect requires a host-generated baked image cache;
unsupported effect containers remain source-backed, and map output without cached geometry
reports its approximate fallback. SmartArt native captures establish picture role/media retention
and editing behavior for the admitted layouts, not identical geometry/typography for all layouts.
The comparison opens Office with alerts suppressed and therefore is not a no-repair assertion.
Binary codecs remain in ole2; product sessions and framework lifecycle remain in their hosts.
