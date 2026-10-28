# Further PowerPoint logic sharing review

Status: implementation and acceptance evidence are recorded in [the delivery ledger](pptx-reuse-delivery-2026-10-28.md). The findings below describe the reviewed baseline.

Reviewed 2026-10-28 against `16edc2fd9`, after the six priorities in
[Shared Office parity delivery](office-shared-parity-delivery.md). Three GPT-6.1-sol agents
independently examined core semantics, shared UI and chart/SmartArt rendering in detached
worktrees. Findings were reconciled against source and a fresh chart-gap harness run.

This is an extraction/adoption review, not an implementation milestone or a native fidelity test.
The main checkout's uncommitted presentation-timing work was excluded from the snapshot.

## Recommendation

Prioritize visible cross-product improvements over moving files for its own sake. Much of
PowerPoint's logic already lives under `src/core/pptx` and `src/ui/src/pptx`: the remaining task
is to make selected capabilities format-neutral and actually adopt them in another product.

The best next batch is **chart plot/label geometry, custom shape geometry, font policy and live
colour transactions**. Follow with picture semantics and the remaining chart rendering families.
Do not restart the delivered chart/SmartArt codec, editing or gallery migrations.

| Order | Candidate                                                     | Destination                                             | First useful consumers                                   | Nature / relative risk                               |
| ----- | ------------------------------------------------------------- | ------------------------------------------------------- | -------------------------------------------------------- | ---------------------------------------------------- |
| 1     | Plot-area manual layout, smooth lines and series strokes      | `core/chart/render` and shared chart geometry           | Word, Excel, embedded Visio, then PPTX delegation        | Existing capability adoption / medium                |
| 2     | Guide-driven custom shape paths                               | `core/drawingml` using `core/geometry`                  | Word shapes; Word/Excel SmartArt                         | Shared-reader and host-model extension / medium-high |
| 3a    | Font substitution and mixed-script font selection             | `core/text`, format adapters                            | Word layout/print and PPTX; other painters incrementally | Consolidation and layout correctness / medium-high   |
| 3b    | Live colour preview transactions                              | Shared headless `ui` helper                             | PPTX plus Excel cell fill or Word shape fill             | Small reusable controller / medium                   |
| 4     | Date axes, data tables, pie-label placement, funnel rendering | `core/chart/render`                                     | Word, Excel, embedded Visio and PPTX                     | Separate renderer slices / medium-high               |
| 5     | Picture crop/effects and SmartArt picture painting            | `core/geometry`, `core/drawingml`, shared `ui/smartart` | Excel/Word pictures and diagrams                         | Model, codec and painter extensions / medium-high    |
| 6     | Font browser services, modal sessions and recovery scheduling | Shared `ui`; existing storage and policy contracts      | Word/Excel/PPTX, then Visio as applicable                | Incremental UI adoption / medium; recovery high      |

The ordering reflects identifiable second consumers, existing neutral foundations and bounded
acceptance tests. It is not a time estimate. The UI transaction slice can run alongside the core
work; shared font policy should precede changing layout measurement or browser font services.

## 1. Chart rendering: reuse geometry without transplanting the PPTX model

The neutral chart model already carries substantially more than the shared painter draws.
The following are renderer/view-model gaps, not reasons to create another chart parser.

| Slice                                    | PowerPoint implementation                                                                | Shared gap / existing foundation                                                                                                                                                          | First acceptance gate                                                                                                               |
| ---------------------------------------- | ---------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| Plot-area manual layout                  | `src/ui/src/pptx/render/chart-view-model-layout.ts`                                      | Seven populated `plotArea/layout/manualLayout` fixtures produce identical SVG when removed. Reuse `chart/manual-layout` and `resolveManualLayoutRect`; title/legend layout already works. | Inner/outer layout targets, edge/factor modes, explicit x/y/w/h, small viewport and native fixture coordinates.                     |
| Smooth lines and authored strokes        | `src/ui/src/pptx/render/chart-line-path.ts:19,60`; `chart-cartesian-plots.ts:174`        | `src/core/chart/render/chart-svg-cartesian.ts:236` emits straight polylines with fixed width. Models retain smoothing and paint.                                                          | Smooth and straight series together, null gaps, clipping, short runs and stable point identity.                                     |
| Calendar/date-axis spacing               | `src/ui/src/pptx/render/chart-date-axis.ts:92`; `chart-horizontal-axis.ts:58`            | Neutral axes retain date facts, but `chart-svg-cartesian.ts:71` bands category indices equally.                                                                                           | Irregular dates, monthly ticks, explicit bounds, reversal and 1904 date system.                                                     |
| Chart data tables                        | `src/ui/src/pptx/render/chart-data-table-metrics.ts:33`; `chart-data-table-render.ts:54` | `src/core/chart/model.ts:168` retains the table; it is not represented in the shared rendering view.                                                                                      | Plot/table size reservation, keys/borders, formatted values, long names, and category-label suppression only when appropriate.      |
| Point-label offsets and pie leader lines | `src/ui/src/pptx/render/chart-pie-labels.ts:103,130,168`                                 | Shared `chart-svg-labels.ts:28` is simpler. Collision, best-fit, label text and manual-layout helpers already exist in core.                                                              | One manual offset plus outside pie labels, deleted points, stable source indices and leader lines. Rich callouts follow separately. |
| Funnel, then box-and-whisker             | `src/ui/src/pptx/render/chart-funnel-sunburst.ts:100,159`; `chart-box-whisker.ts:65`     | `src/core/chart/render/summary-from-chartex.ts:1` excludes these families. ChartEx editing and box statistics are already shared.                                                         | Funnel values/colours/labels, explicit zero/negative-value semantics, native comparison; box rendering is a separate follow-up.     |

Keep DOM nodes, browser interaction and three.js in `src/ui`. Put numerical geometry and the
existing SVG-string painter's logic in core. Introduce neutral render descriptors, then make PPTX
delegate while retaining its public compatibility model. Word and Excel should demonstrate the
same slice before declaring it shared; Visio gains it through embedded workbook charts, not by
pretending all native Visio objects are OOXML charts.

### Fresh evidence and limits

Ran the existing `chart-renderer-gap.test.ts` at this snapshot: **2 tests passed, 82 parts, no
rendering exceptions**. The saved [measurement](evidence/chart-gap-2026-10-28.json) contains:

- 75 classic and 7 ChartEx parts. Funnel, box-and-whisker and region map each have one part
  classified as not drawn. Nonempty SVG elsewhere can include placeholders, including surfaces.
- Removing basic labels changes 27/27 relevant SVGs; markers 6/6; trendlines 1/1. These shared
  features should not be proposed again as wholly absent.
- Manual layout changes 0/10 SVGs: seven populated plot-area layouts and three populated
  point-label layouts. There are **no title or legend manual-layout examples** in this set.
- The seven plot layouts are in `bar-picture-fill-hill.pptx` and charts 1-3 of both
  `issue-132-gradient-fill.pptx` and `issue-132-hr-deck.pptx`; the three point offsets are in
  `chart-label-callouts.pptx` charts 1-3. These are existing starting fixtures.
- The detector finds no smooth-line, date-axis or data-table examples in this corpus. Add focused
  fixtures before porting those slices. Zero coverage is not proof of absence or correctness.

Feature detection uses regular expressions and only some features have ablations. An SVG change
shows sensitivity to removed XML, not correct rendering. This run is neither a native image
comparison nor a complete Office feature census. Do not carry forward the historical report's
"78 of 82" as a current rendering score.

## 2. Custom geometry: a concrete Word capability unlocked by shared code

PowerPoint already evaluates structured custom paths in
`src/core/pptx/core/geometry/custom-geometry-parser.ts:88,114`,
`guide-formula-paths.ts:30` and `custom-geometry-guides.ts:49`.
Shared formula evaluation and arc helpers already exist in `core/geometry`; extract the remaining
path/model adapters, not another formula evaluator.

The neutral reader, `src/core/drawingml/drawing-geometry.ts:102,111`, handles literal paths but
reports `CUSTOM_GEOMETRY_GUIDES_UNSUPPORTED` for guide coordinates. Word's
`src/core/docx/shape-parse.ts:104` rejects all `custGeom`. Cached SmartArt also consumes the shared
geometry reader (`src/core/diagram/drawing.ts:81`).

Extend `parseDrawingGeometry` with neutral guide/path inputs and outputs, then extend Word's shape
model and painter. The first slice should render one guide-driven Word shape at two sizes while
preserving ordered commands, arcs, multiple path coordinate spaces and unrelated XML on save.
Word/Excel SmartArt can then adopt it. General Excel shape support is separate host admission work.

## 3. Fonts: distinguish semantic policy from browser resources

### Core font policy and script selection

There are separate substitution maps in `src/core/pptx/core/utils/font-substitution.ts:56,402`,
`src/core/docx/layout/fonts.ts:6,23` and `src/core/text/css-font-family.ts:2,12`.
Extend the existing `text` API with canonical lookup, quoting, optional PANOSE classification and
explicit fallback policy. Preserve current product chains initially: changing a fallback can
change Word pagination and chart measurements.

Next, neutralize `src/ui/src/pptx/render/text-script-fonts.ts:81,131` into run segmentation and
font-slot selection in `core/text`, using the existing Unicode script detector. Word retains
East Asian/complex-script faces (`docx/model-run.ts:68,70`) but its layout adapter chooses the
Latin family (`docx/layout/adapter.ts:115-120`). Prove one mixed Latin/CJK/Arabic Word run through
measurement, printing and unchanged authored font slots on save. Include punctuation, combining
marks, bidi and symbol-font cases; script detection is not a shaping engine.

### Shared browser font services

`src/ui/src/pptx/render/font-catalog.ts:37,70,122`, `font-availability.ts:12,28` and
`custom-fonts.ts:131,176` provide grouped theme/embedded/custom fonts and browser registration.
Word's `ribbon-home.ts:95` uses a static list; Excel's `ribbon/controls-input.ts:31,121` has its
own combo/preview path. Share generic catalogue descriptors and a scoped FontFaceSet service in
UI, with product-provided labels and available font sources.

Start with PPTX and Word catalogue/availability, then session-only uploads with explicit disposal.
Test failed loads, SSR without `document`, two documents and registration ownership. Embedded-font
package discovery, licensing, consent and format-specific trust remain host adapters.

## 4. Live colour transactions: a small shared UI win

`src/ui/src/pptx/render/live-color-input.ts:45,55,108` and `live-color-leaves.ts:95,159` are generic
over target and patch types. They already handle preview, one settled undo step, Escape/no-change,
target changes and preserving concurrent edits to unrelated properties.

Word's `ribbon-color-grid.ts:41`, Excel's `ribbon/color-grid.ts:109` and Visio's
`viewer-color-menus.ts:82` share the colour-grid element, but not this transaction protocol.
Move the generic controller into shared UI and keep patch construction, undo and target lookup
in product adapters. Prove Excel cell fill or Word shape fill plus all five PPTX bindings: N input
events become one undo, cancellation restores theme references, and a concurrent unrelated edit
survives. Word text formatting needs a selection-range adapter and should not be the first slice.

Do not build another colour picker or MRU store. `core/color/theme-palette.ts:165` already exports
`pushRecentColor`; adapt PPTX's `recent-colors.ts:90` policy with explicit casing/limit choices.

## 5. Pictures and SmartArt: codecs, geometry and painting are separate steps

PowerPoint picture codecs remain in runtime `image-color-effects.ts:80,181`,
`image-alpha-effects.ts:93,151`, `image-a14-effects.ts:269` and
`image-a14-effects-writer.ts:168,196`. Shared picture fills carry relationship/source data, while
Excel's `src/core/xlsx/read/drawing.ts:103-121` does not expose those effect semantics.
Word explicitly reports crop/artistic-effect limitations in `docx/parse-warnings.ts:113`.

Move typed picture colour/alpha effect semantics to DrawingML while preserving effect order and
unknown `blip` children. A bounded first slice is grayscale, brightness/contrast and alpha in PPTX
and Excel read/save, followed by shared painting. Existing typed shape effects are different and
should not be re-extracted. Artistic effect fidelity remains separate work.

Crop rectangle/inset maths in `src/ui/src/pptx/render/picture-crop/crop-geometry.ts:65,83,140,181,221,238`
can use neutral `geometry` inputs; source-crop codecs belong in DrawingML. Start with Excel crop
read/render/save before interactive handles. Preserve rotation, flips and negative margins;
UI minimum frame sizes are not document semantics. There is not yet a second implemented consumer,
so this is capability-enabling work, not merely removing duplication.

For SmartArt, `src/ui/src/smartart/smartart-paint.ts:17` reports picture/pattern fills as approximated.
PPTX's `smartart-drawing.ts:133` already projects resolved picture bodies. Add a shared image-resolver
contract and painter support for Word/Excel nodes, with host-supplied URLs and relationship policy.
Test crop, unresolved images, multiple diagrams/paint IDs and save/reopen. Pattern painting is still
a gap in the PPTX path too; do not describe it as an existing implementation ready to move.

## 6. Follow-on consolidation with clear ownership

| Candidate                 | Evidence and existing building block                                                                                                                                         | Boundary / first proof                                                                                                                                                                                                                   |
| ------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Modal session helper      | PPTX `web-components/insert-dialog-frame.ts:44`, Excel `dialogs/frame.ts:55`, Word `dialog-shell.ts:17,44`; shared `office-ui-dialog` already owns focus trapping and Escape | Shared typed result/async-submit/disposal helper, not another dialog. One Word and Excel dialog: exactly-once completion, late result after dismissal inert, shadow-root focus return, second ownerDocument.                             |
| Hyperlink policy adoption | PPTX `render/hyperlink-security.ts:33,78`; existing `core/opc/safe-href.ts:82,100,105` already serves Excel                                                                  | Reconcile deliberate relative/ftp/tel rules and inject host opening policy. Preserve internal `ppaction`/slide actions. Security-sensitive compatibility change, not a generic allowlist replacement.                                    |
| Recovery scheduling       | PPTX `render/controller/autosave-controller.ts:73,84`; Visio `viewer-autorecover.ts:93,106,131,224`; existing `core/embedding` storage contracts                             | Share timers, in-flight/dirty-generation scheduling and cancellation in UI. Keep payloads, serialization, encryption and storage namespaces product-owned. Prove edit-during-write, retry and disposal with fake clocks before adoption. |
| Typed text-body 3D codecs | PPTX `text-body-scene3d.ts:72,171`, `text-body-sp3d.ts:63`; neutral `drawingml/text-model.ts:166` retains opaque XML                                                         | Later DrawingML codec slice with source preservation and PPTX adoption. No 3D rendering promise; lower value than current geometry/font/picture consumers.                                                                               |

## What should stay product- or framework-specific

- Slide/master/layout traversal, placeholder inheritance, slide packaging and presentation actions.
- Animation/media timing, morph, transitions and presenter-window orchestration. Consolidating these
  into `ui/pptx` benefits five bindings but is not automatically reusable Office logic.
- Framework lifecycle, stores/runes/change detection, portals and focus bridges.
- Product selection, document placement, history transactions, sheet formula resolution, relationship
  allocation and public compatibility projections. They should consume shared semantics, not disappear.
- Clipboard payload formats. PPTX cut/paste orchestration is repeated in React `useClipboardHandlers.ts`,
  Vue `useElementClipboard.ts`, Svelte `editor-clipboard.ts` and vanilla `editor-clipboard-actions.ts`.
  Consolidate that into a PPTX controller across all five bindings; Excel ranges and Visio masters
  need their own codecs. A shared clipboard transport is a smaller, separate opportunity.

The generic pointer-gesture, colour-grid, dialog, chart-format and SmartArt gallery foundations already
exist. Do not copy the whole PPTX gesture/editor controller into a neutral area: its snap, connector,
selection and template-routing semantics need explicit host contracts.

## Implementation gate for each future slice

1. Name the current shared API, missing neutral field/operation and actual second consumer.
2. Move or adapt one coherent feature; record source path/commit and changes in `PROVENANCE.md`.
3. Compile new shared core code strictly. Core must not acquire UI, DOM, framework or PPTX imports.
4. Preserve PPTX behaviour with its existing geometry/codec tests; cover preservation, unsupported
   diagnostics and edit/save/reopen where semantics change.
5. Demonstrate the second consumer. Binding changes reach all five PPTX bindings; product controls
   use shared elements and every locale. Rendering claims need native fixtures, not just equal SVGs.
6. Only remove superseded source paths after their admitted cases pass. Run the CI planner's verify
   from an isolated worktree before any later push.

## Documentation findings

The October 17 table in [the renderer gap report](chart-renderer-gap-report.md) is historical.
This review adds fresh measurements without treating uncovered features as failures. The old
header comment in `src/core/pptx/core/utils/chart-from-neutral.ts:9` still lists gradients,
trendlines, error bars, layouts and protection as absent from the neutral model; that is stale
against the current models and the delivered adapters. Update it when touching that adapter.

The delivered operations remain delivered. Remaining source adapters and richer painting do not
invalidate the prior bounded acceptance record, and that record does not establish universal
lossless editing or native visual equivalence.
