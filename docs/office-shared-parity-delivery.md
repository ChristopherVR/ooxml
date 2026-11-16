# Shared Office parity delivery

Implementation started 2026-10-26 from `27cc953db`. This is the active delivery checklist for the
six priorities identified by the roadmap review. Historical wave notes in the other plans remain
evidence of what shipped at the time, not the current capability ledger.

## Completion rules

Each feature is tracked independently for reading/preservation, rendering, editing, save/reopen,
product integration and native Office validation. A shared implementation alone does not establish
consumer parity. A comparison between the old and new implementation does not establish native
Office fidelity. Tests must cover the requested change and preservation of unrelated content.

Core owns document semantics, layout, geometry and serialization. Shared DOM painters and controls
belong in `ooxml-ui`; package relationships, host placement and framework lifecycle remain adapters.
Every changed PowerPoint control must reach all five bindings.

A ticked box below means the admitted subset named in that item shipped with its tests; it does
not mean Office parity for the feature. Open gaps in the shared layer (3-D chart views outside
PowerPoint, region map ChartEx, Visio ChartEx, SmartArt and equations, equations in Excel) are
tracked in the [Office parity checklist](parity-checklist.md).

Equation ownership audit, 2026-10-31: the 11 non-test `latex-*` and `omml-*` files under
`src/ui/src/pptx/render` contain only compatibility re-exports from `ooxml-core/math`, including
the command, construct, environment, symbol, sibling, colour and layout helpers. No copied
converter remains to delete. The eight core math test files passed (192 tests), as did the
four PowerPoint facade test files (170 tests). This checks shared ownership and regression
behavior, not native Office equation rendering parity.

## 1. Safe chart editing and preservation

- [x] One shared source-backed chart writer and edit-eligibility policy.
- [x] Word preserves carried chart children and refuses edits the model cannot preserve.
- [x] Embedded workbook charts, including Visio consumers, follow the same policy.
- [x] Excel consumes the shared policy without losing formula-only source patching.
- [x] Tests cover carried children, unknown nested content, unreadable values, no-op bytes and
      supported title/data/format edits. The UI explains read-only admission.
- [x] Native acceptance verifies supported edited documents and embedded objects.

## 2. SmartArt layout and rendering in Word and Excel

- [x] Retain full layout definitions and context in neutral diagram host models.
- [x] Use the shared layout interpreter for inserted, edited and uncached diagrams; diagnose
      unsupported layouts and retain an explicit approximate fallback.
- [x] Reuse preset and custom geometry, including arcs, in the shared SmartArt painter.
- [x] Render gradients, transparency, rich text and supported shadows through shared primitives.
- [x] Corpus tests compare cached/uncached/inserted/edited geometry; product browser tests cover
      insertion, text edits, layout changes, history, resize and save/reopen.
- [x] Record remaining 3D, typography, effect and layout limits explicitly.

## 3. Shared DrawingML formatting

- [x] Complete line semantics beyond width/fill/dash/cap with preservation-aware readers/writers.
- [x] Share typed effect parsing, writing and resolution, preserving unknown effect children.
- [x] Parse and write theme format schemes, style references and placeholder-colour substitution.
- [x] Adopt the APIs in actual consumers; preserve authored theme tokens through unrelated edits.
- [x] Strict typechecks and parser/writer/resolver parity tests cover real producer fixtures.

## 4. Neutral PowerPoint chart saving and adapters

Implemented for the admitted families below. Shared operations own chart semantics; PowerPoint
projects the neutral models into its public compatibility model and supplies package relationships,
theme context and media allocation. An unclaimed edit still follows an explicit compatibility path;
the remaining source adapters and unsupported families are follow-up work, not a neutral-only or
complete Office-parity claim.

- [x] Data-label visibility/content, per-series and per-point overrides, value-from-cells labels.
- [x] Series/point fills, lines, gradients and transparency, preserving unrelated theme tokens.
- [x] Rich title runs, legend entries, axis titles/fonts, gridlines and number/date/display units.
- [x] Series structure, admitted type switches/combinations, trendlines, error bars and markers.
- [x] Layouts, 3D/chrome, data tables, stock lines/bars, protection and supported extensions.
- [x] Neutral ChartEx reading, editing, chart-part generation and admitted family switching.
- [x] Project neutral chart facts and formatting through the PowerPoint compatibility adapter,
      refreshing that adapter after saves. Retain source paths for explicitly unsupported cases.
- [x] Run the XML/model corpus harness with every difference classified, then native reopen and
      embedded-workbook refresh checks. Do not normalize away authored semantics to match a baseline.

These checks cover supported source semantics and operations, not native pixel fidelity. XY
(Scatter) series in vertical 2-D category-axis combinations are admitted (see the 2026-11-16
evidence below); 3-D and bubble series types remain whole-chart switches, as in Office. Combinations
on horizontal bars or on two value axes, raw geographic/Pareto creation, ChartEx custom label text
or manual layout (no schema node, see below), and full removal of the PowerPoint source-tree model
remain outside this delivery.

## 5. Shared SmartArt style and colour authoring

- [x] Neutral style/colour operations and definition writers preserve unknown XML.
- [x] Share cached-drawing generation independently of native Office cache regeneration.
- [x] Word and Excel mount shared Styles and Change Colors galleries with history and persistence.
- [x] Cover per-node overrides, rich text, host theme changes and save/reopen.
- [x] Keep Visio diagram conversion a separate host feature with explicit native-shape or embedded
      semantics; sharing diagram primitives is not native Visio SmartArt support.
- [x] Offer Office's full built-in catalogue (38 colour, 14 quick-style definitions, nine 3-D) in
      every product's galleries; a pick writes Office's own definition and `uniqueId`.
- [x] Write quick-style 3D (label `scene3d`, `sp3d`, text 3D) into Word and Excel generated caches
      as Office's own caches carry it; keep the spatial camera in the quick-style part.
- [x] Name the galleries in every product locale in Word and Excel, from PowerPoint's translations
      (agent-authored; not verified against Office's localized gallery names).
- [x] Draw 3D in the shared painter, approximately: lit bevels, extrusion along the camera's depth
      and the Scene styles' spatial camera (`smartart-3d.ts`, `diagram/scene-camera.ts`).
- [x] Show a 3-D pick in PowerPoint's editor before reopening (its 3D; colours and effects stay
      at the nearest intensity until reopen).
- [x] Regenerate a source cache Office laid out after an edit (round-trip writer), bound to
      Office's presentation points; placement is the layout engine's, measured below.

## 6. Shared chart controls and ChartEx consumers

- [x] Add shared formatting intents and controls for effects, 3D, text, pattern fills,
      admitted secondary axes, trendlines, error bars, markers and axis titles.
- [x] Wire the supported controls in PowerPoint, Word, Excel and embedded Visio charts.
- [x] Enable admitted Word/Excel ChartEx data-sheet, family-option and structural edits.
- [x] Synchronize workbook references/caches with safe external-source and hierarchy policies.
- [x] Verify applicable bindings, history and save/reopen, including native Word/Excel ChartEx cases.

These checks cover the admitted controls, not every Office formatting option. Secondary-axis
authoring keeps the first flattened series as the primary rendering anchor. A stacked or
percent-stacked group can give up its first or last series (each axis keeps one stack), and a
series moved back rejoins its neighbouring group and drops an unused secondary pair; a series inside
a stack, horizontal combinations and unknown group-extension splits are refused. ChartEx secondary
axes stay disabled. Each trendline of a series has its own pane section, with Add trendline. Custom
error bars take literal values or a single-row or single-column range: Excel resolves ranges from
the workbook, Word and embedded Visio only from cells the chart
already caches, and an unknown cell refuses the edit. Picture fill relationships are allocated by
one shared helper (`chart/picture-parts.ts`, over `opc/package-image.ts`) for every host's media
folder. The shared painter draws picture fills (stretched, tiled, stacked along bars, with
transparency) from `ChartSpace.pictureSources`; a Format pane picture is filed under a pending id
and stored as a media part on save, so Word, Excel and Visio (whose chart is an embedded workbook,
pictures in its `xl/media/`) offer Picture or texture fill; tiles mirror, `stackScale` stacks
by value-axis units and `a:srcRect` crops. PowerPoint's pane edits several trendlines and error
bars (amount, direction, caps and custom values, kept as literals) in all five bindings. ChartEx exposes
its own supported controls and refuses classic-only advanced requests.

- [x] Broaden classic secondary-axis splits to stacked ends, with merge-back and pair removal.
- [x] Edit several trendlines per series and custom error-bar sources (values and ranges).
- [x] Share picture fill relationship allocation across the pptx, docx, xlsx and visio packages.
- [x] Paint chart picture fills in the shared renderer, then admit them in Word, Excel and Visio.
- [x] Custom error sources in PowerPoint's own chart pane source, and native reopen evidence for
      the new secondary, trendline and custom-source saves (see the 2026-11-16 entry below).
- [ ] PowerPoint's save of a stacked group's end series moved to the secondary axis (natively read
      as primary).

## Integration and evidence

- [x] Reconcile `roadmap.md`, `agnostic-core-plan.md`, `structure-plan.md`, product capability
      pages and PowerPoint ribbon reviews with the delivered implementation.
- [x] Run the repository CI planner's verification from the isolated integration worktree and
      affected browser suites; retain native version/font/capture details where fidelity is claimed.

Implementation and validation evidence is recorded below as each increment is integrated. No
milestone is complete merely because its foundations or first supported cases have landed.

### Chart preservation, 2026-10-26

Integrated shared `writeChartSpaceOverSource`, Word edit/save guards, embedded workbook admission
and the Visio read-only dialog. Unsupported nested elements and unreadable values cannot be lost
through an edit; supported top-level passthrough nodes survive, and unchanged parts retain bytes.
The Visio diagnostic is localized in all five locales.

Validation: 21 focused core tests, the Visio UI cases and locale checks, strict core/full UI
typechecks and a core build passed. Native Word and Excel acceptance scripts reopened a title edit
with a retained `c:clrMapOvr` without repair; Word's chart values and embedded workbook cells agreed,
and only the chart part changed. Visio 16 reopened and resaved the embedded workbook object while
retaining that mapping and the edited title. OLE activation and visual preview fidelity were not
measured by this check. A synthetic protection probe was inaccessible in Word before core loading,
so it is unit preservation evidence only, not native acceptance evidence.

### First neutral chart save families, 2026-10-26

PowerPoint now maps supported data-label show/content/position/format edits and solid series colour
changes to shared operations. The explicit Show Labels command now requests group-level values
even when existing series already show labels. The committed corpus compares 72/72 label edits and
75/75 series colour edits identically in canonical XML and reloaded models. The three focused SDK,
save-plan and corpus files passed 479 tests. These are implementation comparisons; hidden-label
policy, value-from-cells edits and the remaining chart families are still tracked above.

### DrawingML and SmartArt foundations, 2026-10-26

Integrated shared outlines, effects, theme format schemes and source-preserving pattern/picture fill
writers, with real Word/Excel/PowerPoint producer fixtures. DrawingML's broad affected suite passed
13,355 tests (16 expected failures); final focused checks passed 84 DrawingML and 109 fingerprint
compatibility cases. Both compiler projects passed. These are codec/preservation checks, not native
rendering-fidelity claims.

Word and Excel now retain SmartArt layout/style context and use the shared interpreter and richer
SVG painter. The built-in corpus interprets 174/176 definitions and explicitly falls back for two.
Validation passed 923 core and 46 UI tests, both core/UI compiler projects and builds, plus Word's
10 and Excel's 6 browser cases across six frameworks. Unsupported connector routing, presentation
variables, typography, effects, picture/pattern rendering and flat 3D remain reported limitations.

At this foundation checkpoint, the integrated code plus ChartEx storage passed 1,085 tests across
145 files and both core compiler projects. The later ChartEx UI/native evidence is recorded below.

### Further neutral chart saves, 2026-10-26

Shared operations now preserve styled title runs, axis source-linked formats and date/label/display
units, hidden-label content and value-from-cells extension siblings. PowerPoint group spacing maps
to the neutral schema-validated operation. Formatting-only changes leave embedded workbook cells
alone; data edits retain blank cells. The title/axis increment passed the 300-case saved-part corpus
plus 41 focused tests. Label-range and group edits have focused preservation regressions; expanded
corpus and native acceptance remain required before closing priority 4.

### SmartArt styles and generated caches, 2026-10-26

Original implementation commits: `a58e44d20` (interpreter/painter/theme), `2056cc734` (partial-layout
diagnostics), `26af3ba4c` (shared style/colour controls and writers), `4a3c51714` (host cache wiring).
Word and Excel share three style intensities and five theme-accent colour schemes. Edits preserve
per-node fill/line/rich-text overrides and unknown definition XML. Generated drawing caches use
semantic node associations, host relationships/content types and an integrity marker. Cache
generation admits interpreted layouts with no reported issues; spatial 3D, folding, unsupported
text and partial geometry remain explicit uncached/fallback cases.

The final SmartArt increment passed 908 core tests across 120 files and 50 UI tests across 11 files,
both core/UI compiler projects and builds. Browser validation covered Word's eight and Excel's
seven cases across six bindings. Native Word 16 and Excel 16 reopened six style/colour variants
with the expected semantic text and style/colour ids and nonempty node shapes; Excel reported no
repair. These are semantic/package checks, not native pixel comparisons or the complete Office
style catalogue.

### Built-in SmartArt catalogue and 3D caches, 2026-11-16

`make-smartart-definitions.ps1` captured every entry of PowerPoint 16's `SmartArtColors` (38) and
`SmartArtQuickStyles` (14) through COM; the generator stores them gzip-packed in
`diagram/builtin` and the gallery metadata (names, categories, tile traits) in the main entry.
Word, Excel and PowerPoint galleries list all 52; Word and Excel apply a pick by replacing the
definition part, PowerPoint on save (in its editor a pick is previewed through the colour patch
and nearest intensity, an approximation). Quick-style 3D is written into generated caches per
shape, matching the structure of Office's caches observed for all 14 styles; the spatial camera
of the five Scene styles stays in the quick-style part, as in Office's cache. Tests: every
definition parses, matches its metadata, and saves/reopens on the Word fixture with a generated
cache (`diagram/builtin/style-definitions.test.ts`); PowerPoint saves both parts and the
document-point ids. Native: 41 Word 16 and 41 Excel 16 files (all 38 colours, all 14 styles)
reopened with matching texts and ids, native node shapes and no Excel repair. Browser: the Word
SmartArt specs (8, six bindings) and Excel SmartArt specs (13) pass with the new galleries; the
PowerPoint browser suite was not run for this change. Not established:
pixel parity of the generated caches or our flattened 3-D painting.

### SmartArt round-trip cache, 3D painting and gallery names, 2026-11-16

Round-trip writer: `buildGeneratedDiagramCache` binds each regenerated shape for a node and layout
role Office presented to Office's own presentation point (`write/cache-presentation.ts`), so an
edited diagram keeps Office's points and connections and only new nodes get generated points.
Placement comes from the layout engine, not Office: on the unedited diagrams of the 229
PowerPoint gallery fixtures and the Word fixture (`scripts/diagram/measure-source-cache-regeneration.ts`)
the writer admits 52 of 230; 311 of their 414 regenerated shapes bind to an Office point; the
worst bound-shape frame deviation is within 1% of the diagram's larger side in 27 diagrams, 5% in
34 and 10% in 36, with large gaps in Captioned Pictures, Square Accent List, Step Down Process and
the bracket and block lists (21% to 40%). The other 178 stay uncached as before. Native: the
acceptance build gained `edit-3d1` and `edit-3d9` (text edit, added node, then Polished or Bird's
Eye Scene) for Word and Excel; all 86 files reopened in Word 16 and Excel 16 with the edited
texts, the expected ids, one native shape per node and no Excel repair log.

3D painting: Word and Excel draw a bevel lit by an SVG specular filter from the label's light rig,
an extrusion as stacked darkened copies along the camera's depth, and the Scene styles' camera as
a per-shape affine projection through the scene camera PowerPoint fits (`diagram/scene-camera.ts`,
moved from the pptx renderer; loaders carry the quick style's camera on the drawing). Tests:
`diagram/scene-3d.test.ts`, `src/ui/src/smartart/smartart-3d.test.ts`. Not established: any
pixel comparison with Office; there is no hidden-surface removal, material model or contour.

PowerPoint preview: a SmartArt Styles pick applies the definition's per-label 3D to the drawing
shapes and a Scene style's camera to the diagram (`smartart-style-3d-preview.ts`), so the lit
scene shows it before reopening; colours and effects stay at the nearest intensity until then.
Tested in `smartart-style-3d-preview.test.ts`; the PowerPoint browser suite was not run.

Gallery names: Word and Excel name the titles, sections and all 52 entries in every product
locale from PowerPoint's `pptx.gallery.smartArt*` translations (`diagram/strings/gallery.ts`).
Those translations are agent-authored and not verified against Office's localized names.

### Advanced chart controls and ChartEx hosts, 2026-10-26

Original commits `580d123fb`, `abfa28fb3`, `8d4ac7a86`, `fcfa5dbf9` and `4ae7858f3` introduced
shared advanced operations, resolved appearance conversion, gradient geometry preservation, rich
text formatting and the shared controls. `4864c8622` expanded secondary-axis splits with stable
selection indices and corrected column clustering. The final focused control checks passed 60 UI
and host tests; advanced control/save browser cases passed across all six Word and Excel bindings
and the targeted Visio chart workflow. Later secondary-axis checks passed 90 core and 40 UI cases.

ChartEx original commits `725de251`, `4f84b46a`, `1f90137d`, `14bf8847` and `3af9f25c8` add shared
data-sheet/action contracts, safe local workbook mapping/growth, Word/Excel history and save, family
options, axis controls, point value labels and common formatting. Twelve Word/Excel browser cases
passed across six bindings. Native Office 16 checks verified Word ChartData values, workbook
growth and a two-series box chart; five Excel variants reopened without repair with expected
chart counts/types and series. External data and ambiguous hierarchy sources are refused.

### PowerPoint chart migration checkpoints, 2026-10-26

Original `4ccd7e9fc` adds sparse analysis/marker/series/point save planning and existing-axis-pair
reassignment. `d78fb64f5` adds category-axis combo type edits while preserving document-order
targets. The latter checkpoint passed 130 focused tests and both core compiler projects. Unsupported
3D/XY combo switches and unsafe splits remain compatibility cases. Root integration and broader
corpus results are required before closing priority 4.

ChartEx read facts (`02c04bf5`) and shared creation/save increments (`a63107c6b`, `1689fbc6e`) cover
six authored families with source/workbook synchronization. PowerPoint 16 reopened a valid
20-slide gallery after a waterfall-to-funnel edit; six generated neutral chart parts grafted into
valid gallery hosts had the expected native chart types. A fresh six-slide SDK deck failed native
opening with `0x80070570` using both the baseline and new generators. This is a known new-deck
template limitation, not evidence that complete new-deck generation has native parity. Raw
geographic/Pareto creation, wider ChartEx styling/labels/axes and remaining compatibility adapters
still need explicit coverage. These checkpoints do not establish Office equivalence or lossless
export for arbitrary documents.

### Remaining PowerPoint source adapters, 2026-11-16

A probe over the 75 classic chart parts of the committed decks listed the `PptxChartData` fields
whose edits no neutral family claimed. The first follow-up family moves the group subtype settings
the object-tree writers write (3-D bar shape, 3-D gap depth, radar style, surface wireframe,
pie-of-pie split options and bubble options) onto the shared `setChartGroupSubtype` operation, under
the same chart-type conditions and removal rules. `chart-subtype-migration.test.ts` saves each edit
through both paths over every deck: 33 parts reload with the same edited value both ways; 5 save to
the same canonical XML and the rest differ only in the recorded object-tree side effects (theme font
pinning, JavaScript number caches, manual-layout number form, resolved outline colours, one series
`c:shape` position). No fixture holds a pie-of-pie chart, so that case is unit evidence only. Cleared
settings the object-tree writers keep remain on that path.

Group settings on their own: `chart-neutral-save-group-flags.ts` claims `varyColors`, the bar
direction, the scatter style and the `standard` grouping flag of a kept chart type through
`setChartVaryColors`, `setChartBarDirection` (every group a bar group; the axes swap sides),
`setChartScatterStyle` and `setChartGrouping`. A probe that saved each edit through the object-tree
writers found that they wrote no `c:varyColors`, `c:barDir` or `c:scatterStyle` on an existing part,
so those edits were lost on save. `chart-group-flags-migration.test.ts` saves each edit both ways:
70, 31 and 2 parts reload with the edit through the neutral writer and none through the object-tree
writers; the grouping flag reloads on all 8 3-D bar parts both ways (one to the same canonical XML).

Gridline flag: the chart-wide `style.hasGridlines` is read off the first value axis's
`c:majorGridlines` and the object-tree writers never wrote it, so the flag edited on its own was
lost on save. `chart-neutral-save-gridlines.ts` sets the major gridlines of that axis
(`setChartAxis`), as the shared inspector's Show Gridlines patch does, and claims a flag that only
mirrors an axis edit with it. `chart-gridline-flag-migration.test.ts`: 51 parts keep the toggled
flag through the neutral writer; the object-tree save kept it on the 5 parts whose axis already drew
gridlines.

Series number format: `PptxChartSeries.numberFormat` is read from the series' label number
format, falling back to the value cache's code, and the object-tree writers never wrote it.
`chart-neutral-save-series-format.ts` writes it as the series' label number format, not linked to
the source; `setChartDataLabels` now gives a series without labels of its own its group's switches
and position when an edit only formats them, so the series keeps showing what it showed.
`chart-series-number-format-migration.test.ts`: all 75 parts keep the edit through the neutral
writer and none through the object-tree path. Clearing the format stays unclaimed (the cache code
would read back).

XY combinations: Excel 16, driven through its object model (`Series.ChartType = xlXYScatter` on a
column or line series), saves a `c:scatterChart` that references the same category and value axes
as the column or line group, with the categories as `c:xVal`, a 5 pt circle marker in the series
colour, no line and `c:smooth val="0"`; asked to make one series 3-D column or bubble, it changes the
whole chart. The shared `setChartSeriesType` now writes that structure. The object-tree writers
wrote the scatter series with `c:cat` and `c:val`, which CT_ScatterSer does not allow;
`chart-xy-combo-migration.test.ts` saves the edit both ways over the corpus: 16 parts reload with the
same series types, names and values, 12 save through the neutral plan and 4 (two splits inside a
stack, two groups with an unknown extension) keep the
object-tree path. A whole-chart switch from such a combination now restores each series' data from
its own group, and a picture-filled series switched to a line family keeps the picture off its
outline (CT_LineProperties has no picture fill), which previously made the shared writer throw and
the part's save fail; a writer failure now falls back to the object-tree path.

Native acceptance (`pptx-classic-chart-acceptance.ps1 -RequireActiveDataSheet`, PowerPoint 16.0,
2026-11-16): all 11 sequential cases passed chart type, series, category, value and Edit Data cell
checks, including the column chart whose third series became scatter (series types 51, 51 and
-4169), the 3-D column with a cylinder shape and gap depth 220 (92, `xlCylinderColClustered`) and the
filled radar (82). An earlier run without `c:smooth` reported the scatter series as 72
(`xlXYScatterSmooth`), which is why the conversion writes it explicitly. This is open and data
evidence, not a pixel comparison. Surface wireframe, pie-of-pie and bubble option edits were not
opened natively.

ChartEx custom data labels: CT_DataLabel of the chartex schema ([MS-ODRAWXML]) holds a number
format, shape and text properties, visibility flags, a separator and `@pos`, but no label text and
no manual layout. Excel 16, driven through its object model on a waterfall chart, raised "not
implemented" for `DataLabel.Text` and `DataLabel.Left`, and the part it saved kept only the series'
visibility flags. Per-point visibility, position, separator, number format, font, shape and hiding
already go through the shared `pointLabel` action in PowerPoint, Word and Excel; a PowerPoint edit
of a ChartEx label's text or manual position stays refused, now with a message naming the cause.
No custom-label node is invented.

### Chart picture fills, error bars and native analysis checks, 2026-11-16

- Shared painter: picture fills on series, points, plot area, chart area, title and legend
  (`chart/render/chart-picture-paint.ts`); `picture` intent and pending pictures
  (`chart/edit/picture.ts`); load and save helpers in `chart/picture-parts.ts`.
- Hosts: Word reads chart pictures at load and stores pending ones for edited and new charts;
  Excel resolves them from the source package and stores them once the package is zipped; Visio's
  embedded workbook does the same. Save and reopen tests: `chart/picture-fill-flow.test.ts`
  (docx, xlsx and visio packages), `xlsx/edit/chart-picture-fill.test.ts`,
  `xlsx/embedded-chart-space.test.ts`, `ui/src/docx/chart-format-pane.test.ts`,
  `ui/src/xlsx/chart-format-pane.test.ts`, `ui/src/visio/viewer-chart-editor.test.ts`.
- PowerPoint: `withChartSeriesTrendline`, `withChartSeriesErrorBars` and the pane's Trendlines
  and Error Bars sections (`ui/src/pptx/render/chart-format/chart-format-error-bars.test.ts`), saved
  and reopened through `PptxHandler` (`pptx/core/builders/sdk/chart-analysis-save.test.ts`).
- Painter gaps closed: tile mirroring (`flip` x, y, xy), `stackScale` with `c:pictureStackUnit`
  and `a:srcRect` crops (`chart/render/chart-picture-paint.test.ts`).
- Native (`src/core/scripts/chart-analysis-native-acceptance.ps1`, Word, Excel and PowerPoint
  16.0 through COM, one application at a time; PowerPoint's trendlines set through
  `withChartSeriesTrendline`): a clustered column with two trendlines, custom
  error values and picture fills on a series and the plot area, and a stacked column whose last
  series moved to the secondary axis. Word (`word-analysis.docx`, 2 charts), Excel
  (`excel-analysis.xlsx`, 2 charts, no repair log) and PowerPoint (`powerpoint-analysis.pptx`)
  open and read back the series axis groups, trendline counts, error bars and picture or texture
  fill types. `powerpoint-stacked-secondary.pptx` fails: PowerPoint reads series 3 on the primary
  axis, because PowerPoint's save does not write the stacked end move (open above). Reads only;
  painting is not compared.

### Documentation reconciliation, 2026-10-27

Active roadmap, architecture plans, product capability pages and ribbon reviews now distinguish
implemented shared controls from unfinished PowerPoint adapter coverage. Historical audits and
older translations are identified explicitly. Word, Excel and PowerPoint documentation sites
built successfully; Visio's documentation build/link/theme suite passed 76 tests. Added relative
links, formatting and whitespace checks passed. Integrated runtime verification was still pending
at that checkpoint; its completed result is recorded in the final verification section below.

### Final chart integration and native acceptance, 2026-10-27

Classic saves now include typed numeric/date/hierarchy vectors, blank points, rich runs, source-backed
protection and stock decorations. Structural workbook edits allocate new sheets without rewriting
existing worksheet parts. The package adapter refreshes its chart model after saving, so repeated
type and data edits use current series identities and axes. Numeric categories can become text
without losing source formulas or authored series ids. The 300-test classic corpus passes with
the unchanged category total of 71/71 for both canonical XML and reloaded models.

Independent review added regressions for theme-preserving text outline/gradient/effect edits,
rich-label fonts, numeric point format codes, mutually exclusive dash choices and foreign layout
attributes. ChartEx now shares DrawingML style operations and typed axis/group/point-label edits;
unsupported axis and custom-label requests are refused explicitly. Native PowerPoint 16 reopened
the extended-chart example with the requested pattern, hidden point, title font, axis title/bounds
and serialized tick units. This is semantic evidence, not a pixel comparison.

The committed `src/core/scripts/pptx-classic-chart-acceptance-build.ts` and
`pptx-classic-chart-acceptance.ps1` exercise five sequential edits in a real producer deck: growth,
pie, line3D, a repeated line save, and shrink to bar. PowerPoint 16 passed all five chart-type,
series, category, value and referenced workbook-cell checks. With `-RequireActiveDataSheet`, the
active sheets were `ChartData1`, `ChartData2`, `ChartData3`, `ChartData3`, `ChartData4`, matching the
actual chart references. All 100 chart-rendering browser cases passed across the five PowerPoint
bindings. Word/Excel/Visio native and browser evidence remains recorded in the sections above.

The final Word/Excel ChartEx controls use the same style projection for pattern fills, effects,
text, shape 3D and outlines; native ChartEx axis/family controls remain separate. Twelve browser
cases passed across all six bindings for both products, including pattern edits, data growth and
save/reopen. Secondary-axis authoring remains disabled for ChartEx.

### Final repository verification, 2026-10-27

At `7776ac5d5`, `bun run verify --jobs 4` passed all 46 CI-planned steps from the isolated integration
worktree after integrating current `origin/main`. The affected core suite passed 20,940 tests
(16 existing expected failures and 423 skips); the affected UI suite passed 11,020 tests
(11 skips). Product and PowerPoint framework typechecks/tests, locale checks, public-prop checks,
tree-shaking limits, builds, lint and changed-file formatting all passed.

The clean-install package smoke test imported every entry point and checked 130 browser entries
for Node-only dependencies. The integrated Office suite also built successfully. Browser and native
acceptance were run separately as detailed above; the default `verify` command does not run them.

The broad run exposed and fixed an omitted-origin PowerPoint SmartArt reflow case: all 14 reflow
families now produce finite coordinates for size-only bounds, while invalid-geometry rejection
remains strict. The ChartEx corpus now compares hidden-only labels and numeric scaling explicitly,
with authored-source assertions. No corpus totals were lowered to complete this delivery.
