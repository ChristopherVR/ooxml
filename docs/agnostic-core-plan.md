# Plan: making ooxml-core format-agnostic

Status on 2026-10-23, after Wave 6 (`2fedfaae7`): steps 1 to 5 are done for the parts listed below,
and step 6 (pptx off `XmlObject`) is started for charts only; section 6 was not pursued. The shared
areas `drawingml`, `diagram` and `chart` (with `chart/edit`, `chart/ex`, `chart/workbook` and
`chart/render`) exist, and docx, xlsx, visio and pptx consume them. Chart and SmartArt editing, which
was PowerPoint-only, now runs in Word and Excel (SmartArt) and in Word, Excel and Visio (charts)
over the neutral models. The "Remains" lines are the open items. Counts below were taken from
`origin/main` (`c3e0979`, 2026-10-08) with `find`/`grep`; they are approximate and describe coupling,
not quality.

The step labels used in code comments and commit messages (chart steps C1 to C6, chart UI steps U1
to U5, SmartArt steps S1 to S5) are defined in sections 3.5 to 3.7. Their order follows
[structure-plan.md](structure-plan.md) section 5.3.

Goal: everything that is not specific to one Office format lives in a format-neutral area with a
neutral public model, and `docx` and `pptx` (and now `xlsx` and `visio` for charts and SmartArt) are
thin adapters over it. Word can then use what only PowerPoint had: SmartArt, charts, then DrawingML
primitives (fills, lines, effects, text bodies, theme).

## 1. Where the shared logic lived (starting point, 2026-10-08)

`src/core/pptx/core` held 79,630 lines of SmartArt, chart and DrawingML code mixed with slide code in
flat folders (`core/utils` about 1,700 files, `core/core/runtime`, `core/core/builders`,
`core/types`). Only file names separated generic DrawingML from slide logic.

### 1.1 The coupling, in four layers

| Layer                               | What it was                                                                                                                                                                                                                                                                                                          | Why it was not portable as is                                                                                                                                                                |
| ----------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| XML model                           | `XmlObject` (`types/common.ts`): `fast-xml-parser` trees, attributes as `@_name`, same-name siblings grouped into arrays, prefixes in keys (`a:solidFill`) looked up by local name through `xmlLookupService`.                                                                                                       | docx and `src/core/xml` use a namespace-aware DOM (`@xmldom/xmldom`). Every pptx parser read `node['@_x']`. 87 of 162 `utils/chart-*` files and nearly every smartart file took `XmlObject`. |
| Handler runtime (`this`)            | A mixin chain `PptxHandlerRuntime` (`core/core/runtime/*`): `parseSmartArtDrawingShapes`, `parseDrawingShape`, `parseSmartArtQuickStyle` and chart parsing called `this.parseColor`, `this.resolveThemeFillRef`, `this.readXmlPartByRelationshipId(slidePath, relId)` and `this.compatibilityService.reportWarning`. | The base class owned the JSZip instance, the theme maps, the rels maps and the slide being loaded, so a shape parser could not run without a presentation.                                   |
| Types                               | `PptxSmartArtDrawingShape extends PptxCustomPathProperties`, `PptxSmartArtData` embeds `TextSegment`, `Pptx3DScene`, `PptxSmartArtChrome`; `PptxElement`/`PptxSlide` appear in `smartart-decompose-*` and the chart builders.                                                                                        | Public types were named and shaped for PowerPoint; the pure data (nodes, connections, colours, quick style) was tangled with slide editing fields.                                           |
| Pure algorithms (no XML, no `this`) | `smartart-engine/` (layout algorithms, constraint solver, text fit, font metrics tables), `smartart-layout-interpreter-*`, `smartart-hierarchy-*`, `smartart-editing-*` (reflow, node ops) over typed `PptxSmartArt*` models; chart data utils; `color/` and `geometry/` (already moved).                            | Cleanest to move, but they imported the pptx types barrel (`from '../types'`), so the typed model had to move first.                                                                         |

Coupling verdict: the **data shapes** (point, connection, relIds, custom layout, drawing geometry and
text) are thin and reached `XmlObject` only through attribute reads. The **cached drawing shape
parser** was wired to the runtime for colour, gradient and theme resolution. The **layout engine** was
XML-free after its input was parsed, but its types came from the pptx barrel. The **writer and
fabrication** code was object-tree specific.

### 1.2 SmartArt map

- Part model and XML parsing: `core/builders/PptxSmartArtParser.ts` (data points, connections),
  `core/runtime/PptxHandlerRuntimeSmartArtParsing.ts` (connections with transition labels, quick style,
  cached drawing shapes, `parseDrawingShape`), `SmartArtXmlUtils`,
  `utils/smartart-data-model-attributes.ts`, `smartart-connector-labels.ts`,
  `diagram-relationship-ids.ts`, `core/runtime/smartart-drawing-part.ts` (resolves `dsp:drawing`
  through `dsp:dataModelExt` or the data part's rels), `smartart-colors-builder`,
  `smartart-style-label-refs`, `smartart-drawing-shape-{style,3d,font-ref}`, `smartart-text-paragraphs`,
  `smartart-text-style-resolution`.
- Types: `types/smart-art*.ts` (`PptxSmartArtNode`, `PptxSmartArtConnection`,
  `PptxSmartArtDrawingShape`, layout definition, colour and quick-style definitions, constraint rules,
  chrome).
- Layout engine and interpreters: `utils/smartart-engine/*` (about 60 files, own ordered-XML
  layout-definition parser), `utils/smartart-layout-*`, `smartart-hierarchy-*`,
  `smartart-constraint-*`, `smartart-decompose*`.
- Editing and writing ("fabrication", generating parts from scratch):
  `core/runtime/smartart-fabrication-*`, `smartart-xml-builders`,
  `PptxHandlerRuntimeSaveSmartArtFabrication`, `utils/smartart-editing*`.
- PPTX-only glue: `PptxGraphicFrameParser` (finds `dgm:relIds` in a `p:graphicFrame`),
  `SmartArtElementProcessor` (converter), `smartart-decompose` (explodes SmartArt into slide shapes),
  the save pipeline.

### 1.3 Charts map

`types/chart*.ts` (about 2,000 lines of model), `utils/chart-*` (162 files, about 87 importing
`XmlObject`; 10 pure), `core/runtime/PptxHandlerRuntimeChart*` (parsing 1,105 lines, colour style,
external data, user shapes), `builders/sdk/chart-operations.ts` (1,064 lines), `chart-xlsx-parser`
(embedded workbook), ChartEx (`chart-cx-*`), colour and style parts (`chart-color-style-writer`,
`chart-style-definition`). Parsing started from a `p:graphicFrame` on a slide; the part loader was
`readXmlPartByRelationshipId(slidePath, relId)`. Word needs the same part (`c:chart r:id` inside
`w:drawing`) and the same `c:chartSpace` model.

### 1.4 DrawingML primitives map

Fills, lines, effects and text: `PptxShapeStyleExtractor`, `PptxGradientStyleCodec`,
`PptxColorStyleCodec`, `PptxColorTransformCodec`, `PptxEffectDagExtractor`, `effect-dag-*`,
`shape-style-line-helpers`, `drawing-fill-xml`, `body-properties-parser`,
`paragraph-properties-parser`, `text-run-properties-parser`, `drawing-line-dash`, `gradient-angle`;
theme in `PptxHandlerRuntimeTheme*` (loading, format scheme, ref resolution, `phClr` substitution) and
`types/theme.ts`; custom geometry (`custom-geometry*`, left in pptx by the earlier `geometry`
extraction). `docx` had its own small theme model (`theme.ts`, `theme-color.ts`) and its text runs use
WordprocessingML, so only the DrawingML text body (inside shapes, charts, diagrams) was shared.

## 2. Target areas

Every area is a subpath of the single package, strict TypeScript, written against the shared `xml`
DOM and `opc`. Neutral public types never mention `Pptx*`, `Docx*`, `XmlObject`, a slide or a document.

| Area        | Contents                                                                                                                                                                                                                                                            | Depends on                          |
| ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------- |
| `drawingml` | `DrawingColor` (srgb/scheme/sys/prst/hsl/scrgb + transform list) and its resolver, `Fill`, `Line`, `Effect` models, `TextBody`/paragraph/run models, xfrm, `ThemeLike` (colour scheme, font scheme, format scheme) and `phClr` resolution. DOM parsers and writers. | `xml`, `color`, `geometry`, `units` |
| `diagram`   | The SmartArt part model: data model, layout definition header, colours, quick style, cached `dsp:drawing` shape tree; the relationship resolver; the layout engine, node text, edit operations and part writers.                                                    | `drawingml`, `xml`, `opc`           |
| `chart`     | `c:chartSpace` model, parse/serialize, `cx:` ChartEx model, colour/style parts, embedded workbook reader and writer, edit operations, a renderer-neutral series/axis view-model and SVG painter.                                                                    | `drawingml`, `xml`, `opc`           |
| `theme`     | A sub-module of `drawingml` rather than its own area; `docx`, `xlsx` and `pptx` feed it.                                                                                                                                                                            | `drawingml`                         |

`diagram` shipped first with a **provisional** private colour/fill reader (`diagram/drawing-color.ts`
and `drawing-fill.ts`) because `drawingml` did not exist yet. Those files were written against the DOM
and a `DrawingColor` type, with no diagram knowledge, so they moved to `drawingml` unchanged (step 3).

### 2.1 Adapters

A neutral area exposes **parsers over DOM elements or part text plus small injected interfaces**; a
format area supplies the host:

- `PartHost`: `readText(partName)`, `resolve(basePart, target)`, relationships by part. The pptx
  adapter implements it over its JSZip and `slideRelationships`; the docx adapter over the document
  package. This replaces `this.readXmlPartByRelationshipId(slidePath, relId)`.
- `ColorResolver`: `(DrawingColor) => hex | undefined`. pptx binds its theme maps; docx binds
  `src/core/docx/theme.ts`.
- Diagnostics: parsers return `issues: {code, message}[]` instead of calling a compatibility service.
- pptx adapter also converts between the neutral model and its legacy `Pptx*` types (aliases where
  shapes are identical, mappers where not) so the public pptx API does not change.
- docx adapter owns the host markup (`w:drawing`/`wp:inline`/`wp:anchor`, extent, wrap, alt text) and
  calls the neutral area for the graphic part.

For code that still takes `XmlObject`, a neutral parser may be driven through a **neutral attribute
reader** (`(name) => string | undefined`) instead of a DOM element; pptx passes
`(n) => node['@_'+n]`. That is how the connection, custom-layout and relationship-id parsers are shared
without converting pptx to the DOM first. A full object-tree view of a DOM is deliberately not built: it
would keep two XML models alive.

## 3. Migration order

1. **`diagram` part model and cached drawing (done).** Neutral data model, connection, custom-layout
   and relIds parsers, relationship resolver, colours and quick-style summaries, cached drawing shape
   tree, loader. pptx re-imports the neutral parsers and types. docx maps `w:drawing` to `DocxDiagram`.
2. **`diagram` model types and layout engine (moved; remains listed).** Survey of what step 2 covers
   (counts are non-test source files; "XML" means the module walks the pptx `XmlObject` tree):

   | Family                                                                       | Files               | Imports from pptx                                                                                                                                                                   | Coupling                                                                                                                                                                                                        | Status                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
   | ---------------------------------------------------------------------------- | ------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
   | `types/smart-art*.ts`                                                        | 8                   | `XmlObject` (`rawXml` slots), `TextStyle` (node runs), `Pptx3DScene`/`Pptx3DShape`/`Text3DStyle` (quick style), `TextSegment` and `PptxCustomPathProperties` (cached drawing shape) | Node, layout definition, constraints, colour transform, quick style and layout variables are pure data with raw-XML, run-style and 3D slots; the drawing shape, chrome and definition headers carry pptx models | **Moved** to `diagram/model` as `Diagram*` with type parameters for the slots; `PptxSmartArt*` are aliases (`pptx/core/types/smart-art-model.ts`), `PptxSmartArtData extends DiagramData`. Drawing shape, chrome and definition headers stay in pptx.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
   | `utils/smartart-engine/`                                                     | 54 (+7 deck tests)  | the types barrel, `geometry` (preset evaluator), `smartart-layout-item-font-size` (`resolveFontTable`)                                                                              | Own ordered-XML layout-definition parser over the raw part text; no `XmlObject`, no `this`                                                                                                                      | **Moved** to `diagram/engine`, strict. The preset evaluator went to `geometry`, the glyph tables and kerning to `text/font-metrics`, the shared text-fit and layout-result modules to `diagram/layout`. Seven tests that load decks through `PptxHandler` stay in pptx.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
   | `smartart-constraint-*`                                                      | 8                   | types barrel, `clone-utils`, interpreter model                                                                                                                                      | Reads `dgm:constr`/`dgm:rule` through the neutral attribute parsers; the solver used the interpreter model, whose `dgm:choose` walkers read the layout node's raw XML                                           | **Moved** to `diagram/layout` (solver, resolve, branch index, declared-by, sibling roles, ratio fallback, interpreter constraint helpers). Attribute parsing, validation and write values are neutral (`diagram/layout/smartart-constraint-rules.ts`); `smartart-constraint-rules` stays in pptx as the object-tree reader and writer.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
   | `smartart-hierarchy-*`                                                       | 39                  | types barrel, constraint solver, interpreter model                                                                                                                                  | Pure arithmetic over typed nodes, but every family reaches the solver and interpreter model; `smartart-hierarchy-tailed-transpose` walks `rawXml`                                                               | **Moved** to `diagram/hierarchy` (the 39 modules and `arrangeHierarchy`), with its closure in `diagram/layout` (node forest `smartart-tree`, item font size, tier and role, hierarchy font fit, preset node, render and shape-preset helpers). The raw `dgm:choose` readers take the `RawXmlView`; pptx keeps `arrangeHierarchy` as a bound facade and `smartart-helpers` for the slide-element factories.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
   | `smartart-layout-interpreter-*`, `smartart-layout-item-*`, `smartart-node-*` | 93 interpreter + 20 | types barrel, solver, `core/runtime/smartart-text-reconciliation`, `color`, `xml-access`                                                                                            | 5 interpreter files walk `XmlObject`; text projection reconciled paragraphs through the pptx runtime's text parsers                                                                                             | **Moved**, except the pptx-bound pieces. The `dgm:choose` walkers, `dgm:if` evaluation, axis navigation, composite detection, interpreter model and every arranger and `interpretSmartArtLayout` live in `diagram/layout`; everything that reads raw `dgm:choose`/`dgm:forEach` XML takes a `RawXmlView`, and pptx binds its entry points to its raw slots through `smartart-ordered-xml-adapter` (bound facades keep the old pptx module names and signatures). The node text projection, its reconciliation and the plain-text views moved to `diagram/text` (2026-10-23, see S2). Stay in pptx: `projectFoldedNodeText` (a thin pptx wrapper), `smartart-node-style-apply` (writes `PptxElement`s), `smartart-node-style-role` and `smartart-node-pres-vars` (object-tree readers of the data model the neutral `data-model` parser already covers), `smartart-interpreter-drawing-bridge` (builds `PptxElement`s). |
   | `smartart-decompose*`                                                        | 3                   | `PptxElement`, `TextSegment`                                                                                                                                                        | Explodes SmartArt into slide elements                                                                                                                                                                           | **pptx-only**, never moves.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
   | `smartart-editing*`                                                          | 8                   | types barrel, interpreters                                                                                                                                                          | Reflow calls the interpreters                                                                                                                                                                                   | **Moved** to `diagram/edit` (2026-10-23), strict and generic over the neutral data model (see S3). The pptx modules are facades with the historical names.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |

   **Remains:** the pptx `TextStyle` styler (`smartart-node-text-projection`) and the object-tree
   paragraph parser and builder (`smartart-text-paragraphs`) stay until pptx reads SmartArt text through
   the neutral model; `smartart-decompose*`, the drawing bridge and the node style writers never move.
   The hierarchy arranger is generic over the raw slot (`RawXmlView<R>`) while the other interpreters take
   `RawXmlView<unknown>`; unifying them is cosmetic. The engine's `layout-def-parse` still parses its own
   ordered XML rather than the shared `xml` DOM; the interpreter walkers share that ordered tree, so
   moving both to the DOM is one later step. Parity for the choose walkers is pinned by
   `pptx/core/utils/smartart-choose-walkers-corpus.test.ts` (every walker over the 353 distinct layout
   definitions of the committed decks and built-in layouts).

3. **`drawingml` extraction (started; remains listed).** `src/core/drawingml` exists
   (`ooxml-core/drawingml`). Moved from `diagram` with history: the colour reader and resolver, ordered
   colour transforms, fill and line readers, text-body reader, xfrm and geometry readers, shadow, colour
   and fill writers, gradient presets and geometry, rect path gradients and the DOM attribute helpers;
   the neutral types are `Drawing*` (`DrawingColor`, `DrawingFill`, `DrawingTextBody`...) and `diagram`
   keeps its `Diagram*` names as aliases. Also new: the theme model (`DrawingTheme`: colour scheme as
   `DrawingColor`, font scheme with script overrides), `parseTheme` over the DOM, `parseThemeColorMap`
   and `resolveSchemeColor` with the `clrMap` indirection. docx (`theme.ts`, `diagram-theme.ts`,
   `theme-color.ts` defaults) and xlsx (`read/theme.ts`, slot order in `write/theme.ts`) read their
   themes through it; `chart` and xlsx chart code import the readers from `drawingml`. visio reads its
   `a:clrScheme` through `parseTheme` and resolves theme colours with `parseDrawingColor` and
   `resolveDrawingColor` behind a thin guard for its stricter limits (colour kinds, `a:` namespace,
   shade/tint/alpha only, `hueMod` ignored); its `vt:` extensions and `fmtScheme` selection stay in visio.
   The text model (`drawingml/text-model.ts`, with its reader and writer, 2026-10-23) is pinned equal to
   the pptx parsers by `pptx/__tests__/integration/smartart-text-neutral-parity.test.ts`.

   Shared outlines now cover compound/alignment, joins, custom dashes and arrow ends; typed effect
   lists, source-preserving pattern/picture fill codecs, theme format schemes and `phClr` style
   resolution are implemented (2026-10-26). Chart and diagram consumers use them, with real
   Word/Excel/PowerPoint producer fixtures and strict parser/writer/resolver tests.
   **Remains:** effect DAG interpretation, broader effect rendering and the remaining pptx codecs
   (`PptxHandlerRuntimeTheme*`, `types/theme.ts`, colour, gradient and effect adapters) stay on their
   object tree until step 6 or a parity-tested shim; the pptx text parsers still read through their own
   object tree (the neutral model is pinned, not yet used by pptx); the xlsx theme writer still patches
   the part itself (a neutral writer exists as `drawingml/write-theme.ts`; whether xlsx should adopt it is
   open).

4. **Drawing shape parsing off the runtime (done for the neutral reader; remains listed).**
   `DiagramShapeView` (`diagram/drawing-view.ts`) is the resolved shape that renderers and edit operations
   use, and `PptxSmartArtDrawingShape` extends it. `readDiagramDrawingShapes` and
   `resolveDiagramDrawingShape` (`diagram/drawing-resolve.ts`) read a `dsp:drawing` through
   `parseDiagramDrawing` (which keeps each shape's text body on the neutral rich text model and its shadow
   colour) with a host `DiagramColorResolver`.
   `pptx/__tests__/integration/smartart-drawing-shape-neutral-parity.test.ts` reads all 371 committed
   drawing parts (3014 shapes) through both readers and requires every view field to agree; the one
   counted divergence is a colour that repeats a transform (two `a:tint`), which the pptx object tree
   drops and the neutral resolver applies. The data part writer for a SmartArt created in memory moved to
   `diagram/write/data.ts` (`writeFabricatedDiagramData`, byte-identical to the pptx fabrication under a
   seeded id generator).

   **Remains:** pptx's `parseSmartArtDrawingShapes` still parses its own object tree (switching it to a
   mapper over the neutral reader means carrying its custom-geometry, 3D and text-segment slots across).
   The cached drawing writer (`smartart-fabrication-drawing`, which emits pptx custom geometry, 3D and
   text segments), the layout, colours and quick-style fabrication, and the merge of edits into existing
   data parts (`smartart-xml-builders`, object trees in the pptx save pipeline) stay in pptx: a DOM port
   of the merge path cannot match the object-tree serializer byte for byte without porting the save
   pipeline (step 6).

5. **`chart` model, parsers and writer (done for classic charts; remains listed).** Neutral `c:chartSpace`
   model (`src/core/chart/model.ts`, `model-series.ts`, `model-analysis.ts`): title, auto-title-deleted,
   view3D, plot area with one or more chart groups for every classic family, series with their references
   and caches, `spPr` through `drawingml`, markers, data points, data labels, axes with scaling and number
   formats, legend, `plotVisOnly`, `dispBlanksAs`, `c:style`, `c:roundedCorners`, raw `extLst` strings.
   `parseChartSpace` (`parse-*.ts`) reads it over the shared DOM and reports every child it does not model
   as an issue instead of throwing. It is tested on real parts from the chart-gallery, chart-data-fidelity,
   chart-top-axis, chart-filtered-series, chart-title-runs, chart-stacked-line-markers, three-d-charts and
   pie3d decks and the Excel and openpyxl xlsx fixtures. xlsx `read/chart.ts` builds its `ChartObject` from
   the model (identical output on every chart part in the fixture corpus). docx attaches a `DocxChart`
   (part name, parsed model, title, embedded workbook part name) to `w:drawing` charts; Word draws the
   chart since 2026-10-08 (`docx/chart-paint.ts`); unsupported families (bubble, stock, surface, chartex)
   were drawn as a labelled placeholder when this was written (see `docs/chart-renderer-gap-report.md`).

   - **Writer (2026-10-08).** `writeChartSpace` (`write-*.ts`) serialises the model in the ECMA-376 element
     order: booleans `1`/`0`, numbers in a canonical decimal form without an exponent (`chartNumber`),
     cached point text as held, `extLst` and `c:printSettings` verbatim, namespace declarations once on the
     root. To reproduce Office's bytes the model keeps the children of `c:spPr`, text bodies and `c:layout`
     as written (`sourceXml`, re-emitted while the modelled fields still read back from them; changed fills,
     lines and effects are patched into the kept `c:spPr`, changed text and layouts are written from the
     model), `c14:style`, `externalData/autoUpdate`, extra root declarations, gridline, chart-line, up/down
     bar and leader-line shapes, the multi-level cache point count and the line group's `smooth`. `drawingml`
     gained an `a:ln` writer. Round-trip tests (`write-space-fixtures.test.ts`): parse, write and parse keeps
     the model on every part the parser tests read and on the 278 parts embedded in the Excel acceptance
     fixtures; Office-written parts come back byte-identical except for the elements reported as not modelled
     (walls and floors, surface band formats at the time), literal CR LF inside `a:t` (normalised by the XML
     parser) and namespace declarations repeated below the root. Values Excel writes with 17 digits or an
     exponent outside a kept layout (axis bounds, units) come back in the canonical form.
   - **Parity harness (2026-10-08).** `src/core/pptx/__tests__/integration/chart-neutral-parity.test.ts`
     loads every committed deck that holds a chart part (22 decks, 81 parts) through `PptxHandler`, runs
     `parseChartSpace` on each raw part, projects both models onto one comparison shape and asserts equality
     per part. The 74 classic parts compare equal (141 series, 108 axes): chart type per group, bar direction,
     grouping, gap width, overlap, gap depth, vary colours, first slice angle, hole size, scatter and radar
     style, 3-D shape, wireframe, series order, `c:idx`, names, value, x and bubble-size caches (blanks
     included), format codes, smooth, invert-if-negative, explosion, markers, data points, series, group and
     point data labels, categories and multi-level categories, axes (ids, cross axis, position, delete,
     crosses, cross-between, orientation, min, max, log base, units, gridlines, tick marks, tick label
     position, number format, title), chart title, legend position and overlay, `plotVisOnly`,
     `dispBlanksAs`, `autoTitleDeleted`, `c:style`, view3D, `date1904` and rounded corners. The pptx
     normalisations (trimmed text, the `Series` default name, placeholder values for an empty cache, the
     legend-overlay default, the first-run flat title, the grouping collapse to `clustered` plus
     `groupingStandard`, unknown marker symbols as `auto`) are applied explicitly on the neutral side of the
     projection; axes are compared by id because pptx groups its axis list by element name. Recorded
     expectations: the 7 ChartEx parts are reported, not modelled, by the neutral parser, and pptx takes
     categories from the embedded workbook when the part caches none (2 parts). Bugs the harness found, all
     fixed: the neutral parser did not read the line group's own `c:smooth`; pptx missed a `c:style` written
     only under `mc:Fallback` (35 parts), fixed after `getChartStylePalette` in `ooxml-ui/pptx` was corrected
     to Office's gallery columns (style 2 starts at accent 1); and pptx read a bare data-label switch
     (`<c:showVal/>`) as unset, although `CT_Boolean` defaults to true. The other pptx `CT_Boolean` readers
     that go through `getChildByLocalName` share that blind spot (fast-xml-parser turns a bare element into an
     empty string); no committed deck writes one.
   - **Model gaps.** What one model carries and the other does not, from the harness and the two type sets:

     | Area             | pptx model only (`PptxChart*`)                                                                                                                                                                                                                                                                                                                                                     | Neutral model only (`ChartSpace`)                                                                                                                                                                                                                                                                                                                                                                                     |
     | ---------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
     | Part and package | resolved external data (target, auto-update), embedded workbook rows used as a cache fallback, colour style (`colorsN.xml`) and chart style (`styleN.xml`) parts, user-shape drawings, `clrMapOvr`, pivot source, print settings, protection                                                                                                                                       | `c:lang`, parse issues for everything not modelled, raw `extLst` of the part, chart, plot area, groups, series, axes, labels, markers and points                                                                                                                                                                                                                                                                      |
     | ChartEx (`cx:`)  | theme-resolved series, axis and label colours and fonts, `cx:valueColors` resolved to hex and `cx:valueColorPositions` read as `pos` children (the neutral model keeps both raw), the geography cache as an object tree                                                                                                                                                            | since 2026-10-23 (`ChartExSpace`, step C5): formulas (`cx:f`, `cx:nf`, `@dir`), level format codes and names, every point index, series `uniqueId`, `ownerIdx`, `formatIdx`, `hidden` and axis ids, data points, hidden labels and separators, plot surface, axis scaling, tick marks, chart title and legend alignment and overlay, external data, format overrides, print settings, `clrMapOvr` and extension lists |
     | Chart groups     | one chart-level set of group settings, read from the first group of a combo; surface band formats; drop-line, high-low-line and up/down-bar styling                                                                                                                                                                                                                                | every group of a combo with its own settings, axis id list, `c:marker` and `c:smooth`; drop lines, high-low lines, up/down bars and series lines as presence flags only                                                                                                                                                                                                                                               |
     | Series           | the resolved line colour, width and dash of trendlines and error bars (their structure is neutral since 2026-10-15), picture fills (`c:pictureOptions`, implied blip fills) resolved to images, `c16:uniqueId`, filtered series and titles (c15), data-label ranges and field tables (c15), callout shapes, `xForSave`, resolved colour, line width and dash, gradient descriptors | `c:order`, formulas (`c:f`) and source kind (reference or literal) for names, categories, values, x, y and bubble sizes, `c:ptCount` and per-point format codes, per-series categories (pptx keeps one list from the first series)                                                                                                                                                                                    |
     | Axes             | display-unit label styling, resolved font family, size, bold and colour, gridline styling as resolved colours                                                                                                                                                                                                                                                                      | document order, title rich text, layout and overlay, `txPr` as a full text body                                                                                                                                                                                                                                                                                                                                       |
     | Chart chrome     | floor, wall and data-table styling as resolved colours, legend entry text styles and typefaces resolved through the theme (`+mn-lt`), title runs with resolved styles                                                                                                                                                                                                              | unresolved `DrawingColor` (scheme colours and transforms) and `DrawingFill`/`DrawingLine` for every `spPr`, effect lists as raw XML                                                                                                                                                                                                                                                                                   |

   - **pptx reads classic charts through the neutral parser (2026-10-08, wave 4).** pptx builds
     `PptxChartData` for every classic chart part from `parseChartSpace` through an adapter
     (`pptx/core/utils/chart-from-neutral*.ts`): chart type, the first group's settings, categories and
     multi-level categories, every series' name, `c:idx`, value, x and bubble-size caches with their blanks,
     number format, smoothing, invert-if-negative, explosion and 3-D shape, the structural fields of
     markers, data points, series and point data labels and the group labels, the axes (ids, position,
     deletion, scaling, gridlines, number format, title text, tick marks and labels, crossing, units, display
     units), the flat title, legend presence, position and overlay, `c:style`, `plotVisOnly`, `dispBlanksAs`,
     `autoTitleDeleted`, `showDLblsOverMax`, `dispNaAsBlank`, view3D, `date1904`, `roundedCorners`, and whether
     a data table, floor or walls exist. The neutral model gained floors and walls, data tables, display units
     and `readDisplayNaAsBlank` for this. What the neutral model does not carry is still read from the part's
     object tree by the existing readers, handed the matching subtree (each chart group is paired with its
     object-tree node by element and position, each series by position): resolved colours, line width and dash,
     gradients and fonts (series, points, markers, labels, axes, title runs, legend entries, chart and plot
     area, floor, walls, data table, drop lines, high-low lines, up/down bars), trendlines, error bars, picture
     fills, `c16:uniqueId`, c15 filtered series, titles, label ranges and `xForSave`, point-label text and
     layouts, ofPie and bubble options, surface band formats, date-axis major and minor time units, layouts,
     `clrMapOvr`, the external data, embedded workbook, colour-style, chart-style and user-shape parts, pivot
     source and formats, print settings and protection.

     The switch was proven by loading every committed deck (22 decks, 81 chart frames: 74 classic, 7 ChartEx)
     through the previous object-tree parser and the adapter and asserting the full `PptxChartData` deeply
     equal (`toStrictEqual`) for every frame; the previous parser was then deleted. Where the adapter reads
     differently it follows the schema and is pinned in `chart-from-neutral.test.ts`: a repeated chart-group
     element (two `c:lineChart` groups) keeps the series of both groups, which the object tree (same-named
     siblings grouped into one array) lost; bare `CT_Boolean` switches (`<c:smooth/>`, `<c:delete/>`) read as
     true; a `<c:marker/>` without `c:symbol` is the automatic marker. The DOM parser rejects parts that are not
     namespace-well-formed, which exposed two pptx writers that wrote `xmlns:c15`/`xmlns:c16` as elements
     (`chart-filtered-series-writer`, `chart-data-labels-range`, fixed); the pptx loader repairs parts those
     writers saved (declares the known Office prefixes, maps `c15:v` back to `c:v`, reports
     `CHART_PREFIX_REPAIRED`) and leaves any other malformed part unloaded.

   - **Trendlines and error bars (2026-10-15).** The neutral model holds them (`chart/model-analysis.ts`:
     `ChartTrendline` with its `c:trendlineLbl`, `ChartErrorBars` with custom `c:plus`/`c:minus` sources;
     parsed by `parse-analysis.ts`, written in schema order by `write-analysis.ts`), so a part with trendlines
     round-trips byte for byte (`chart-gallery.pptx` chart2 left the not-modelled list) and an unreadable
     element is kept in place for positional pairing. The fitting and sizing moved out of the PowerPoint
     renderer into `chart/trendline-fit.ts` (`fitChartTrendline`: every kind, data-space samples, Office's
     equation text, R-squared, coefficients) and `chart/error-bar-amount.ts` (`chartErrorBarAmount`, sample
     variance); `ooxml-ui/pptx` maps their output to pixels (`chart-overlays-trendline-fit.ts`,
     `chart-error-bars.ts`, and the older public `chart-trendlines.ts` entry). pptx builds
     `PptxChartSeries.trendlines` and `.errBars` from the neutral model
     (`pptx/core/utils/chart-from-neutral-analysis.ts`); only the theme-resolved line colour, width, dash and the
     label's manual layout still come from the object tree, paired by position. `PptxChartTrendlineType`,
     `PptxChartErrBarDir`, `PptxChartErrBarType` and `PptxChartErrValType` alias the core types. The full
     `PptxChartData` of all 82 committed chart frames serialised identically before and after, and the
     `ooxml-ui` chart view-model parity hashes did not move. Read as the schema defines them (pinned in
     `chart-from-neutral-analysis.test.ts`): `c:name` is now read (the object-tree reader never found its text),
     a bare `<c:dispRSqr/>` is true and a bare `<c:trendlineLbl/>` is a label. The chart summary carries resolved
     trendlines and error bars (`ChartSummaryTrendline`, `ChartSummaryErrorBars`) and the shared painter draws
     them over bar, column, line and area charts (category positions 1..n) and scatter and bubble charts (their
     x values), with Office's default strokes, so Word and Excel charts show them. The data table and 3-D view
     were already read through the neutral model (`ChartDataTable`, `ChartView3D`, mapped by
     `chart-from-neutral.ts`); `PptxChartDataTable` and `PptxChartView3D` keep their own public field names
     (`showHorzBorder`, `rAngAx`, `hPercent`).

     **Remains (trendlines and error bars):** the deprecated object-tree readers `parseSeriesTrendlines` and
     `parseSeriesErrBars` stay exported from `ooxml-core/pptx` until a major release removes them; the pptx
     trendline and error-bar writers (`chart-trendline-serializer.ts`, `chart-errbars-serializer.ts`) still edit
     the object tree; the shared painter draws no trendline label (equation and R-squared text), nothing over a
     combination chart's further groups or a stacked group, and no x error bars on category charts; xlsx
     regenerates a chart from its `ChartObject` (losing trendlines) only for a type switch to a family the
     operations do not build; the PowerPoint renderer still builds its own view model from `PptxChartData`
     rather than from the chart summary.

   - **Renderer unification (2026-10-17, started).** `docs/chart-renderer-gap-report.md` ranks what the
     PowerPoint renderer draws that core's painter does not (measured by ablation over the 82 committed chart
     parts). Slice 1, data labels: `chart/data-label-text.ts` (flag cascade and text, shared with the PowerPoint
     renderer), `render/summary-labels.ts`, `render/chart-view-labels.ts` and `render/chart-svg-labels.ts`. Word,
     Excel and Visio charts now draw value, category, series name, percentage and bubble-size labels at
     `c:dLblPos` on bar, column, line, scatter, bubble, radar, pie and doughnut charts; per-point deletion and
     custom text, label number formats and the chart-level text properties apply. Not drawn yet: label callout
     shapes and leader lines, manual label layout, "value from cells" text (`c15:datalabelsRange`), rich-text
     label runs, per-point fonts, labels on stock, waterfall, histogram and hierarchy charts, and the xlsx writer
     does not write `c:dLbls` when it regenerates a chart.
     Slices 2 and 3 (2026-10-17), native axes and markers: the PowerPoint range, tick and marker geometry
     implementations now live in strict `chart/axis-range*`, `axis-ticks.ts` and `marker-shape.ts`; PowerPoint
     delegates to these helpers. The shared painter retains native axis IDs, bounds, log and reversed scaling,
     number formats and major and minor gridline presence and styling. Word and PowerPoint's neutral SVG path
     consume these axes directly; Excel retains them in its chart model, also reaching Visio's embedded charts.
     Line, scatter and radar plots honour series and point marker symbols, sizes, visibility and paint. Palette
     edits update live marker paint as well as saved XML. Axis titles, display-unit captions, full tick-mark
     placement, date-axis spacing, series line and smoothing, advanced legend layout and remaining labels, ChartEx
     and 3-D work are still pending.

6. **Port pptx area by area from `XmlObject` to the DOM** (charts and SmartArt first, slides, text and tables
   last), deleting `xml-reorder`/`ordered-xml-merge` as each area is ported, then drop `fast-xml-parser`. This is
   the dominant cost in the whole programme.
   - **Status: charts started.** Classic chart parts are read by the neutral parser and adapted (step 5).
     Charts are saved through `writeChartSpace` and `writeChartExSpace` for the families in section 3.5 (C3, C5).
     The object tree is still parsed for the style fields, ChartEx reading (`utils/chart-cx-parser.ts`), and the
     writers that have not been claimed, so `fast-xml-parser` stays for charts until the `drawingml` models
     resolve theme colours and fonts and the remaining writers move to the neutral writers.
7. **Rendering view-models** (chart 2D and 3D, SmartArt 3D) are UI and render code. The 2D chart painter is core
   (`chart/render`); the 3D charts, the pptx DOM chart renderer and SmartArt 3D stay in `src/ui` (section 5). They
   consume only the neutral models.

Each step ships independently. After each, the `src/core/pptx` tests must pass unchanged (byte-identical output
is the acceptance bar), and the area has its own tests and real parts from `src/core/pptx/__tests__/fixtures`.

### 3.5 Chart steps C1 to C6

- **C1, neutral edit operations and typed intents over `ChartSpace` (done for the classic families).**
  `src/core/chart/edit/` (exported from `ooxml-core/chart`) holds pure operations that return a new `ChartSpace`:
  chart type and grouping switch (`setChartType`, `setChartGrouping`), series add, remove, reorder, rename, values,
  categories and points, chart and axis titles, legend, data labels, axis visibility, title, bounds, units, log
  scale, reverse, number format, ticks, crossing, display units and gridlines (with their outline), series, point
  and element fill, line, colour, transparency and gradient, the built-in style id and palette recolouring. Each
  module states whose semantics it follows: PowerPoint's `builders/sdk/chart-operations.ts` was the checklist and
  decides addressing (flattened series index, first axis of a kind, a position shows a hidden legend, showing empty
  labels turns the value on) and the type switch (keep what the new group element allows, add required children,
  demote `clustered`, drop orphaned axes); Excel decides where PowerPoint is silent (new title and legend markup,
  default axes, pie `varyColors`, 50% doughnut hole, full overlap for stacked bars, the 2.25 pt line a filled series
  gets when it becomes a line series, colour edits keeping transparency). The model gained the pie-of-pie split
  options and date-axis time units for it (`write-space-gaps.test.ts`; the fixture byte round trip is unchanged).
  Intents (`chart/edit/intent.ts`) cover the Format pane fields, including target paint and group options. Proof:
  `pptx/__tests__/integration/chart-neutral-edit-parity.test.ts` makes nine edits on the four bar and line charts of
  `chart-gallery.pptx` both ways (neutral op plus `writeChartSpace`, and the PowerPoint op plus the pptx writers) and
  compares the whole reloaded `PptxChartData`: identical except the listed line width and cap of a bar-to-line
  switch. It found two type-switch differences, fixed in the neutral op (series colour moving to the line,
  `crossBetween` kept). The data sheet maps onto the series operations (`chart/edit/data-sheet.ts`), and
  `chart/edit/new-chart.ts` builds a new chart part through `setChartType`.
  **Remains:** the operations have no per-series combination type (`setChartSeriesChartType`), no trendline or
  error-bar setters, no marker or point-label edits, and no axis title font styling or user shapes; the type switch
  does not build stock, surface or 3-D line charts and makes a combination chart one group; the neutral ops edit only
  the chart part's caches, never an embedded workbook (`chart/workbook` writes that, see C6); the colour style part
  (`colorsN.xml`) is not in `ChartSpace`, so `applyChartColorPalette` recolours series but the caller records the
  palette id.
- **C2, Excel as a client of C1 (done).** `ChartObject.chartSpace` carries the parsed model of a kept part (and of a
  chart `addChart` builds, parsed from the part the summary writer produces), and every edit is a `chart/edit`
  operation on it: `editChart` applies a `ChartEditIntent`, and `updateChart` keeps its `ChartPatch` API by turning
  the difference between the summary the model showed and the patched summary into operations
  (`xlsx/edit/chart-space-patch.ts`, `chart-space-series.ts`: type and grouping, Quick Layout, title, legend, bar
  spacing, series membership, names, references, caches, primary, point and marker paint, effects, axes, element
  fills). The summary the painter reads is derived from the edited model after each edit (`read/chart.ts`
  `chartObjectFromSpace`, with direct formatting read back from the written part, as Word does); structure edits
  rewrite the model's references (`shift-formulas.ts`), and a new chart's caches are refreshed into the model on save
  (`chart-cache.ts`). On save an edited model is written with `writeChartSpace` in the source's declaration style,
  and an unedited one keeps its part byte for byte (`write/drawing.ts` `editedChartPart`). `chart-patch.ts` and the
  DOM half of `write/chart-colors.ts` are gone; only the colour style part (`colorsN.xml`, not in `ChartSpace`) and
  the pivot, protection and colour-map children of `c:chartSpace` and `c:chart` (`chart/keep-unmodelled.ts`) are still
  written beside the model. Excel's Format pane is the shared pane over `chartSpaceFormatSource`, with its intents sent
  to `editChart`, so it gains the border, axis, gridline, data-label, legend-overlay and series options PowerPoint's
  pane shows; Excel's own fill controls keep the chart area, plot area, title, legend and series Fill section.
  Proof: `xlsx/write/chart-save-parity-*.test.ts` saves every classic chart part under `xlsx/__fixtures__` (109: the
  three packaged workbooks and the native Excel parts of the JSON evidence files) after each of 26 edit kinds and
  compares the saved model with the fingerprint of what the DOM path saved (`__fixtures__/chart-save-parity.json`).
  Every non-regenerating kind matches (byte for byte where the DOM path was, else the same model) except where the
  operations are deliberately better, each checked for what the edit asked: a type, grouping or series-count change
  keeps the part instead of regenerating it (the DOM path turned a stock chart into columns), a hidden axis is written
  as `c:delete` where the DOM path ignored axes without direct formatting, inserted rows move every group's and the
  bubble sizes' references (the DOM path moved the first group's), and a point colour its marker already shows is not
  written again as a fill. `src/core/scripts/xlsx-excel-acceptance.ps1` opens formatted and retyped copies of every chart fixture
  in Excel without repair (`xlsx-acceptance-charts.ts`).
  Advanced effects, shape 3D, text, pattern, secondary-axis, analysis, marker and axis-title controls
  now use shared intents within U2's admission limits. Picture relationship allocation remains
  host-owned. A kept part with content the model lacks below the top level (inside the
  plot area, a series or an axis) or a value the parser could not read is not editable: the reader reports it as a
  workbook warning, `updateChart` and `editChart` refuse to change it, the chart commands are disabled, and the part is
  saved from its source bytes with only moved references patched in (`chart/keep-source.ts`,
  `xlsx/edit/chart-space-blocked.test.ts`). The model keeps surface `c:bandFmts` and series and point `c:pictureOptions`
  verbatim (`bandFormatsXml`, `pictureOptionsXml`), the only elements the parser still reported in the committed corpus,
  so none of the 182 distinct classic chart parts of the xlsx fixtures, the Excel evidence files and the pptx and docx
  decks is blocked (three were before: two surface charts and a picture-filled 3-D bar). An edited surface chart keeps
  its band formats and opens in Excel without repair (acceptance script). A combination chart becomes one group on a
  whole-chart type change (the operations' rule). Extended (`cx:`) chart editing is now admitted
  through the separate neutral ChartEx operations and source policy described in C5.
- **C3, pptx saves through the shared writer (admitted families implemented).**
  `chart-neutral-save-plan.ts` plans sparse edits over the stored `ChartSpace`, and
  `writeChartSpaceOverSource` preserves unchanged children and source bytes. Implemented families
  include title text with rich-run preservation, legend and axis fonts/layout, data-label flags,
  value-from-cells ranges and per-point labels, series/point fill and line formatting, gradients,
  appearance/effects/text/3D properties, number/date/display units, series data and structure,
  trendlines, error bars and markers. Formatting-only edits leave embedded workbook cells alone;
  structural data edits allocate fresh workbook ranges. Existing coherent axis-pair reassignment
  and vertical 2D category-axis combination edits preserve document-order selection indices.
  Chrome, protection and stock helpers also use shared source-preserving operations.
  Unclaimed or rejected edits retain an explicit compatibility path. Unsafe group splits,
  unsupported picture relationship changes and wider 3D/XY combinations are not silently admitted.
  The saved-part corpus compares canonical XML and reloaded models; it is implementation parity,
  not a native rendering oracle. The 300-test classic corpus and five sequential native PowerPoint
  edit/reopen/workbook checks pass. See [the delivery ledger](office-shared-parity-delivery.md) for evidence;
  no claim is made that every object-tree serializer can already be deleted.
- **C4, `PptxChartData` as a compatibility adapter (shared projections implemented).** Classic
  facts and supported formatting come from `ChartSpace` through `chart-from-neutral*.ts`; saved
  parts refresh compatibility views with scoped host theme and relationship context. Remaining
  source projections retain unsupported text/spatial details and unmodeled XML. The public model
  remains a compatibility API rather than a neutral alias,
  and the PowerPoint DOM chart painter remains separate from `chart/render`.
- **C5, ChartEx model, writer, edits and consumers (implemented families, migration in progress).**
  `chart/ex` reads/writes the neutral `ChartExSpace` model and carries unsupported source children
  explicitly. Shared operations cover titles/legend/labels, values/categories/names, common fills
  and outlines, series structure, layout/family options and safe local-workbook growth. Word and
  Excel mount the shared ChartEx data sheet and family/axis/point-label controls, with history,
  safe-source admission and save/reopen tests across all six bindings. External data and ambiguous
  hierarchy mappings are refused. ChartEx common formatting uses its own capability source;
  unsupported classic advanced controls are hidden or refused.
  PowerPoint now projects neutral ChartEx facts through a host colour/font bridge, and shared
  creation/save increments cover six authored families and admitted family switches. Neutral
  chart-part grafts reopen with the expected native PowerPoint types in valid gallery decks;
  this does not certify fresh SDK presentation templates. A new six-slide SDK deck failed native
  opening with both the old and new generators, a known template limitation.
  Typed axis/style/point-label operations now serve PowerPoint as well as Word and Excel. The
  style bridge reuses shared DrawingML edits while keeping ChartEx data and layout semantics native
  to its model. **Remains:** raw geography/Pareto creation, wider axis/custom-label authoring and
  removal of the remaining compatibility paths only
  after regression and native checks. The complete Office ChartEx authoring surface is not claimed.

- **C6, Word and Visio chart editing (done for Word and for inserted charts in Visio).** `chart/workbook/` holds the
  embedded-workbook writer moved from pptx (`chart-xlsx-writer`, its cell and validation helpers and the save policy of
  `chart-external-workbook-save`), ported to the shared DOM; pptx forwards to it. Beside it: the standard data layout
  (`layout.ts`), the writes a `ChartSpace` implies (`space-writes.ts`) and a new embedded workbook (`create.ts`).
  `chart/edit/data-sheet.ts` maps the data sheet onto the series operations (formulas follow the table in the standard
  layout) and `chart/edit/new-chart.ts` builds a new chart part. Word inserts, edits, formats, resizes and saves charts
  (`docx/chart-edit.ts`, `docx/chart-save.ts`, `docx/ui/chart-commands.ts`); its Format pane is the shared pane (U2) and
  its charts can be resized with the shared drawing handle. Desktop Word checks the results in the acceptance script. Visio
  edits the data and format of an inserted chart object (`src/ui/src/visio/viewer-chart-editor.ts`) through the embedded
  chart reader and writer of `ooxml-core/xlsx` (`embedded-chart-space.ts`), which apply the `chart/edit` operations.
  **Current limits:** Word/Excel ChartEx editing is admitted as described in C5. Native chart style
  and colour-part authoring remains incomplete; curated style presets do not replace it. Visio's
  chart workflow edits eligible embedded Excel chart objects through the shared data sheet and
  Format pane, regenerating the object's preview. Unsupported source parts are read-only. Arbitrary
  OLE activation, external workbooks and native preview pixel fidelity are not established.

### 3.6 Chart UI steps U1 to U5

The labels U3 to U5 are assigned here; the remaining chart UI work was previously listed under the shared Format pane.

- **U1, shared Format pane model and element (done).** `office-ui-chart-format-pane` (`src/ui/src/chart/format-pane*.ts`)
  draws PowerPoint's Format pane from a neutral model over a `ChartFormatSource`. `chartSpaceFormatSource` reads a
  `ChartSpace`, and each field is emitted as a typed `ChartEditIntent` (`chart/edit/intent.ts`, with the new target paint
  and group options operations).
- **U2, products mount the pane (implemented with capabilities and limits).** PowerPoint,
  Excel, Word and embedded Visio charts use `office-ui-chart-format-pane`. Core owns the neutral
  operations and resolved appearance conversion; products supply source/relationship/history
  adapters. Shared controls cover effects, shape 3D, text, pattern fills, outline joins/ends,
  markers, trendlines/error bars, axis titles and admitted secondary-axis moves. Excel keeps its
  richer gradient editor in a slot and uses the shared pattern editor. Picture media allocation
  remains host-owned. Secondary authoring retains the first series as the primary render anchor;
  horizontal combinations, stacked multi-series and unknown-extension splits are refused. The
  pane edits one trendline and preserves custom error ranges without a range editor. ChartEx
  exposes only its supported controls. Browser cases verify Word/Excel controls, history and
  save/reopen across six bindings and the embedded Visio workflow; native Word/Excel checks cover
  admitted ChartEx saves. Exact text, effects and 3D rendering still need broader native evidence.

- **U3, shared chart data sheet (`office-ui-chart-data-sheet`, `src/ui/src/chart/data-sheet.ts`).** One element for the
  PowerPoint bindings, Word and Visio. Excel edits its cells directly. The tokens for its lengths are the `--office-*` tokens.
- **U4, one insert-chart picker (`office-ui-chart-picker`, `src/ui/src/dialog/chart-picker.ts`).** Used by PowerPoint's insert
  dialog, Word, Excel and Visio.
- **U5, shared chart design galleries (`src/ui/src/chart/design-galleries.ts`).** Chart Design galleries (colours, styles, quick
  layouts) shared by the Word, Excel and PowerPoint ribbons.

### 3.7 SmartArt steps S1 to S5

- **S1, neutral DrawingML text model (done).** `drawingml/text-model.ts`, with `read-text-body.ts` and `write-text-body.ts`;
  pinned equal to the pptx object-tree parsers by `pptx/__tests__/integration/smartart-text-neutral-parity.test.ts`.
- **S2, node text projection into `diagram/text` (done).** `projectDiagramNodeText` is generic over a styler; pptx keeps
  `smartart-node-text-projection` as its `TextStyle` styler over its own parsers, and `readDiagramNodeParagraphs` plus the
  neutral styler serve Word and Excel. **Remains:** the pptx styler and object-tree paragraph parser stay until pptx reads
  SmartArt text through the neutral model (step 2).
- **S3, neutral drawing view and edit operations (done).** `diagram/drawing-view.ts` and `drawing-resolve.ts` (step 4), and
  `diagram/edit` (node commands, tree order, outline levels, reflow, inserts). The insert picker and the SmartArt Design command
  list are shared (`design-commands.ts`).
- **S4, part writers (done for data parts).** `diagram/write`: `data.ts` (a SmartArt created in memory), `data-edit.ts`
  (`writeEditedDiagramData`, an existing data part) and `inserted.ts` (an inserted diagram's data part). **Remains:** the
  full source-cache round-trip writer and layout fabrication stay in pptx (step 4). Colour and quick-style fabrication
  now live in `diagram/write/fabricated-styles.ts`, with the PowerPoint path as a compatibility facade. The neutral
  source-backed formatter patches definition fields while preserving unknown XML. `writeGeneratedDiagramDrawing`
  serializes generated 2D drawings and returns presentation GUIDs a host can associate with its data model; PowerPoint's
  ordinary 2D preset shape properties use the same serializer. It refuses 3D input rather than flattening it silently.
  Word and Excel drop stale cached drawing parts on edit and now attach a generated replacement when the ordered
  interpreter and 2D painter admit the diagram. The shared writer adds presentation GUID associations; each host adds
  the drawing relationship and content type. A private extension records the generated approximation and a SHA-256 of
  its drawing bytes. Loaders label matching caches `generated-drawing`; a native rewrite invalidates that status.
  Unsupported, partial and spatial/3D interpretations stay uncached and retain the explicit fallback.
- **S5, Word and Excel SmartArt editing (done, with the shared UI).** Word edits and inserts SmartArt (`docx/diagram-edit.ts`,
  `docx/diagram-save.ts`); Excel does the same through `xlsx/edit/smart-art.ts` and `xlsx/write/smart-art.ts`. Both use the
  shared `office-ui-smartart-picker` (`src/ui/src/smartart/smartart-picker.ts`) and `office-ui-smartart-text-pane`
  (`src/ui/src/smartart/`, with the `SmartArtTextPaneController` both editors wire up). PowerPoint's bindings use the same text
  pane through the `pptx-ui-smartart-text-pane` alias (`src/ui/src/pptx/web-components/smartart-text-pane.ts`), which adds
  PowerPoint's inspector tools (append, move, sub-item, remove, per-node style) and applies every request through
  `render/smartart-text-pane-model.ts` and the pptx SmartArt facades.
  Word and Excel now mount the same SmartArt Styles and Change Colors galleries as PowerPoint, over the shared catalogue:
  three style intensities and five theme-accent colour schemes. `formatDocxDiagram` and `EditSession.formatSmartArt` update
  source definitions, keep node overrides and unchanged rich text, invalidate stale drawings, and save the selected
  definition ids. History retains the matching definitions, including after formatting an inserted diagram.
  **Remains:** rich text authoring in Excel's text pane, the wider native Office style catalogue, and 3D formatting controls.

  Acceptance on 2026-10-26: the shared formatting builder and
  `src/core/scripts/docx/smartart-format-acceptance.ps1` produced and reopened six documents, three combinations each in desktop
  Word 16 and Excel 16. Native node texts, colour ids and quick-style ids matched; Excel produced no repair log. These
  checks also verify native semantic-node drawing associations. Both host loaders independently verify the generated
  caches before native opening. These checks establish supported edit/save/reopen behavior, not pixel-identical effect or layout rendering. Unit tests cover
  retained extension XML, theme references, direct node fills, rich text, undo/redo, and neutral generated cache round trips.

Wave 6 did not change the Office evidence below; it is kept because the save decisions depend on it.

#### 3.7.1 Word and an edited data part (evidence, 2026-10-23)

Desktop Word 365 (16.0, `WINWORD.EXE` driven through COM) was given a SmartArt it had inserted itself (Basic Block List,
nodes Alpha to Epsilon; `docx/__fixtures__/smartart-word.docx`) with the data part changed behind its back. Each variant was
opened read-only, its `SmartArt.AllNodes` texts read, the text of the picture Word draws for it taken from
`InlineShape.Range.EnhMetaFileBits` (the EMF records), and the file saved again so the drawing Word writes could be counted.
Results:

- data text changed, stale cached drawing kept: Word reports and draws the new text ("Beta EDITED") and saves a drawing with the new text;
- a node added to the data, stale five-shape drawing kept: Word draws the sixth node and saves a six-shape drawing;
- the same with the drawing part, its relationship, content type and `dsp:dataModelExt` removed; with only the `dsp:dataModelExt`
  removed (drawing part left dangling); with every `pres` point and `presOf`/`presParOf` connection removed as well (with and without
  the drawing): every variant opens without repair, Word draws all six nodes and saves a full presentation tree (12 `pres` points)
  and a six-shape drawing;
- the data unchanged but the cached drawing's text altered ("Gamma" to "Cacheonly"): Word draws "Gamma". Word never paints the
  cached drawing; it always lays the diagram out from the data and layout parts;
- stub definitions (a `colorsDef`/`styleDef` holding only a built-in `uniqueId`): Word keeps the stub and draws uncoloured shapes; a
  stub `layoutDef` naming `layout/process1` made Word fall back to Basic Block List. Word does not substitute its built-in copy for a
  known `uniqueId`, so a new diagram must carry complete definitions.

Decision: an edited diagram drops its cached drawing (relationship, part, content type and the `dsp:dataModelExt` link) instead of
regenerating it, and keeps the presentation points of the nodes that remain; Word recomputes the layout. A new diagram carries a
built-in layout definition (`ooxml-core/diagram/builtin`) and Word's default `colors/accent1_2` and `quickstyle/simple1`
definitions, with no `pres` points and no drawing. The probe is `src/core/scripts/docx/smartart-word-probe.ps1` (variants from
`smartart-word-variants.mjs`; its output on 2026-10-23: `cache-only: nodes [Alpha, Beta, Gamma, Delta, Epsilon]; drawn [Alpha, Beta,
Gamma]`, `stale-add: nodes [..., Zeta]; drawn [Alpha, Beta, Gamma, Zeta]; re-saved drawing 6 shapes`, `nodrawing-nopres: ...
re-saved drawing 6 shapes`, `stale-text: drawn [Alpha, Beta, EDITED, Gamma]`). The acceptance check for the writer is
`src/core/scripts/docx-word-acceptance.ps1` (with `docx-acceptance-build.ts`), which saves the fixtures untouched, after text, add,
remove, move, promote and demote edits, and with four presets inserted, and checks in Word that each opens without repair, reports
the expected node texts and draws every node's text. Last run 2026-10-23: all 6 documents passed.

#### 3.7.2 Excel and an edited data part (evidence, 2026-10-23)

Desktop Excel 365 (16.0, `EXCEL.EXE` driven through COM) was given its own SmartArt (Basic Block List, Plan, Build, Ship;
`xlsx/__fixtures__/excel-smartart.xlsx`) with the data part or cached drawing changed behind its back, the variants of
`src/core/scripts/docx/smartart-word-variants.mjs` ported to the workbook (`src/core/scripts/xlsx/smartart-excel-variants.mjs`). Each was
opened read-only, its `SmartArt.AllNodes` texts read, the picture Excel draws taken from the clipboard as an EMF
(`Shape.CopyPicture`) and searched for the probe words, and the workbook saved again. The probe is
`src/core/scripts/xlsx/smartart-excel-probe.ps1`; its output on 2026-10-23:

- `cache-only` (data unchanged, the cached drawing's "Ship" changed to "Cacheonly"): nodes [Plan, Build, Ship]; drawn [Plan, Build,
  Ship]; re-saved drawing 3 shapes [Plan|Build|Ship]. Excel never paints the cached drawing.
- `stale-text` (data text changed, stale drawing kept): nodes [Plan, Build EDITED, Ship]; drawn [Plan, Build, EDITED, Ship]; re-saved
  [Plan|Build EDITED|Ship].
- `stale-add`, `stale-nopres`, `dangling-ext`, `nodrawing-add` and `nodrawing-nopres` (a node added; with the stale drawing, without
  the `dsp:dataModelExt` link, without the drawing part, relationship and content type, and without every presentation point): every
  variant opens, Excel lists and draws all four nodes and saves a four-shape drawing [Plan|Build|Ship|Zeta].

Decision: Excel behaves like Word, so an edited diagram drops its cached drawing (relationship and part) and the `dsp:dataModelExt`
link, keeps the presentation points of the nodes that remain, and Excel recomputes the layout; a new diagram carries the built-in
layout definition and Word's default colour and quick-style definitions, with no `pres` points and no drawing. The acceptance check is
`src/core/scripts/xlsx/smartart-excel-acceptance.ps1` (with `smartart-acceptance-build.ts`): it saves the fixture untouched, after text,
add, move and remove edits, after promote and demote, after Change Layout, and a new workbook with four presets inserted, and checks in
Excel that each opens without repair, lists the expected `AllNodes` texts and layout and draws them. Last run 2026-10-23: all 6
workbooks passed.

### 3.8 Model gaps and parity harnesses in one place

| Harness                           | What it proves                                                | Where                                                     |
| --------------------------------- | ------------------------------------------------------------- | --------------------------------------------------------- |
| Chart neutral parity (2026-10-08) | pptx object tree and neutral parser agree on 74 classic parts | `pptx/__tests__/integration/chart-neutral-parity.test.ts` |
| Chart neutral edit parity (C1)    | nine edits both ways on four charts agree                     | `chart-neutral-edit-parity.test.ts`                       |
| Chart save parity (C3)            | 13 edits on 75 parts saved both ways                          | `chart-save-parity.test.ts`                               |
| Chart ChartEx parity (C5)         | every committed chartex part agrees                           | `chart-ex-neutral-parity.test.ts`                         |
| ChartEx save parity (C5)          | nine edits on 15 chartex parts                                | `chart-ex-save-parity.test.ts`                            |
| Excel chart save parity (C2)      | 26 edit kinds on 109 classic parts                            | `xlsx/write/chart-save-parity-*.test.ts`                  |
| SmartArt text parity (S1)         | every text body of the committed fixtures                     | `smartart-text-neutral-parity.test.ts`                    |
| SmartArt drawing parity (S3)      | 371 drawing parts, 3014 shapes                                | `smartart-drawing-shape-neutral-parity.test.ts`           |
| Choose walkers corpus (step 2)    | every walker over 353 layout definitions                      | `smartart-choose-walkers-corpus.test.ts`                  |

## 4. Risks

- **Two XML models for a long time.** Mitigation: neutral attribute readers and `PartHost` instead of a bridge; do not build an
  object-tree view over the DOM.
- **Behavioural drift in the pptx adapter.** Mitigation: the pptx suite is the oracle (16 minutes); add parity tests that parse the
  same part with both parsers on the committed decks before switching a pptx call site.
- **Colour fidelity.** pptx resolution includes `phClr`, effect colour transforms and a quick-style and colour-list interplay. The
  `diagram` resolver covers `srgb`, `scheme`, `sys`, `prst`, `scrgb`, `hsl` and the common transforms (`alpha`, `lumMod`, `lumOff`,
  `tint`, `shade`, `satMod`) and reports the rest as unresolved rather than guessing.
- **Layout engine size.** About 60 engine files and about 150 interpreter files share the types barrel; moving them piecemeal created
  circular imports. Types moved first, then the engine, then the interpreters, one PR each.
- **SmartArt regeneration has explicit limits.** Word and Excel use the shared neutral layout
  interpreter for uncached and edited diagrams, with an approximate fallback for unsupported or
  partial layouts. Admitted 2D results gain generated source caches; complex text, folding and
  spatial 3D decline cache generation. Native Office can regenerate its own cache, but that does
  not establish that the viewer's geometry, typography or effects match Office (section 3.7).
- **Strictness gap.** `pptx` compiles with relaxed flags; code moved into a neutral area must compile strict (`exactOptionalPropertyTypes`,
  `noUncheckedIndexedAccess`). Budget for bounds-checked access as in the `geometry` move.
- **Bundles.** The pptx bundle inlines shared areas; check bundle size and that `diagram` and `chart` stay tree-shakeable.

## 5. What stays format-specific

- `pptx`: slides, masters, layouts, notes, animation, transitions, presentation properties, the fluent builder SDK, converters, CLI,
  signatures, `a:tbl`, `p:graphicFrame` handling, SmartArt-to-slide-element decomposition, and the save pipeline that writes the diagram
  parts into a presentation.
- `docx`: WordprocessingML (`w:p`, `w:r`, `w:tbl`, numbering, styles, revisions), the anchor, inline and wrap host markup, and the OPC
  wiring that adds a diagram's or a chart's parts to a package.
- `xlsx`: sheets, cells, formulas, the drawing anchor (`xdr:graphicFrame`) and the workbook's own chart and SmartArt hosts.
- Rendering code (three.js scenes, HTML painters, the pptx chart view-models) is UI and stays in `src/ui`; the neutral SVG chart painter is
  the exception and lives in `chart/render`.

## 6. Shared UI: a second package (not pursued as written)

This section proposed a `packages/*` layout (`packages/ooxml-core`, `packages/ooxml-ui`, `packages/office-collab`), `v<version>` tags and a
`pptx-ui-*` to `office-ui-*` migration across separate viewer repositories. None of it shipped in that form and it should not be followed.
The repository is a Bun workspace with the library in `src/core`, the DOM-only `ooxml-ui` in `src/ui`, tags `ooxml-core@<version>` and
`ooxml-ui@<version>`, collaboration as the `collab` area, and the viewers under `viewers/`. See `AGENTS.md` and `src/ui/README.md` for the
current layout and [ooxml-ui-plan.md](ooxml-ui-plan.md) for the UI package's history.

## 7. First slice delivered (history)

See `PROVENANCE.md` ("diagram area"): `src/core/diagram` and `src/core/docx/diagram*.ts`. Everything in steps 2 to 6 above has since moved
or is listed under its "Remains" entry.

### Classic chart read adapter follow-up (2026-10-26)

PowerPoint series solid fills, fill and outline gradients, line widths, preset dashes and no-fill
switches now project from neutral DrawingML fields. Chart/plot/title/legend/axis/gridline/series format-pane fills, outlines and ordinary outer-shadow/glow/soft-edge effects also project from the neutral models; spatial and text-specific source details retain their compatibility adapters. Trendline and error-bar lines, trendline label
layout values, marker and point shape overrides, floor/wall formatting, up/down bars, and drop/high-low lines use the same neutral projection. Theme aliases resolve through
the deck's existing color-map context; authored tokens remain in `ChartSpace` for saving.

`refreshClassicChartFromPart` rebuilds compatibility data from a saved classic part with the current
theme and host relationship context. It refreshes chart groups, axis topology and series identities,
instead of layering a new topology over stale resolved snapshots. Reordered axes retain styling by
axis identity and the retained snapshot is not mutated by the projection.

Remaining source-tree compatibility reads include resolved label/axis text styles,
complex text effects and spatial properties, picture details/relationships, extension payloads, filtered
series, external style/workbook/user-shape parts and unsupported layouts. This checkpoint does not
claim that every classic chart read is DOM-only. The 81-frame chart projection corpus and focused
adapter regressions remain the gate for deleting each compatibility reader.
