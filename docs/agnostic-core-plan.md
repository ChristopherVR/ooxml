# Plan: making ooxml-core format-agnostic

Status: proposal plus first slice (the `diagram` area), 2026-10-02. Counts below were taken from `origin/main` (`c3e0979`) with `find`/`grep`; they are approximate and describe coupling, not quality.

Goal: everything that is not specific to one Office format lives in a format-neutral area with a neutral public model, and `docx`, `pptx` (later `xlsx`) are thin adapters over it. Word can then use what only PowerPoint has today: SmartArt first, then charts, then DrawingML primitives (fills, lines, effects, text bodies, theme).

## 1. Where the shared logic lives today (`src/pptx`)

`src/pptx/core` is 79,630 lines of SmartArt, chart and DrawingML code mixed with slide code in flat folders (`core/utils` ~1,700 files, `core/core/runtime`, `core/core/builders`, `core/types`). Only file names separate generic DrawingML from slide logic.

### 1.1 The coupling, in four layers

| Layer                               | What it is                                                                                                                                                                                                                                                                                                       | Why it is not portable as is                                                                                                                                                                                            |
| ----------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| XML model                           | `XmlObject` (`types/common.ts`): `fast-xml-parser` trees, attributes as `@_name`, same-name siblings grouped into arrays, prefixes in keys (`a:solidFill`) looked up by local name through `xmlLookupService`.                                                                                                   | docx and `src/xml` use a namespace-aware DOM (`@xmldom/xmldom`). Every pptx parser reads `node['@_x']`. 87 of 162 `utils/chart-*` files and nearly every smartart file take `XmlObject`.                                |
| Handler runtime (`this`)            | A mixin chain `PptxHandlerRuntime` (`core/core/runtime/*`): `parseSmartArtDrawingShapes`, `parseDrawingShape`, `parseSmartArtQuickStyle`, chart parsing all call `this.parseColor`, `this.resolveThemeFillRef`, `this.readXmlPartByRelationshipId(slidePath, relId)`, `this.compatibilityService.reportWarning`. | The base class owns the JSZip instance, the theme maps, the rels maps and the slide being loaded. Colour, gradient, shadow and theme resolution are methods of it, so a shape parser cannot run without a presentation. |
| Types                               | `PptxSmartArtDrawingShape extends PptxCustomPathProperties`, `PptxSmartArtData` embeds `TextSegment`, `Pptx3DScene`, `PptxSmartArtChrome`; `PptxElement`/`PptxSlide` appear in `smartart-decompose-*` (SmartArt to slide elements) and the chart builders.                                                       | Public types are named and shaped for PowerPoint; the pure data (nodes, connections, colours, quick style) is tangled with slide editing fields.                                                                        |
| Pure algorithms (no XML, no `this`) | `smartart-engine/` (layout algorithms, constraint solver, text fit, font metrics tables), `smartart-layout-interpreter-*`, `smartart-hierarchy-*`, `smartart-editing-*` (reflow, node ops) over typed `PptxSmartArt*` models; chart data utils; `color/`, `geometry/` (already moved).                           | These are the cleanest to move, but they import the pptx types barrel (`from '../types'`). Moving them needs the typed model to move first.                                                                             |

### 1.2 SmartArt map

- Part model and XML parsing: `core/builders/PptxSmartArtParser.ts` (data points, connections), `core/runtime/PptxHandlerRuntimeSmartArtParsing.ts` (connections with transition labels, quick style, cached drawing shapes, `parseDrawingShape`), `SmartArtXmlUtils`, `utils/smartart-data-model-attributes.ts`, `smartart-connector-labels.ts`, `diagram-relationship-ids.ts`, `core/runtime/smartart-drawing-part.ts` (resolves `dsp:drawing` through `dsp:dataModelExt` or the data part's rels), `smartart-colors-builder`, `smartart-style-label-refs`, `smartart-drawing-shape-{style,3d,font-ref}`, `smartart-text-paragraphs`, `smartart-text-style-resolution`.
- Types: `types/smart-art*.ts` (`PptxSmartArtNode`, `PptxSmartArtConnection`, `PptxSmartArtDrawingShape`, layout definition, colour/quick-style definitions, constraint rules, chrome).
- Layout engine and interpreters: `utils/smartart-engine/*` (~60 files, own ordered-XML layout-definition parser), `utils/smartart-layout-*`, `smartart-hierarchy-*`, `smartart-constraint-*`, `smartart-decompose*`.
- Editing and writing (generation of parts from scratch, "fabrication"): `core/runtime/smartart-fabrication-*`, `smartart-xml-builders`, `PptxHandlerRuntimeSaveSmartArtFabrication`, `utils/smartart-editing*`.
- PPTX-only glue: `PptxGraphicFrameParser` (finds `dgm:relIds` in a `p:graphicFrame`), `SmartArtElementProcessor` (converter), `smartart-decompose` (explode SmartArt into slide shapes), save pipeline.

Coupling verdict: the **data shapes** (point, connection, relIds, custom layout, drawing geometry/text) are thin and only reach `XmlObject` through attribute reads. The **cached drawing shape parser** is wired to the runtime for colour/gradient/theme resolution. The **layout engine** is XML-free after its input is parsed, but its types come from the pptx barrel. The **writer/fabrication** code is object-tree specific.

### 1.3 Charts map

`types/chart*.ts` (~2,000 lines of model), `utils/chart-*` (162 files, ~87 import `XmlObject`; 10 are pure), `core/runtime/PptxHandlerRuntimeChart*` (parsing 1,105 lines, colour style, external data, user shapes), `builders/sdk/chart-operations.ts` (1,064), `chart-xlsx-parser` (embedded workbook), chartex (`chart-cx-*`), colour/style parts (`chart-color-style-writer`, `chart-style-definition`). Parsing starts from a `p:graphicFrame` on a slide; the part loader is `readXmlPartByRelationshipId(slidePath, relId)`. Word needs the same part (`c:chart r:id` inside `w:drawing`) and the same `c:chartSpace` model.

### 1.4 DrawingML primitives map

Fills, lines, effects and text: `PptxShapeStyleExtractor`, `PptxGradientStyleCodec`, `PptxColorStyleCodec`, `PptxColorTransformCodec`, `PptxEffectDagExtractor`, `effect-dag-*`, `shape-style-line-helpers`, `drawing-fill-xml`, `body-properties-parser`, `paragraph-properties-parser`, `text-run-properties-parser`, `drawing-line-dash`, `gradient-angle`; theme in `PptxHandlerRuntimeTheme*` (loading, format scheme, ref resolution, `phClr` substitution) and `types/theme.ts`; custom geometry (`custom-geometry*`, left in pptx by the earlier `geometry` extraction). `docx` already has its own small theme model (`theme.ts`, `theme-color.ts`) and its text runs use WordprocessingML, so only the DrawingML text body (inside shapes, charts, diagrams) is shared.

## 2. Target areas

Every area is a subpath of the single package, strict TypeScript, written against the shared `xml` DOM and `opc`. Neutral public types never mention `Pptx*`, `Docx*`, `XmlObject`, a slide or a document.

| Area        | Contents                                                                                                                                                                                                                                                            | Depends on                          |
| ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------- |
| `drawingml` | `DrawingColor` (srgb/scheme/sys/prst/hsl/scrgb + transform list) and its resolver, `Fill`, `Line`, `Effect` models, `TextBody`/paragraph/run models, xfrm, `ThemeLike` (colour scheme, font scheme, format scheme) and `phClr` resolution. DOM parsers and writers. | `xml`, `color`, `geometry`, `units` |
| `diagram`   | The SmartArt part model: data model, layout definition header, colours, quick style, cached `dsp:drawing` shape tree; the relationship resolver; later the layout engine and fabrication.                                                                           | `drawingml`, `xml`, `opc`           |
| `chart`     | `c:chartSpace` model, parse/serialize, `cx:` chartex, colour/style parts, embedded workbook reader, a renderer-neutral series/axis view-model.                                                                                                                      | `drawingml`, `xml`, `opc`           |
| `theme`     | Possibly a sub-module of `drawingml` rather than its own area; `docx` and `pptx` both feed it.                                                                                                                                                                      | `drawingml`                         |

`diagram` ships first (this PR) with a **provisional** private colour/fill reader inside it (`diagram/drawing-color.ts` and `drawing-fill.ts`) because `drawingml` does not exist yet. These two files are written against the DOM and a `DrawingColor` type, with no diagram knowledge, precisely so they move to `drawingml` unchanged in step 3.

### 2.1 Adapters

A neutral area exposes **parsers over DOM elements or part text plus small injected interfaces**; a format area supplies the host:

- `PartHost`: `readText(partName)`, `resolve(basePart, target)`, relationships by part. The pptx adapter implements it over its JSZip and `slideRelationships`; the docx adapter over the document package. This replaces `this.readXmlPartByRelationshipId(slidePath, relId)`.
- `ColorResolver`: `(DrawingColor) => hex | undefined`. pptx binds its theme maps; docx binds `src/docx/theme.ts`.
- Diagnostics: parsers return `issues: {code, message}[]` instead of calling a compatibility service.
- pptx adapter also converts between the neutral model and its legacy `Pptx*` types (aliases where shapes are identical, mappers where not) so the public pptx API does not change.
- docx adapter owns the host markup (`w:drawing`/`wp:inline`/`wp:anchor`, extent, wrap, alt text) and calls the neutral area for the graphic part.

For code that still takes `XmlObject`, a neutral parser may be driven through a **neutral attribute reader** (`(name) => string | undefined`) instead of a DOM element; pptx passes `(n) => node['@_'+n]`. That is how this slice shares the connection, custom-layout and relationship-id parsers without converting pptx to the DOM first. A full object-tree view of a DOM is deliberately not built: it would keep two XML models alive.

## 3. Migration order

1. **`diagram` part model and cached drawing (done in this PR).** Neutral data model, connection/custom-layout/relIds parsers, relationship resolver, colours and quick-style summaries, cached drawing shape tree, loader. pptx re-imports the neutral parsers and types. Add docx `w:drawing` -> `DocxDiagram`.
2. **`diagram` model types and layout engine.** Move `PptxSmartArtNode`/layout-definition/colour/quick-style types to neutral names with `Pptx*` aliases, then the XML-free code: `smartart-engine`, interpreters, hierarchy, constraint solver, editing reflow. The engine's own `layout-def-parse` (ordered XML) becomes a DOM parser. Needs `Pptx3D*` and `TextSegment` decoupled first (a `drawingml` text model).
3. **`drawingml` extraction.** Lift the provisional colour/fill code from `diagram`, add line/effect/text-body/theme-format models and their DOM parsers; pptx codecs (`PptxColorStyleCodec`, `PptxGradientStyleCodec`, `PptxEffectDagExtractor`) become object-tree shims that delegate or are replaced once their callers are ported. docx gets `w:drawing` shape fills and wps text boxes on top of it.
4. **Drawing shape parsing off the runtime.** `parseDrawingShape` takes a `ColorResolver` and `drawingml` parsers instead of `this`; pptx's `parseSmartArtDrawingShapes` becomes a mapper over `diagram.parseDiagramDrawing`. Requires pptx fixtures (smartart-gallery, corpus) as a parity test.
5. **`chart` model and parsers.** Move `types/chart*.ts` as neutral `Chart*` types (aliases in pptx), then DOM parsers for chartSpace; pptx object-tree chart code stays as a shim until step 6. docx adds `c:chart` in `w:drawing`, loading the chart part and its embedded workbook.
6. **Port pptx area by area from `XmlObject` to the DOM** (charts and SmartArt first, slides, text and tables last), deleting `xml-reorder`/`ordered-xml-merge` as each area is ported, then drop `fast-xml-parser`. This is the dominant cost in the whole programme and is the existing plan in `docx-viewer/docs/ooxml-core-plan.md`.
7. **Rendering view-models** (chart 2D/3D, SmartArt 3D) are UI/render code and stay in the viewer repos or the UI package (section 5); they consume only the neutral models.

Each step ships independently; after each, `src/pptx` tests must pass unchanged (byte-identical output is the acceptance bar) and the area has its own tests plus real parts from `src/pptx/__tests__/fixtures`.

## 4. Risks

- **Two XML models for a long time.** Mitigation: neutral attribute readers and `PartHost` instead of a bridge; do not build an object-tree view over the DOM.
- **Behavioural drift in the pptx adapter.** Mitigation: the pptx suite is the oracle (16 minutes); add parity tests that parse the same part with both parsers on the committed decks before switching a pptx call site.
- **Colour fidelity.** pptx resolution includes `phClr`, effect colour transforms and a quick-style/colour-list interplay; the provisional `diagram` resolver covers `srgb`, `scheme`, `sys`, `prst`, `scrgb`, `hsl` and the common transforms (`alpha`, `lumMod`, `lumOff`, `tint`, `shade`, `satMod`) and reports the rest as unresolved rather than guessing.
- **Layout engine size.** ~60 engine files and ~150 interpreter files share the types barrel; moving them piecemeal creates circular imports. Move types first, then engine, then interpreters, in one PR each.
- **Word has no regeneration path.** The cached `dsp:drawing` is optional (Word may omit it, and files from other producers do). Until the layout engine is neutral, a Word SmartArt without a cached drawing can only be preserved and shown as a placeholder, which the model reports honestly.
- **Strictness gap.** `pptx` compiles with relaxed flags; code moved into a neutral area must compile strict (`exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`). Budget for bounds-checked access as in the `geometry` move.
- **Bundles.** The pptx bundle inlines shared areas; check bundle size and that `diagram` stays tree-shakeable.

## 5. What stays format-specific

- `pptx`: slides, masters, layouts, notes, animation, transitions, presentation properties, the fluent builder SDK, converters, CLI, signatures, `a:tbl`, `p:graphicFrame` handling, SmartArt-to-slide-element decomposition, the save pipeline that writes the diagram parts into a presentation.
- `docx`: WordprocessingML (`w:p`, `w:r`, `w:tbl`, numbering, styles, revisions), the anchor/inline/wrap host markup, the OPC wiring that adds a diagram's parts to a package.
- Rendering code (three.js scenes, chart view-models, HTML/SVG painters) is UI and does not enter the logic package.

## 6. Shared UI: a second package in this repository

Today `docx-viewer` and `pptx-viewer` each hold a web-component layer. Identical or near-identical pieces: the ribbon elements (`ribbon-command`, `ribbon-group`, `ribbon-gallery`, `ribbon-section`, `ribbon-toggle`, `ribbon-view`, `ribbon-draw`), `select`, `select-ribbon`, `checkbox`, `search-field`, `theme-editor`, host and control styles, icon sets (`pptx-viewer/packages/shared/src/web-components`, 55 files), the i18n catalogue plumbing (`shared/src/i18n`, `packages/locales`: en/de/es/fr), and collaboration (docx: `collaboration-*.ts` on `prosemirror-collab`; pptx: `shared/src/render/collaboration-*` on Yjs, not the same protocol, see below).

The core package stays logic-only: **no DOM, no custom elements, no CSS, no framework code**. The UI goes to a different published package from the same repository.

### 6.1 Layout (Bun workspaces)

```
ooxml-core/                  repository root (private workspace root, no package published from here)
  packages/
    ooxml-core/              ooxml-core   logic only (today's package, moved down one level)
    ooxml-ui/               ooxml-ui    shared web components, styles, icons, i18n keys
    office-collab/           @christophervr/office-collab (optional, later) sync protocol + Yjs/ProseMirror adapters
  package.json               { "private": true, "workspaces": ["packages/*"] }
```

- Dependency direction: `ooxml-ui` may import `ooxml-core` types; `ooxml-core` never imports `ooxml-ui`. A CI check (like `check-shared.mjs` in docx-viewer) enforces it.
- `ooxml-ui` exports subpaths per element family (`/ribbon`, `/select`, `/theme-editor`, `/i18n`), each registering its custom elements idempotently, SSR-safe (pptx already has `ssr.test.ts` and `registration-contract.test.ts` to port). Locale catalogues stay data (`/i18n/en` ...), like `pptx-viewer-locales`.
- Viewers depend on both packages; the framework bindings (React, Vue, Angular, Svelte, Solid, vanilla) stay in the viewer repositories because they are lifecycle adapters of the viewer's own element.
- Collaboration: the two products use different engines (ProseMirror steps vs Yjs). The shared part is the transport-neutral session/identity/authority API; extract it to `office-collab` only after both sides agree on one session interface. Until then it stays in the viewers. Yjs is an optional peer, as in pptx today.

### 6.2 Release and versioning

pptx-viewer publishes several packages with **independent semver** (`pptx-viewer-core` 4.9.1, bindings 4.16.0, CLI 2.33.0) driven by a `release-plan.json` computed from conventional commits and per-package `includePaths`, tags `<name>@<version>`, git-cliff changelogs, and its `shared` and `locales` packages are private and bundled into each binding. docx-viewer publishes six packages at one shared version with `v<version>` tags. ooxml-core today publishes one package with `v<version>` tags.

Proposal: keep **independent versions with per-package tags** (`ooxml-core@x.y.z`, `ooxml-ui@x.y.z`) using the pptx-viewer release-plan mechanism (affected-package detection by path, conventional commits, git-cliff per package), because the logic package changes far more often than the UI and the UI must not force logic releases (or vice versa). `ooxml-ui` declares `ooxml-core` as a peer dependency range. npm provenance and trusted publishing per package, as today. Existing `v<version>` tags stay valid for `ooxml-core` 0.x; the first release from the workspace switches to `ooxml-core@<version>` tags.

### 6.3 Migration steps and what is needed

1. Convert the repo to workspaces: move the current package to `packages/ooxml-core` (history preserved with `git mv`), root `package.json` becomes private; update CI paths, `tsup`/`tsdown` configs, `files`, the pack smoke test, the `OOXML_CORE_REF` pinning used by docx-viewer CI, and docx-viewer's `file:../ooxml-core` dependency (becomes `file:../ooxml-core/packages/ooxml-core`). Do this alone, behaviour unchanged.
2. Add `packages/ooxml-ui` with an empty build, lint and test pipeline, its own release-plan entry, and a dependency-direction check.
3. Move the **already identical** elements first: `checkbox`, `select*`, `search-field`, `ribbon-command`, `ribbon-group`, with their tests, from pptx-viewer `packages/shared/src/web-components`; docx-viewer swaps its copies. Provenance recorded in `PROVENANCE.md` per module.
4. Move styles and icons, then the ribbon section/gallery/view elements (pptx-specific commands are injected, not imported), then i18n key plumbing and the en/de/es/fr catalogues for the shared strings (Word-specific and PowerPoint-specific strings stay in each viewer's catalogue; the catalogue is merged at runtime).
5. Decide collaboration separately (6.1).
6. Needed before starting: npm scope access for the new package name and trusted-publisher entry; agreement on the element tag prefix (today `pptx-ui-*`; shared elements need a neutral prefix such as `office-ui-*`, with the old tags kept as thin aliases for one major version to preserve the public API); a visual/browser regression baseline (Playwright suites live in the viewers and must keep running against the shared elements).

Nothing in section 6 is implemented in this PR.

## 7. First slice delivered (this PR)

See `PROVENANCE.md` ("diagram area"): `src/diagram` and `src/docx/diagram*.ts`. Not moved yet, on purpose: layout engine, interpreters, fabrication, the typed layout-definition tree, 3D, and everything in steps 2-6 above.
