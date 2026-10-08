# Structure and parity plan (2026-10-08)

Where the repository stands after a full survey, and the order of work toward
one shared implementation per Office concept, shared UI bindings, and honest
Office 365 parity. Facts below come from reading the tree on 2026-10-08
(`origin/main` d9ef4e0); counts are approximate.

## 1. State

Health: `bun run typecheck` passes (strict and pptx projects). The core
vitest suite passes except two `modify-verifier*` tests that time out
(section 3, bug 1). No open GitHub issues or pull requests.

Size (non-test lines): `src/core/pptx` ~320k, `src/ui/src/pptx` ~195k,
`src/core/xlsx` ~49k, `src/core/visio` ~36k, `src/ui/src/docx` ~27k,
`src/core/docx` ~24k, `src/ui/src/xlsx` ~19k. The shared areas together are
under 30k, most of it geometry tables.

### 1.1 Logic that exists more than once

| Concept                                      | Copies today                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | Shared home                                                             |
| -------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------- |
| Theme (colour, font and format scheme)       | pptx `core/types/theme.ts` + 3 runtime files + switching utils (~2.6k); xlsx `write/theme.ts` (patches the part). docx, xlsx `read/theme.ts` and visio `theme.ts` read `a:clrScheme`/`a:fontScheme` through `drawingml` `parseTheme`; visio keeps only its `vt:` extensions and the `fmtScheme` style selection                                                                                                                                                                                  | `drawingml` (`theme*.ts`)                                               |
| Colour transforms                            | pptx `core/color/color-transforms.ts`; `drawingml/drawing-color.ts` + `ordered-color-transforms.ts` (docx, xlsx, chart and visio `drawingPaint` resolve through it); `math/omml-color.ts`                                                                                                                                                                                                                                                                                                        | `drawingml`, primitives in `color`                                      |
| Fill, line, effects                          | `drawingml/drawing-fill.ts`; pptx `drawing-fill-xml.ts`, `save-shape-fill-stroke.ts`, gradient and effect codecs; xlsx chart fill writers; visio fill, line and gradient modules (theme `fmtScheme` selectors keep their strict reject-on-malformed checks, so they do not use the lenient shared fill/line readers)                                                                                                                                                                             | `drawingml` (readers and writers, `a:ln` included; no effect model yet) |
| DrawingML text body (a:bodyPr, a:pPr, a:rPr) | pptx `text-run-properties-parser.ts` (557), `paragraph-properties-parser.ts` (371), `body-properties-parser.ts`; `diagram/drawing-text.ts` (57)                                                                                                                                                                                                                                                                                                                                                  | none                                                                    |
| Chart model (`c:chartSpace`)                 | pptx `core/types/chart.ts` (1574) + ~160 `chart-*` utils (object tree, own writers); xlsx `write/chart-patch.ts` and `chart-colors.ts` (patch kept parts); xlsx `read/chart.ts`, xlsx `write/chart.ts` (new parts) and docx `chart.ts` go through the shared model                                                                                                                                                                                                                               | `chart/model.ts`, `parseChartSpace`, `writeChartSpace`                  |
| Chart rendering                              | `chart/render/` (moved from xlsx `layout/chart-*.ts`) emits SVG strings for xlsx (`chartView`) and docx (`renderChartSpaceSvg`); pptx UI `render/chart-*` still emit descriptors                                                                                                                                                                                                                                                                                                                 | `chart/render` (pptx descriptors not yet on it)                         |
| Axis nice bounds                             | xlsx `layout/chart-scale.ts` `niceScale`; `chart/axis-nice.ts`                                                                                                                                                                                                                                                                                                                                                                                                                                   | `chart/axis-nice.ts`                                                    |
| Message digests and password spin            | pptx `core/utils/digests/` (1,619 lines: md2, md4, md5, ripemd128/160, whirlpool, dispatcher) + `modify-verifier-codec.ts` `iteratedHash` (awaits one Web Crypto digest per round); `digest/` (788 lines, synchronous SHA family, `password-hash.ts` spin)                                                                                                                                                                                                                                       | `digest`                                                                |
| Document properties                          | pptx `PptxDocumentPropertiesUpdater.ts` (466, fast-xml-parser) + `app-properties-*.ts`; `opc/properties` (used by xlsx and docx)                                                                                                                                                                                                                                                                                                                                                                 | `opc/properties`                                                        |
| Elbow connector geometry                     | `src/ui/src/pptx/render/connector-elbow-geometry.ts` (204) and `src/core/geometry/connector-elbow-geometry.ts` (160), same `elbowWaypoints` signature                                                                                                                                                                                                                                                                                                                                            | `geometry`                                                              |
| Locale normalisation and translator          | `src/ui/src/docx/localization.ts`, `src/core/xlsx/ui/localization.ts` (both `normalizeEditorLocale`), `src/core/xlsx/ui/locales/merge.ts`, `src/ui/src/pptx/i18n/translator.ts`; visio and teams have no i18n                                                                                                                                                                                                                                                                                    | none                                                                    |
| Web-control registration contract            | `src/ui/src/registry.ts` (`office-ui.web-control-contract`) and `src/ui/src/pptx/web-components/registration-contract.ts` (`pptx-viewer.web-control-contract`)                                                                                                                                                                                                                                                                                                                                   | `registry.ts`                                                           |
| Theme token bridge                           | docx `--dve-*` via `shadcnBridge`; xlsx `--xve-*` via `office-bridge.css`; pptx `--pptx-*` via `office-token-bridge.ts`; visio `--vv-*` with no bridge; only teams uses `--office-*` directly                                                                                                                                                                                                                                                                                                    | `theme-bridge.ts`, `tokens.ts`                                          |
| Chrome widgets                               | title bar and status bar: one shared element each (`office-ui-title-bar`, `office-ui-status-bar`), which pptx subclasses and themes through `:host`-only token bridges; comments pane: `office-ui-comments-pane` draws Word's pane and Excel's Show Comments list, the five pptx bindings still draw their own panels; backstage, ruler (shared for Word since Wave 2), find and replace, zoom, colour grid, keytips wrappers, presence: a divergent copy per product next to a shared primitive | `src/ui/src/{chrome,comments,form,menu,presence}`                       |

### 1.2 Bindings

- docx, xlsx and visio each have a `bindings` package with one imperative
  mount factory and thin framework glue. Teams has none: every framework
  repeats the same ~40-line lifecycle. pptx hand-rolls each framework root
  (React 1.7k, Vue 2.3k, Angular 3.9k, vanilla 2.1k lines).
- Drift found: pptx `showToolbar`, `initialSlide`, `showThumbnails` exist in
  svelte and vanilla only; docx Svelte exposes `isDirty()` and no `element`
  where the others expose `dirty` and `element`; xlsx Svelte exposes
  `getElement()`; visio accepts `aria-label` in React only; teams accepts
  `className` in React only; the teams `open-file` payload shape differs per
  framework; `pptx-viewer-shared` pins `ooxml-ui ^0.30.1` while the bindings
  pin `^0.31.0`; visio framework packages import `../../bindings/src/...`
  relatively instead of declaring the dependency.
- The pptx changelog shows the same bug fixed five times, once per binding
  (SmartArt fill opacity, ribbon popups, locale re-translation): the cost of
  having no shared binding layer.

### 1.3 Documentation and configuration drift

Resolved in Waves 1 and 2 (2026-10-08):

- `AGENTS.md`: the PowerPoint row of the "Where does my change go?" table, the
  package names and the status line were corrected (Wave 1 item 8).
- `viewers/xlsx/docs/features.md` now names the labelled-frame chart families,
  states that chart-style authoring is unsupported and takes its function count
  (481) from the registry (Wave 1 item 8).
- `.github/workflows/ci.yml`: the `lint` job no longer waits on `typecheck.ui`
  (Wave 1 item 8).
- `src/ui/tsconfig.pptx.json` extends the base (Wave 2).
- The Playwright Chromium helper is shared by docx and xlsx (Wave 2).
  `tsconfig.release.json` is still duplicated between them.
- Every viewer typecheck is green (Wave 2).
- Four of the five `.oxfmtrc.json` copies were dropped on 2026-10-08 (docx,
  teams, visio and xlsx only restated the root settings).

Still open:

- `.oxfmtrc.json`: the pptx viewer keeps its own copy (it spells out every
  option and adds `ignorePatterns` for `dist/` and minified files); fold those
  into the root config before removing it.
- `src/ui` `sideEffects` lists only the pptx entry.
- The `pptx/*` and `pptx/editor/*` wildcard exports resolve to fixed bundler
  lists that differ between tsup and tsdown.
- 306 of 4,378 non-test modules under `src/` exceed 300 lines (recounted on
  2026-10-08; tests, fixtures and declaration files excluded).

## 2. Target structure

The direction in `agnostic-core-plan.md` stands. The additions below come from
the survey.

1. **`drawingml` area** (new): neutral `DrawingColor` and resolver, fill,
   line, effect and text-body models with DOM parsers and writers, and a
   neutral theme model (colour scheme, font scheme, format scheme, `phClr`).
   The provisional files in `diagram/` move here unchanged. docx, xlsx and
   visio consume it first (they are already on the shared DOM); pptx follows
   through adapters as it leaves `fast-xml-parser`.
2. **`chart` grows a model**: `c:chartSpace` and `cx:` types, DOM parsers and
   writers, a renderer-neutral view model. The xlsx reader and writer become
   the first client, pptx `types/chart.ts` aliases it, and docx gains chart
   loading. One SVG painter for xlsx and the pptx DOM renderer follows.
3. **`digest` owns every hash** Office uses on protection and verifier
   elements, synchronously. `crypto` keeps encryption key derivation.
4. **`i18n` area** (new, DOM-free): locale normalisation, dotted-key
   dictionary merge, translator factory with English fallback. Every product
   UI uses it; visio and teams get dictionaries.
5. **`opc/properties` for every format**: the pptx updater becomes an adapter.
6. **Shared UI**: product UIs build on the `chrome/`, `form/`, `menu/` and
   `presence/` primitives and `--office-*` tokens; the per-product bridges
   shrink to a token alias sheet; one registration contract.
7. **Bindings**: one `bindings` package per product with a mount factory
   (teams gets one; pptx gets a factory for options, events and handle, with
   the view code behind it), one handle vocabulary across products
   (`element`, `dirty`, `load`, `save`), and a per-binding parity test that
   diffs the prop, event and handle lists against the shared contract.
8. **Honest parity ledgers**: each product keeps one capability ledger
   (`viewers/<p>/docs/parity.md` or `features.md`) that the parity reviews
   update, with native-reference evidence per row. No parity percentage.

## 3. Outstanding bugs

1. **Modify-password verification takes over 30 s**
   (`src/core/pptx/core/utils/modify-verifier-codec.ts`): `iteratedHash`
   awaits `crypto.subtle.digest` 100,000 times. The shared
   `digest/password-hash.ts` spins synchronously. Two tests fail from this.
2. **Autosave recovery snapshot missing in a binding**:
   `e2e/pptx/autosave-recovery-prompt.spec.ts` (lines 226, 271) and
   `autosave-recovery-encryption.spec.ts` (203) skip when a binding "never
   wrote a recovery snapshot (no dirty flag)". Root cause to find and fix in
   the shared layer, with per-binding tests.
3. **Known gradient raster gaps** pinned as expected failures:
   `e2e/xlsx/chart-gradient-raster.spec.ts`, `chart-series-path-raster.spec.ts`,
   `src/core/chart/gradient-raster.test.ts` (three-transparent-stop and
   coincident-angle cases).
4. **Continuous section balancing** after multi-column sections starts a new
   page (`src/core/docx/layout/page-flow.ts:83`); floats and footnotes are not
   balanced.
5. **Drift listed in 1.2 and 1.3** (each item is a bug for one binding or one
   reader of the docs).
6. `digest/algorithm-names.ts` recognises MD5, RIPEMD and WHIRLPOOL but cannot
   compute them, while pptx has pure-TypeScript implementations. Excel and
   Word protection with those digests therefore report "unsupported".

## 4. Waves

Each item is one small conventional commit series on `main`, strict
TypeScript, tests next to the code, provenance recorded for moved modules,
and no claim of parity without evidence.

### Wave 1 (this session)

| #   | Work                                                                                                                                                                                    | Scope       |
| --- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------- |
| 1   | Move pptx pure-TS digests into `digest`, make `modify-verifier` use `digest/password-hash`, delete the duplicate dispatcher (forwarders stay). Fixes bugs 1 and 6.                      | core        |
| 2   | Create `drawingml`: move the provisional diagram files, add the neutral theme model and DOM parser, port docx `theme.ts` and xlsx `read/theme.ts` onto it. Update plan docs.            | core        |
| 3   | Create `i18n` (`ooxml-core/i18n`): locale normalisation, dictionary merge, translator. docx, xlsx and pptx UIs consume it; remove the three copies.                                     | core, ui    |
| 4   | Remove the duplicate elbow geometry, the duplicate `niceScale` (if outputs agree), and the second web-control contract symbol.                                                          | core, ui    |
| 5   | Binding drift: visio `aria-label` everywhere; docx and xlsx Svelte handles aligned; teams `class` everywhere and a `teams-bindings` mount factory; dependency ranges; relative imports. | viewers     |
| 6   | pptx binding drift: `showToolbar`, `initialSlide`, `showThumbnails` in React, Vue and Angular, with tests.                                                                              | viewers     |
| 7   | Autosave recovery snapshot (bug 2): root cause and fix across bindings.                                                                                                                 | ui, viewers |
| 8   | Docs and CI hygiene: `AGENTS.md` table, names and status; `features.md` chart section and function count from the registry; `ci.yml` lint gate; `.github/README.md` names.              | docs, ci    |

### Wave 2 (landed 2026-10-08, see the commit log for details)

- `chart` model and parsers (plan step 5): the neutral `ChartSpace` model,
  `parseChartSpace` over the shared DOM, xlsx as first client (identical output
  on 364 real chart parts) and Word chart loading are in. Remaining: a neutral
  writer, the pptx alias, chartex (`agnostic-core-plan.md`, step 5).
- pptx document properties through `opc/properties`: done. The shared writers
  now patch parts in place (unchanged bytes survive for every format); the
  pptx updater is a thin adapter. Output differs from the old pptx writer only
  in formatting and is closer to the source file.
- Shared chrome: the Word ruler drives `office-ui-ruler` (markers, margins,
  zoom and extent were added to the shared element); Visio follows the shared
  `--office-*` theme through one alias block; the xlsx bridge is the shared
  `shadcnBridge`. Title bar and status bar: done. docx, xlsx and pptx draw
  `office-ui-title-bar`, and all four products `office-ui-status-bar` (Visio
  has no title-bar row; its Quick Access Toolbar sits in the ribbon head).
  `pptx-ui-title-bar` and `pptx-ui-status-bar` are subclasses that keep their
  published events and hooks; their look is `:host` tokens only (the shared
  elements gained `--office-title-bar-name-foreground`,
  `--office-title-bar-result-selected` and
  `--office-status-bar-separator-opacity` for it), and the pptx style modules
  are gone. Backstage and find bar remain.
- Diagram layout engine out of pptx (plan step 2): model types, the engine,
  the `dgm:choose` walkers (ported to the engine's ordered-XML tree, pptx
  converts its raw slots at one adapter), the constraint solver and the
  interpreter model moved to `diagram`, strict; the hierarchy family and the
  remaining interpreters stay in pptx (step 2 status in
  `agnostic-core-plan.md`).
- Visio reads its theme and DrawingML colours through `drawingml`; the
  `Diagram*` aliases are gone outside `diagram`.
- Gradient raster gaps (bug 3): the coincident-stop case is fixed from native
  measurements (Excel paints a one-step ramp, not a hard edge); the sixteen
  translucent three-stop cases stay pinned with measured differences in the
  test and in `xlsx-parity-review.md`.
- `src/ui/tsconfig.pptx.json` extends the base; `src/core/pptx/ui` compiles
  strict; the Playwright Chromium helper is shared by docx and xlsx.
- Every viewer typecheck is green (the `e2e/docx`, `e2e/xlsx`, teams Svelte
  demo and pptx CLI breaks were fixed).

### Wave 3 (landed 2026-10-08)

- SmartArt (plan step 2): the `dgm:choose` walkers now read the engine's
  ordered-XML tree (pptx converts its raw slots once at one adapter), and the
  constraint solver and interpreter model moved to `diagram`, strict, with
  byte-identical results on 353 layout definitions. Remaining: the hierarchy
  family (51 modules), the remaining interpreters, editing, and a neutral
  paragraph model for text projection.
- Neutral chart writer: `writeChartSpace` round-trips Office-written parts
  byte for byte where the model is complete; new Excel charts go through it.
  Remaining: trendlines, error bars, walls and floors, data tables, display
  units in the model; `chart-patch.ts`; the pptx writers.
- Chart parity harness: every chart part in every committed deck (81 parts)
  is parsed by both the pptx and the neutral parser and compared; all 74
  classic parts agree. It found and fixed three bugs: the neutral parser
  missed the line group's `c:smooth`; pptx ignored `c:style` written under
  `mc:Fallback`; pptx treated a bare data-label switch as off. The pptx chart
  style palette now follows Office's gallery (style 2 starts at accent 1).
- Shared chrome: the pptx title and status bars are themed through host tokens
  only (their style modules are gone). `office-ui-comments-pane` carries Word's
  pane and Excel's Show Comments list. Remaining: backstage, find bar, the pptx
  comments panels (one per binding), Excel threaded-comment editing.
- Collaboration: `ooxml-core/xlsx/collab` binds an Excel edit session to a Yjs
  workbook (cells, formulas, styles, rows, columns, merges, sheets, names,
  selections; conflict rules documented in `collab-area.md`). Remaining: UI
  wiring in `<xlsx-editor>`, and the docx adapter moving into its area.
- Excel sparklines are read, laid out and drawn (line, column, win/loss) from a
  fixture generated by real Excel. Remaining: authoring, exact pixels.
- Hygiene: the slow Word Yjs test exposed a real cost (ribbon control lookups
  re-scanned the DOM on every transaction) that is now cached; Visio theming is
  documented; four redundant `.oxfmtrc.json` copies are gone.

### Wave 4 (next)

- pptx off `fast-xml-parser` area by area (plan step 6), starting with charts
  through the parity harness and SmartArt through the ordered-XML tree. Charts
  started: classic chart parts are read by the neutral parser and adapted to
  `PptxChartData` (proven deeply equal to the old parser on every committed
  chart); styles, ChartEx and the chart writers still use the object tree.
  SmartArt layout (agnostic-core plan step 2): the hierarchy arranger and its
  closure moved to `diagram/hierarchy` and `diagram/layout`; the other
  interpreters, editing reflow and node text projection are still in pptx.
- One chart painter for xlsx and the pptx DOM renderer.
- pptx binding factory; per-binding contract parity tests for all products;
  the pptx comments panels on the shared pane.
- xlsx collaboration UI wiring; docx `DocumentAdapter` into `src/core/docx`.
- Shared backstage and find bar.
- Product parity: Word editable Print Layout, charts and shapes drawn; Excel
  pivot rendering and Page Layout; Visio connector routing and glue.
