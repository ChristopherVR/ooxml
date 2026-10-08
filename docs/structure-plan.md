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

| Concept                                      | Copies today                                                                                                                                                                                                                                                                                                         | Shared home                                            |
| -------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------ |
| Theme (colour, font and format scheme)       | pptx `core/types/theme.ts` + 3 runtime files + switching utils (~2.6k); xlsx `write/theme.ts` (patches the part). docx, xlsx `read/theme.ts` and visio `theme.ts` read `a:clrScheme`/`a:fontScheme` through `drawingml` `parseTheme`; visio keeps only its `vt:` extensions and the `fmtScheme` style selection      | `drawingml` (`theme*.ts`)                              |
| Colour transforms                            | pptx `core/color/color-transforms.ts`; `drawingml/drawing-color.ts` + `ordered-color-transforms.ts` (docx, xlsx, chart and visio `drawingPaint` resolve through it); `math/omml-color.ts`                                                                                                                            | `drawingml`, primitives in `color`                     |
| Fill, line, effects                          | `drawingml/drawing-fill.ts`; pptx `drawing-fill-xml.ts`, `save-shape-fill-stroke.ts`, gradient and effect codecs; xlsx chart fill writers; visio fill, line and gradient modules (theme `fmtScheme` selectors keep their strict reject-on-malformed checks, so they do not use the lenient shared fill/line readers) | `drawingml` (readers and writers; no effect model yet) |
| DrawingML text body (a:bodyPr, a:pPr, a:rPr) | pptx `text-run-properties-parser.ts` (557), `paragraph-properties-parser.ts` (371), `body-properties-parser.ts`; `diagram/drawing-text.ts` (57)                                                                                                                                                                      | none                                                   |
| Chart model (`c:chartSpace`)                 | pptx `core/types/chart.ts` (1574) + ~160 `chart-*` utils (object tree); xlsx `write/chart*.ts`, `layout/chart-*.ts`; xlsx `read/chart.ts` and docx `chart.ts` read the shared model                                                                                                                                  | `chart/model.ts`, `parseChartSpace`                    |
| Chart rendering                              | xlsx `layout/chart-svg*.ts` emit SVG strings; pptx UI `render/chart-*` emit descriptors; docx shows a placeholder                                                                                                                                                                                                    | none                                                   |
| Axis nice bounds                             | xlsx `layout/chart-scale.ts` `niceScale`; `chart/axis-nice.ts`                                                                                                                                                                                                                                                       | `chart/axis-nice.ts`                                   |
| Message digests and password spin            | pptx `core/utils/digests/` (1,619 lines: md2, md4, md5, ripemd128/160, whirlpool, dispatcher) + `modify-verifier-codec.ts` `iteratedHash` (awaits one Web Crypto digest per round); `digest/` (788 lines, synchronous SHA family, `password-hash.ts` spin)                                                           | `digest`                                               |
| Document properties                          | pptx `PptxDocumentPropertiesUpdater.ts` (466, fast-xml-parser) + `app-properties-*.ts`; `opc/properties` (used by xlsx and docx)                                                                                                                                                                                     | `opc/properties`                                       |
| Elbow connector geometry                     | `src/ui/src/pptx/render/connector-elbow-geometry.ts` (204) and `src/core/geometry/connector-elbow-geometry.ts` (160), same `elbowWaypoints` signature                                                                                                                                                                | `geometry`                                             |
| Locale normalisation and translator          | `src/ui/src/docx/localization.ts`, `src/core/xlsx/ui/localization.ts` (both `normalizeEditorLocale`), `src/core/xlsx/ui/locales/merge.ts`, `src/ui/src/pptx/i18n/translator.ts`; visio and teams have no i18n                                                                                                        | none                                                   |
| Web-control registration contract            | `src/ui/src/registry.ts` (`office-ui.web-control-contract`) and `src/ui/src/pptx/web-components/registration-contract.ts` (`pptx-viewer.web-control-contract`)                                                                                                                                                       | `registry.ts`                                          |
| Theme token bridge                           | docx `--dve-*` via `shadcnBridge`; xlsx `--xve-*` via `office-bridge.css`; pptx `--pptx-*` via `office-token-bridge.ts`; visio `--vv-*` with no bridge; only teams uses `--office-*` directly                                                                                                                        | `theme-bridge.ts`, `tokens.ts`                         |
| Chrome widgets                               | title bar, status bar, backstage, ruler, find and replace, zoom, colour grid, keytips wrappers, presence, comments pane: a divergent copy per product next to a shared primitive                                                                                                                                     | `src/ui/src/{chrome,form,menu,presence}`               |

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

- `AGENTS.md`: the "Where does my change go?" table sends PowerPoint drawing
  to `viewers/pptx/packages/shared`, which is a facade; package names are
  wrong (docx bindings are unscoped `docx-*-viewer`, teams are
  `openteams-*-viewer`); the status line omits `chart`, `math`, `text`,
  `visio` and `teams` and describes `drawingml` as planned.
- `viewers/xlsx/docs/features.md` did not name the chart families that draw
  as a labelled frame (bubble, stock, surface) or that chart-style authoring
  is unsupported, and its function count was approximate (the registry holds
  481; the survey's own grep undercounted it). Fixed in Wave 1 item 8.
- `.github/workflows/ci.yml`: the `lint` job is gated on `typecheck.ui`.
- `src/ui/tsconfig.pptx.json` does not extend the base; `src/ui`
  `sideEffects` lists only the pptx entry; the `pptx/*` and `pptx/editor/*`
  wildcard exports resolve to fixed bundler lists that differ between tsup and
  tsdown.
- Five `.oxfmtrc.json` copies; docx and xlsx duplicate `playwright-chromium.ts`
  and `tsconfig.release.json`.
- 309 of 4,310 non-test modules exceed 300 lines.

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

### Wave 2

- `chart` model and parsers (plan step 5), xlsx reader first, docx `c:chart`.
  Started: the model, `parseChartSpace`, the xlsx reader and docx chart
  loading are in; the pptx alias, chartex and a neutral writer remain
  (`agnostic-core-plan.md`, step 5).
- pptx document properties through `opc/properties`.
- Shared chrome: title bar, status bar, backstage and find bar used by all
  four products; token bridges reduced to alias sheets.
- Diagram layout engine out of pptx (plan step 2): model types and engine moved to
  `diagram`; constraint solver, hierarchy and interpreters remain (see the step 2
  status in `agnostic-core-plan.md`).
- Gradient raster gaps (bug 3) with native reference images.
- `src/ui/tsconfig.pptx.json` extends the base; tighten one pptx directory at
  a time; shared `playwright-chromium` and `tsconfig.release` for docx and
  xlsx.

### Wave 3

- pptx off `fast-xml-parser` area by area (plan step 6).
- One chart painter for xlsx and the pptx DOM renderer.
- pptx binding factory; per-binding contract parity tests for all products.
- Collaboration: xlsx and docx `DocumentAdapter` implementations in their
  areas.
- Product parity: Word editable Print Layout, charts and shapes; Excel pivot
  rendering, sparklines, Page Layout; Visio connector routing and glue.
