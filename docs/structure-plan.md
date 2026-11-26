# Structure and parity plan (2026-10-08)

Where the repository stands after a full survey, and the order of work toward
one shared implementation per Office concept, shared UI bindings, and honest
Office 365 parity. Facts below come from reading the tree on 2026-10-08
(`origin/main` d9ef4e0); counts are approximate and sections 1 to 3 are that
snapshot.

**Status on 2026-10-22.** Waves 1 to 4 landed and the viewers now live in this repository, so the
separate-repository observations (version pins, repository names) in section 1.2 are historical.
Since the snapshot: the duplicate digests, the slow modify-verifier and the missing MD5, RIPEMD and
WHIRLPOOL digests are fixed (`digest` computes them; bugs 1 and 6); the elbow connector geometry
duplicate is gone (the pptx renderer imports `ooxml-core/geometry`); `drawingml/write-theme.ts`
exists; the shared backstage, find bar and print preview live in `src/ui/src/chrome`; continuous
section column balancing exists (`docx/layout/balance-columns.ts`, bug 4); editable Word Print
Layout, Word chart drawing, Excel pivots, Page Layout and collaboration, and Visio collaboration
landed (see the parity pages). `niceScale` is still in `chart/render/chart-scale.ts` next to
`chart/axis-nice.ts`, pinned equal by `xlsx/layout/chart-scale-vs-axis-nice.test.ts`. Wave 5 below is
the open list.

## Current shared Office status (2026-10-27)

The [delivery ledger](office-shared-parity-delivery.md) is the active checklist. Shared DrawingML
outlines/effect lists/format schemes and pattern/picture codecs are implemented. Word and Excel
use the neutral SmartArt interpreter, richer painter, shared style/colour galleries and conservatively
admitted generated drawing caches. They also edit supported ChartEx data, family options and
shared advanced DrawingML formatting.
PowerPoint, Word, Excel and embedded Visio charts mount the shared advanced Format pane.

PowerPoint classic saves now have neutral label/paint/text/axis/structure/analysis/point families,
with passing corpus, browser and native edit/reopen checks. ChartEx has shared factory, save,
style/axis/label operations and compatibility projections. The remaining architecture work is
adaptation of unsupported source details and removal of covered
legacy paths, rather than reimplementing these shared codecs or controls. Picture media allocation,
unsafe combo splits, painting the approximated effect DAG members, broader native typography/3D fidelity and the complete
Office style catalogue remain explicit limits. Sections and wave notes below are dated snapshots.

## 1. State

Health at the snapshot: `bun run typecheck` passed (strict and pptx projects) and the core vitest
suite passed except two `modify-verifier*` tests that timed out (section 3, bug 1, since fixed).

Size (non-test lines): `src/core/pptx` ~320k, `src/ui/src/pptx` ~195k,
`src/core/xlsx` ~49k, `src/core/visio` ~36k, `src/ui/src/docx` ~27k,
`src/core/docx` ~24k, `src/ui/src/xlsx` ~19k. The shared areas together are
under 30k, most of it geometry tables.

### 1.1 Logic that exists more than once

| Concept                                      | Copies today                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | Shared home                                                             |
| -------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------- |
| Theme (colour, font and format scheme)       | pptx `core/types/theme.ts` + 3 runtime files + switching utils (~2.6k); xlsx `write/theme.ts` (patches the part; `drawingml/write-theme.ts` is the neutral writer pptx and visio use). docx, xlsx `read/theme.ts` and visio `theme.ts` read `a:clrScheme`/`a:fontScheme` through `drawingml` `parseTheme`; visio keeps only its `vt:` extensions and the `fmtScheme` style selection                                                                                                             | `drawingml` (`theme*.ts`)                                               |
| Colour transforms                            | pptx `core/color/color-transforms.ts`; `drawingml/drawing-color.ts` + `ordered-color-transforms.ts` (docx, xlsx, chart and visio `drawingPaint` resolve through it); `math/omml-color.ts`                                                                                                                                                                                                                                                                                                        | `drawingml`, primitives in `color`                                      |
| Fill, line, effects                          | `drawingml/drawing-fill.ts`; pptx `drawing-fill-xml.ts`, `save-shape-fill-stroke.ts`, gradient and effect codecs; xlsx chart fill writers; visio fill, line and gradient modules (theme `fmtScheme` selectors keep their strict reject-on-malformed checks, so they do not use the lenient shared fill/line readers)                                                                                                                                                                             | `drawingml` (readers and writers, `a:ln` included; no effect model yet) |
| DrawingML text body (a:bodyPr, a:pPr, a:rPr) | pptx `text-run-properties-parser.ts` (557), `paragraph-properties-parser.ts` (371), `body-properties-parser.ts` (object tree, pinned equal to the neutral reader by `smartart-text-neutral-parity.test.ts`); `drawingml/text-model.ts` with `read-*`/`write-*` (the full model, used by `diagram/text`); `drawingml/drawing-text.ts` (the flat summary charts use)                                                                                                                               | `drawingml` (`text-model.ts`)                                           |
| Chart model (`c:chartSpace`)                 | pptx `core/types/chart.ts` (1574) + ~160 `chart-*` utils (object tree, own writers); xlsx `write/chart-patch.ts` and `chart-colors.ts` (patch kept parts); xlsx `read/chart.ts`, xlsx `write/chart.ts` (new parts) and docx `chart.ts` go through the shared model                                                                                                                                                                                                                               | `chart/model.ts`, `parseChartSpace`, `writeChartSpace`                  |
| Chart rendering                              | `chart/render/` (moved from xlsx `layout/chart-*.ts`) emits SVG strings for xlsx (`chartView`) and docx (`renderChartSpaceSvg`); pptx UI `render/chart-*` still emit descriptors                                                                                                                                                                                                                                                                                                                 | `chart/render` (pptx descriptors not yet on it)                         |
| Axis nice bounds                             | xlsx `layout/chart-scale.ts` `niceScale`; `chart/axis-nice.ts`                                                                                                                                                                                                                                                                                                                                                                                                                                   | `chart/axis-nice.ts`                                                    |
| Message digests and password spin            | Resolved: the pptx copies moved into `digest/` (md2, md4, md5, ripemd128/160, whirlpool, `password-hash.ts`); `modify-verifier-codec.ts` calls the synchronous spin                                                                                                                                                                                                                                                                                                                              | `digest`                                                                |
| Document properties                          | pptx `PptxDocumentPropertiesUpdater.ts` (466, fast-xml-parser) + `app-properties-*.ts`; `opc/properties` (used by xlsx and docx)                                                                                                                                                                                                                                                                                                                                                                 | `opc/properties`                                                        |
| Elbow connector geometry                     | Resolved: `src/ui/src/pptx/render/connector-elbow-geometry.ts` now imports `src/core/geometry/connector-elbow-geometry.ts` (`elbowWaypoints`)                                                                                                                                                                                                                                                                                                                                                    | `geometry`                                                              |
| Locale normalisation and translator          | `src/ui/src/docx/localization.ts`, `src/core/xlsx/ui/localization.ts` (both `normalizeEditorLocale`), `src/core/xlsx/ui/locales/merge.ts`, `src/ui/src/pptx/i18n/translator.ts`; visio and teams have no i18n                                                                                                                                                                                                                                                                                    | `i18n` (`ooxml-core/i18n`)                                              |
| Web-control registration contract            | `src/ui/src/registry.ts` (`office-ui.web-control-contract`) and `src/ui/src/pptx/web-components/registration-contract.ts` (`pptx-viewer.web-control-contract`)                                                                                                                                                                                                                                                                                                                                   | `registry.ts`                                                           |
| Theme token bridge                           | docx `--dve-*` via `shadcnBridge`; xlsx `--xve-*` via `office-bridge.css`; pptx `--pptx-*` via `office-token-bridge.ts`; visio `--vv-*` with no bridge; only teams uses `--office-*` directly                                                                                                                                                                                                                                                                                                    | `theme-bridge.ts`, `tokens.ts`                                          |
| Chrome widgets                               | title bar and status bar: one shared element each (`office-ui-title-bar`, `office-ui-status-bar`), which pptx subclasses and themes through `:host`-only token bridges; comments pane: `office-ui-comments-pane` draws Word's pane and Excel's Comments side pane, the five pptx bindings still draw their own panels; backstage, ruler (shared for Word since Wave 2), find and replace, zoom, colour grid, keytips wrappers, presence: a divergent copy per product next to a shared primitive | `src/ui/src/{chrome,comments,form,menu,presence}`                       |

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

1. **Resolved.** Modify-password verification took over 30 s
   (`src/core/pptx/core/utils/modify-verifier-codec.ts`): `iteratedHash`
   awaited `crypto.subtle.digest` 100,000 times. It now uses the synchronous
   `digest/password-hash.ts` spin.
2. **Autosave recovery snapshot missing in a binding**:
   `e2e/pptx/autosave-recovery-prompt.spec.ts` (lines 226, 271) and
   `autosave-recovery-encryption.spec.ts` (203) skip when a binding "never
   wrote a recovery snapshot (no dirty flag)". Root cause to find and fix in
   the shared layer, with per-binding tests.
3. **Known gradient raster gaps** pinned as expected failures:
   `e2e/xlsx/chart-gradient-raster.spec.ts`, `chart-series-path-raster.spec.ts`,
   `src/core/chart/gradient-raster.test.ts` (three-transparent-stop and
   coincident-angle cases).
4. **Mostly resolved.** Continuous section balancing after multi-column
   sections started a new page; `src/core/docx/layout/page-flow.ts` now balances
   columns through `balance-columns.ts`. Floats and footnotes in balanced
   sections were not re-verified.
5. **Drift listed in 1.2 and 1.3** (each item is a bug for one binding or one
   reader of the docs).
6. **Resolved.** `digest` now computes MD2, MD4, MD5, RIPEMD-128/160 and WHIRLPOOL
   (the pptx copy was moved into it), so Excel and Word protection with those
   digests verifies.

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
  test and in `xlsx-parity-status.md`.
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
  pane and Excel's Comments side pane. Remaining: backstage, find bar, the pptx
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

### Wave 4 (landed 2026-10-08)

Browser verification ran for the first time since Wave 1, one agent per
product, with `--workers=2` on Chromium (Playwright 1.63):

- PowerPoint: 148 tests across the autosave-recovery, web-control-contract,
  title-bar, status-bar, ui-primitives, every chart and every SmartArt spec
  pass on all five demos (the WebGL-gated 3D specs ran on SwiftShader); new
  specs cover the root display options and the chart style palette. It caught
  two real cross-binding bugs: Vue, Angular and Svelte never bound the chart
  axis-label fonts the shared view model emits, and chart style palettes were
  built from the fixed Office accents rather than the deck theme (Angular
  ignored the style entirely). Both are fixed in the shared layer with a test
  per binding; the chart now carries the theme accents for rendering only.
- Excel: every `chart-*` spec (387 tests) plus `framework-matrix`,
  `clipboard-parity`, `formatting`, `editing` pass in all six frameworks; new
  specs cover sparklines, the Show Comments list, every binding's own handle
  and two-window collaboration.
- Word: 452 tests across the comment, review, field, locale, parity and
  ribbon specs pass in all six frameworks; new specs cover the shared comments
  pane, the ruler drag and the binding handles. It caught a real Wave 2
  regression: Word never registered `office-ui-ruler` (the unit test had
  masked it by registering everything), now fixed with a test that would have
  failed. Two flaky specs were made deterministic.
- Visio: 68 ungated tests pass (the `VISIO_NATIVE_*` captures stay skipped);
  new specs cover the shared theme tokens and the `aria-label` on every host.
  It caught a real flaky-share bug (the host never published a drawing when a
  load-time zoom swallowed the load event), now fixed with a regression test.
- Teams: the whole suite (30 tests) passes in all six frameworks; new specs
  cover the host class, the `open-file` event and the PowerPoint reading view.
- The demos for Excel and Word now drive each binding's real handle (before,
  no browser test reached the bindings' handles at all).

Implementation:

- pptx charts off `fast-xml-parser` (plan step 6): classic chart parts are read
  by the neutral parser and adapted to `PptxChartData`, proven deeply equal to
  the old parser on every committed chart before the old parser was deleted;
  the neutral model gained walls, floors, data tables and display units. It
  exposed two pptx writers that saved undeclared `c15`/`c16` prefixes (fixed),
  and decks this library saved that way are repaired on load with a reported
  issue. Styles, ChartEx and the chart writers still use the object tree.
- SmartArt layout (plan step 2): the hierarchy arranger, every other
  interpreter up to `interpretSmartArtLayout` and their closure (253 files)
  moved to `diagram/hierarchy` and `diagram/layout`, strict, byte-identical on
  the corpus; node text projection (needs a neutral paragraph model) and the
  editing reflow (needs a neutral drawing shape) remain in pptx.
- One chart painter: the Excel SVG painter moved to `chart/render` over a
  neutral `ChartSummary`; `renderChartSpaceSvg` draws any `ChartSpace`. Word
  now draws bar, column, line, area, pie, doughnut, scatter and radar charts
  in the editor and in Print Layout (read-only; bubble, stock, surface and
  chartex show a labelled frame). The pptx DOM renderer still has its own.
- Excel collaboration is wired: `collaboration` on `<xlsx-editor>` and every
  binding, File > Share, remote selections painted from a DOM-free layout,
  presence in the title bar, undo routed through the Yjs undo manager while
  bound. The browser spec found a join-time race (a guest could seed a blank
  workbook over the host's); the fix makes the shared transport announce
  presence before answering a sync request and makes a joiner wait for an
  earlier writer's workbook.

### Wave 4 follow-ups (landed 2026-10-08)

- Excel opened a real 66,000-formula workbook (one 6.5 MB sheet) in about
  1.6 s of core time because every worksheet was parsed into a full DOM, and
  the first edit recalculated every formula. Worksheet cell data now goes
  through a lightweight reader in `xml/lite.ts` (falling back to the DOM for
  anything it does not handle), the first edit trusts stored values and
  computes only what it reaches, and the formula tokenizer no longer slices
  the source per match: core load 1.6 s to 0.33 s, time to first painted
  grid in the browser roughly halved. `src/core/scripts/xlsx/profile-load.ts`
  measures each stage. Still slow: the first edit (about 0.5 s, building the
  dependency graph) and structural edits (full recalculation).
- Excel chrome now matches Excel 365: the editing-mode selector, Comments and
  Share sit at the right end of the ribbon tab row (a new `actions` slot on
  the shared ribbon tabs), the search box is centred in the title bar, the
  title bar's collaboration slot holds presence only, and the status bar
  reads Ready, selection statistics (a new `summary` slot on the shared
  status bar), view buttons, zoom. A reported double "Ready" bar could not be
  reproduced in any framework, width or theme, or on the deployed demo; a
  unit test and a browser spec now assert one status bar. Word and
  PowerPoint still place Comments and Share elsewhere (follow-up).

### Excel engine follow-ups (landed 2026-10-08)

- The formula dependency graph is now prepared in idle time after the first
  paint (`EditSession.prepareCalculation`, `src/ui/src/xlsx/calc-preparation.ts`,
  slices under 8 ms, paused while the user types or scrolls), and parsing is
  cached by formula shape (digits blanked; 129 shapes cover this workbook's
  66,434 formulas): the first edit after preparation costs about 12 ms instead
  of about 400 ms, first paint unchanged.
- Row and column inserts and deletes, and sheet renames, shift the graph in
  place and recalculate only what the edit can change, each checked against a
  full recalculation from scratch: about 3 s to about 0.5 s on that workbook,
  rename to 0.13 s. F9 now recalculates only dirty and volatile cells as Excel
  does (`calculateNow({ full: true })` is Ctrl+Alt+F9). Still rebuilding:
  undo of a row or column edit, edits on sheets with spills, cut array
  formulas, cell shifts, sort, move, sheet add/delete/move.
- Selection statistics use the first numeric cell's number format.
- The formula number lexeme is unambiguous (a CodeQL polynomial-regex alert
  raised by the tokenizer change); the three lint errors another commit left
  in the office app shell are fixed.

### Chrome consistency (landed 2026-10-08)

- One shared `office-ui-ribbon-actions` element (mode selector, Comments,
  accent Share, in Office's order) sits in the ribbon tab row's `actions`
  slot for Excel and Word; PowerPoint's five bindings place a
  `pptx-ui-ribbon-actions` subclass at the right end of their own tab row
  (Record, Comments, Share). Word gained a Share entry point (a cancelable
  `file-command`, the host supplies the transport). The pptx tab row itself
  moved onto `office-ui-ribbon` on 2026-11-16 (section 5).

### Wave 5 (open items, status 2026-10-22)

Landed since the plan was written: the shared backstage, find bar and print preview
(`src/ui/src/chrome`), editable Word Print Layout, Excel pivot authoring and Page Layout, Excel and
Visio collaboration. Still open at that historical checkpoint (see the current ledger above):

- pptx binding factory; per-binding contract parity tests for all products;
  the pptx comments panels on the shared pane; the pptx tab row on
  `office-ui-ribbon`.
- pptx chart styles, ChartEx and the chart writers onto the neutral model;
  the pptx DOM chart renderer onto `chart/render`.
- SmartArt text editing out of pptx (the neutral paragraph model in `drawingml` and the shared
  node text projection in `diagram/text` landed 2026-10-23; Word and Excel read node paragraphs
  read-only); the engine's ordered XML onto the shared `xml` DOM.
- The docx Yjs mapping lives in `src/core/docx/ui/yjs-collaboration.ts`; whether it should become
  a `DocumentAdapter` over `collab` like the Excel mapping is undecided.
- Product parity is tracked in the per-product pages, not here: Word shapes and page-story
  editing, Excel pivot filters and sheet-part collaboration, Visio connector routing and glue.

## 5. Framework layers and shared chart and SmartArt editing (2026-10-23)

A survey on 2026-10-23 (`origin/main` 3b2ecd9a6) found that the largest remaining duplication is not
Office logic but the framework layer. At that survey's start, chart and SmartArt editing was
concentrated in PowerPoint. Wave 6 and the 2026-10-26 delivery changed that; the observations in
section 5.1 are historical and the current shared status is recorded above.

### 5.1 Historical survey state

Tracked, non-test source lines:

| Product | Shared implementation                   | Per framework                                             |
| ------- | --------------------------------------- | --------------------------------------------------------- |
| pptx    | `src/ui/src/pptx` 211k                  | React 106k, Angular 96k, Vue 82k, Svelte 66k, vanilla 64k |
| docx    | `src/ui/src/docx` 38k, `bindings` 1.3k  | 80 to 200                                                 |
| xlsx    | `src/ui/src/xlsx` 27k, `bindings` 1.2k  | 90 to 250                                                 |
| visio   | `src/ui/src/visio` 35k, `bindings` 1.3k | 150 to 290                                                |
| teams   | `src/ui/src/teams` 17k, `bind.ts`       | 50 to 130                                                 |

- **pptx bindings are ports of each other.** Roughly 20 to 25% of each binding (80k to 100k lines
  in all) is framework-free logic: the state machine around each shared policy. Drift found: undo
  depth (React and Vue 120, Svelte and vanilla 100, React keeps its own stack), two autosave engines
  (interval in React and Angular, debounce elsewhere), React and Angular outside
  `createGestureController`, the AI bridge copied into every binding, five keyboard dispatchers with
  differently named handlers. Per binding, the chart and SmartArt files alone are 3.5k to 7.9k lines
  (each has its own data editor, insert SmartArt dialog and text pane).
- **docx, xlsx, visio and teams each wrote their own binding kit** (prop tracker, change-only
  forwarding, deferred handle, contract mapping, near-identical Vue, React and Solid adapters).
- **Chart and SmartArt editing was concentrated in pptx at the survey's start.** Reading went through `chart` and `diagram` for every
  product, and docx, xlsx and visio drew charts through `chart/render`, but pptx edited
  `PptxChartData` through `builders/sdk/chart-*` and saves through about 127 object-tree writers;
  Excel used its own `ChartPatch`; Word and Visio did not yet edit charts. SmartArt node
  operations, reflow, text projection and part writers lived in `src/core/pptx`; Word and Excel showed
  the cached drawing only. Wave 6 and the later shared delivery superseded these findings.

### 5.2 Target layers

1. **Core** (DOM-free): model, editing operations and policies.
2. **`ooxml-ui`** (no framework): headless controllers (`getSnapshot`, `subscribe`, typed actions),
   imperative leaf renderers and shared Lit elements, including the chart and SmartArt editing
   elements every product mounts.
3. **One store adapter per framework**, written once for every controller (React
   `useSyncExternalStore`, Vue `shallowRef`, Svelte runes, Angular signals, Solid).
4. **Native templates** on top, so each framework keeps its idioms (slots, render props,
   `v-model`, DI, SSR). Bindings hold templates, lifecycle, focus and portals only.

The docx, xlsx, visio and teams wrappers share one contract-driven kit in `ooxml-ui/bindings`, with
generic React, Vue and Solid factories (Angular and Svelte need their compilers, so their small
adapters stay per product over the same kit) and one contract-parity test per product.

### 5.3 Chart and SmartArt order

Status at the end of Wave 6 (2026-10-23): the steps below have landed in part. The open items are the
Wave 6 "Next" list and the "Remains" entries of [agnostic-core-plan.md](agnostic-core-plan.md) sections
3.5 to 3.7, which define the step labels (C1 to C6, U1 to U5, S1 to S5).

Charts, core: neutral edit operations over `ChartSpace` (`chart/edit`, with Excel's `ChartPatch` as
a client), then the model gaps, then pptx save through `writeChartSpace` one feature family per
commit behind an old-versus-new part diff over every committed deck, then `PptxChartData` as an
adapter, then a ChartEx model, parser and writer, then Word chart editing (edited parts plus the
embedded workbook), then Visio editing of embedded charts through the same operations.

Charts, UI: a shared format-pane model and element (from `src/ui/src/pptx/render/chart-format`),
shared chart design galleries, and one insert-chart picker (`office-ui-chart-picker`) and one data
sheet (`office-ui-chart-data-sheet`) for every product.

SmartArt: a neutral paragraph and run model in `drawingml`, then node text projection into
`diagram`, then a neutral drawing shape and edit operations (`diagram/edit`), then part writers
(`diagram/write`), then Word and Excel SmartArt editing, with a shared insert picker, galleries and
text pane.

### Wave 6 (landed 2026-10-23)

Framework layers:

- `ooxml-ui/bindings`: one contract-driven binding kit with React, Vue and Solid factories; the
  docx, xlsx, visio and teams bindings run on it, with a contract-parity test per product.
- pptx headless controllers (`src/ui/src/pptx/render/controller/`) with one store adapter per
  binding: undo history (one default depth, 100), keyboard dispatch, AI bridge and panel, autosave
  (one engine, uncommitted inline text included), recovery, collaboration provider and session,
  canvas gestures (React and Angular moved onto `createGestureController`), group selection and
  slide show steps. Inspector Effects, Line, Size and Position and Text Box sections are field
  descriptors with one generic renderer per binding.
- The per-binding chart data grids, insert SmartArt dialogs and SmartArt text panes are gone in
  favour of shared elements.

Charts: neutral edit operations and typed intents (`chart/edit`), a ChartEx model, parser, writer
and edits (`chart/ex`), the embedded workbook writer (`chart/workbook`), one Format pane
(`office-ui-chart-format-pane`, its own translations), one data sheet, one insert picker and shared
Chart Design galleries. pptx saves title, legend, axis and series data edits of classic charts and
title, legend, label and series edits of ChartEx through the neutral writers, behind save-parity
harnesses. Excel edits over `ChartSpace` and never rewrites a part the model cannot carry. Word
inserts, edits, formats, resizes and saves charts; Visio edits inserted charts. Word and Excel
results are checked in desktop Office by acceptance scripts.

SmartArt: a neutral DrawingML text model, node text projection, drawing view, edit operations and
data writers in `diagram`, built-in definitions in `diagram/builtin`, and one picker, text pane and
design command list. Word and Excel insert and edit SmartArt; desktop Word and Excel rebuild the
layout from the data part, so an edit drops the cached drawing.

The Wave 6 next list was a historical checkpoint. Labels/paint saves, Word/Excel ChartEx editing,
SmartArt styles/colours/generated caches and shared advanced chart controls have since landed.
Current work finishes the PowerPoint compatibility-adapter migration and its exhaustive admission,
corpus and native checks. The presentation controllers, the Fill inspector section and the pptx
ribbon tab row landed on 2026-11-16 (below).

The Fill inspector section landed on 2026-11-16: `SHAPE_FILL_MODE_SECTIONS` and
`SHAPE_FILL_OPTIONS_SECTIONS` (`render/inspector/shape-fill-fields.ts`) are drawn by each binding's
generic field renderer, and `resolveShapeFillView` with the typed `shapeFillPatch` actions
(`shape-fill-actions.ts`, stop edits in `shape-fill-stops.ts` over `ooxml-core/chart`'s stop list
operations) drive the compound controls each binding keeps as a native template: the solid colour
with its theme grid, the stop rows, the pattern swatches and the picture source. The bindings had
drifted (React a five-kind select, Vue four kinds and no picture, Svelte and vanilla gradient and
pattern checkboxes, Angular closed disclosures with no kind at all; stop opacity 0-100 in React,
0-1 elsewhere; add-stop colour and position four ways); every binding now offers PowerPoint's six
kinds, shows the solid colour only for a solid fill, transparency 0-100 for a filled shape, per-stop
opacity 0-100 and a pattern swatch grid, and adds a stop halfway along the last gap. A slide
background fill saves as `useBgFill` and draws at once (below).
See the active delivery ledger and agnostic plan for implemented coverage and residual limits.

The Alt Text and picture fields followed (2026-11-16): `ELEMENT_ALT_TEXT_SECTIONS`
(`render/inspector/element-alt-text-fields.ts`, description and title, gated per kind by
`element-non-visual-description.ts`), `PICTURE_ADJUSTMENT_SECTIONS` and `PICTURE_CROP_SECTIONS`
(`picture-fields.ts`: brightness, contrast, saturation, transparency, black and white threshold,
grayscale and the four crop insets) are drawn by the generic renderers, which gained a `text` field
kind (a text box, or a text area when `multiline`, with a placeholder). React and Vue lost
`ImageAdjustmentsPanel`, React `RangeSlider`, and vanilla's per-field image, crop, alt text and
title handlers became one `patchSelectedElement` (its unused gradient and fill opacity handlers
and gradient state went with them). Drift found: Angular showed the stored `alphaModFix` (opacity)
as Transparency, so an opaque picture read 100%; crop insets stopped at 80% in React, Vue and
Angular and 90% in Svelte and vanilla; Angular's crop captions were untranslated English; Svelte and
vanilla had no Grayscale; a zero threshold wrote a bi-level effect (a white picture) everywhere but
Svelte; vanilla committed alt text on change, the others per keystroke; Angular's Alt Text panel
and the Vue and Angular adjustments ignored `canEdit`. Every binding now shows transparency
inverted, as PowerPoint does, crops to the shared 90% clamp on a slider, clears an effect set back
to zero, applies alt text as typed and disables the fields with the pane.

The rest of the inspector followed the same day, one section per commit, each drawn by the
generic renderers and checked per binding:

| Section                | Shared descriptors (`render/inspector/`)                   | Drift removed                                                                                                                       |
| ---------------------- | ---------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| Paragraph and font     | `PARAGRAPH_SPACING_SECTIONS`, `CHARACTER_SPACING_SECTIONS` | each binding showed a different subset (Vue none); all showed model units (px, hundredths of a point) where PowerPoint shows points |
| Table                  | `TABLE_OPTIONS_SECTIONS`                                   | three option orders; band sizes, direction and cell margin each missing somewhere; "Last Row" for PowerPoint's Total Row            |
| Picture colour effects | `PICTURE_EFFECT_SECTIONS`, `pictureDuotonePresetPatch`     | recolour only in React and Vue; duotone written on first touch in vanilla and Angular; preset labels humanised from keys            |
| Slide Background card  | `SLIDE_BACKGROUND_SECTIONS`                                | Hide background graphics missing in three cards; a colour on a picture background was saved but never shown                         |
| Media playback         | `MEDIA_PLAYBACK_SECTIONS`                                  | no fades in vanilla; Angular's Play Across Slides did not start the clip; switched-off flags written as `false`                     |
| OLE card               | `OLE_OBJECT_NAME_SECTIONS`, `resolveOleSummaryRows`        | the name was trimmed per keystroke (a typed space vanished); Angular's group card was untranslated English                          |
| Action Settings        | `ACTION_SETTINGS_SECTION`, `actionSettingsTargetOf`        | five slide pickers (number box or select, titled or not); the custom show and Return after showing missing in some bindings         |
| Media trim             | `MEDIA_TRIM_SECTION`                                       | mm:ss text boxes in React, seconds or milliseconds elsewhere; the end bound ignored the live clip duration                          |
| Table cell formatting  | `TABLE_CELL_SECTIONS`                                      | untranslated B I U and L C R letter buttons; one border width and margin for all sides in vanilla; diagonals only in Vue; px units  |
| Artistic effects       | `PICTURE_ARTISTIC_SECTIONS` (a `gallery` field)            | thumbnails in React and Vue, a select elsewhere; Vue's gallery never disabled; Grayscale left the grayscale toggle on in three      |

The field model resolves what depends on the deck or the clip at render time: a field's
`optionsOf(target)` replaces its static choices (the deck's slides and custom shows, with a verbatim
`label` beside the translation key), `maxOf(target)` its upper bound (the trim end against the live
duration), and `previewImageOf(target)` a gallery's picture. Each generic renderer rebuilds a
section when its resolved choices change.

A slide background fill now draws at once: `getSlideBackgroundStyle` publishes the resolved
background as `--pptx-slide-bg-*` custom properties on every stage, and `resolveComputedFill`
paints a `useBackgroundFill` shape that carries no resolved copy with them, offset by the shape's
position (`render/slide-background-fill.ts`; React's stage adapter passes the properties through).
A framework-neutral spec (`slide-background-fill-live.spec.ts`) checks the five demos.

Still per binding, by choice or for lack of time: the table cell text and fill colour pickers with
their theme grids and the gradient and pattern cell fill, the table style editor (its fields
already come from the shared table-style-editor descriptor, but it edits style definitions per
table part rather than one value through a patch), Merge and Split, and the docked Animation and
Transition panes. Chart and SmartArt formatting already go through the shared elements
(`office-ui-chart-format-pane`, the SmartArt picker and text pane).

### PowerPoint tab row on `office-ui-ribbon` (landed 2026-11-16)

The five PowerPoint bindings draw their ribbon tab row with `pptx-ui-ribbon-tabs`, the shared
`office-ui-ribbon` under the pptx tag (token bridge, `tab-select` and `collapse-toggle` events).
`office-ui-ribbon` gained `items`, a tab row whose panels live elsewhere: the host owns the
selection and hears every choice, so a click on the selected tab still peeks a collapsed ribbon.
One descriptor, `buildRibbonTabRow` (`src/ui/src/pptx/render/ribbon-tab-row.ts`), decides the tabs
(fixed tabs with File first, add-in tabs, then the contextual tabs such as Shape Format, Picture
Format, Chart Design and Table Design), their labels, ScreenTips and kinds, and the selected tab;
`hiddenActions` and Customize Ribbon filter it. Each binding builds that state from its own stores
and routes the two events. Record, Comments and Share, and Ribbon Display Options stay native
templates in the element's `actions` slot; the tab panels and commands are unchanged. Ribbon
Display Options keep their own controller (`createRibbonDisplayController`), so the element's own
collapse button is not used. Binding tests reach the tabs through `ribbonTabButtons` and
`ribbonTab`, since they live in the element's shadow root.

### Presentation controllers (landed 2026-11-16)

Slide-show state is written once in `src/ui/src/pptx/render/controller/` and read by every pptx
binding through its store adapter:

- `createPresentationRehearsalController`: Rehearse Timings (per-slide and total clocks, pause,
  recording, the 250 ms display tick, the summary). Saving goes through `applyRehearsalTimings`.
- `createPresentationInkController`: the show's pen, highlighter, eraser and laser, per-slide
  strokes, Ctrl+M markup visibility and the popup toolbar's auto-hide; `applyPresentationInk`
  keeps ink as slide elements.
- `createPresenterSessionController`: the audience window, deck hand-over, the BroadcastChannel
  protocol on both sides, Swap Displays and the shared snapshot.

Drift found between the bindings and the one behaviour chosen:

- Rehearsal: slide floors of 1 s (Svelte) and 250 ms (vanilla) and none elsewhere; Svelte also set
  "On mouse click" when saving; React, Vue and Angular counted paused time into the total, vanilla subtracted
  a pause twice, and vanilla timed the editor instead of running a show. Now: no floor, both clocks
  stop while paused, "On mouse click" is left as it was, and every binding runs the show.
- Ink: Svelte cleared every slide (and disarmed) on "Erase All Ink on Slide" and erased on hover;
  the pointer options of the context menu and keyboard toggled the tool off in React, Vue and
  Angular, and an Angular colour pick could disarm the pen; React built its kept ink elements by
  hand. Now: the toolbar's tool buttons toggle, every other path arms, Erase clears the current
  slide, the eraser needs a press, and every binding keeps ink through `applyPresentationInk`.
- Presenter window: only React and Angular noticed a popup the user closed; Svelte refused to open
  without a deck while vanilla and the others navigated; Angular answered "audience ready" with the
  slide captured at open; only Svelte, Vue and vanilla numbered their messages. Now: the popup is
  polled, it is navigated with or without a deck, "ready" gets the current snapshot, and the stored
  deck is cleared on close, unload and exit.

Still per binding: the HUD, summary, toolbar and overlay templates, fullscreen requests inside the
user gesture, and vanilla's DOM ink overlay (unit-square strokes, its own gesture tracking), which
stores its strokes in the shared ink controller.

The show toolbar's "has annotations" flag (it enables Erase All Ink on Slide) counted only the
current slide in React and Angular and every slide in Vue, Svelte and vanilla. Every binding now
counts every slide, like the keep-or-discard prompt on exit. The button still clears only the
current slide, so it can be enabled on a slide that has no ink of its own.

### Slide Sorter and Outline view state (landed 2026-11-16)

A survey of the five bindings outside the inspector, ranked by the framework-free lines each kept
around an already shared policy, put the Slide Sorter first (React about 380 lines with its keyboard
hook, Vue about 190, Angular about 150, vanilla about 140, Svelte about 60) and Outline view second
(about 60 to 110 each). Both now run once in `ooxml-ui/pptx`:

- `createSlideSorterController` (`render/controller/slide-sorter-controller.ts`): selection and
  anchor, the sorter clipboard, zoom, the right-click menu, drag reorder and the keyboard, with
  `slideSorterView` for the drawn values. Deck operations reach the host in descending index
  order; each binding keeps its template, the drag `dataTransfer` and its listener lifetime.
- `createOutlineViewSession` (`render/outline-view-session.ts`): refuse on a read-only deck,
  apply, skip a no-op, commit, and hand the caret to the row the edit names
  (`focusOutlineRow`); the binding only chooses when to restore focus after its render.

Drift found and the behaviour chosen:

- Sorter: React opened the menu (and swallowed the native one) on a read-only deck; Vue and
  vanilla reordered a read-only deck by dragging and Svelte's move buttons stayed live; Angular and
  Svelte had no drag reorder; only vanilla undid and redid inside the sorter; React left Escape
  unprevented. Now: the menu, drag and deck-writing keys are refused on a read-only deck, every
  binding drags (Angular gained a `moveSlide` output), undo and redo inside the sorter run when the
  host passes them (vanilla does today), and every handled key is prevented and stopped.
- Outline view: three row lookups (an escaped selector in React, Vue and Angular, an attribute
  scan in Svelte, a map of inputs in vanilla). Now one attribute scan.

Non-test binding source after this work (2026-11-16, this change alone): React 104,367 to 104,066,
Angular 92,113 to 92,093 (it gained drag reorder), Vue 80,113 to 80,014, Svelte 66,446 to 66,358,
vanilla 65,996 to 65,927.

### Comments panel, Selection Pane and the show's Erase button (landed 2026-11-16)

- `createCommentsPanelController` (`render/controller/comments-panel-controller.ts`), one per
  viewer: the new-comment draft and its mentions per slide, the comment being edited per slide,
  one reply composer that opens empty, and add, edit (at any depth), delete, resolve and reply
  through the binding's undoable deck commit. React's behaviour is the one chosen, because the
  other bindings lacked features rather than differed: Vue, Angular and Svelte kept one draft for
  the whole deck (it followed you to the next slide) and dropped it when the panel closed, had no
  edit in place, and only React anchored a new comment to the selected element; vanilla's
  workspace pane and inspector tab kept separate drafts. Vue provides the controller to every
  `CommentsPanel`, Angular through `CommentsPanelStateService`, Svelte through `commentsPanelFor`;
  a panel mounted outside a viewer lists comments read-only.
- `createSelectionPaneController` with `selectionPaneRows` and `selectionPaneLabel`: the rename
  box and drag reorder, and one row label (name, else text, else the translated "{type} {number}").
  Drift: React labelled in English whatever the locale, vanilla capitalised the type id, Angular
  showed the element id, Svelte the lower-case type numbered from the top; the rename box started
  from the label, the id or the bare name; Angular renamed and vanilla dragged on a read-only
  deck; Angular and Svelte could not drag. The box now closes before focus returns to the viewer
  root, so the blur that focus move causes never commits a cancelled rename.
- The show toolbar's state gains `hasSlideAnnotations`: Erase All Ink on Slide is enabled for ink
  on the slide on screen only, while `hasAnnotations` keeps meaning ink anywhere in the show.
  Vanilla's Erase (button and E) cleared every slide's ink; it now clears the slide on stage.

Non-test binding source after this work (2026-11-16, from 104,367 /
92,113 / 80,113 / 66,446 / 65,996): React 103,395, Angular 92,172, Vue 79,816, Svelte 66,438,
vanilla 66,002. Angular, Svelte and vanilla grew where they gained comment editing and drag
reorder; their per-binding state machines went.

Angular keeps its public `CommentsService`, now unused by the viewer.

### Find and Replace bar, slide rail and zoom (landed 2026-11-16)

- `createFindReplaceController`: the bar's open state, its mode (Ctrl+F find, Ctrl+H find and
  replace) and the search session. Ctrl+F and Ctrl+H toggle the bar and the other chord switches an
  open bar's mode; opening searches the kept query again; closing drops the matches and keeps the
  query, replacement and scopes. The building blocks' `createFindBarController` is now this
  controller with its port. Drift: Ctrl+H after Ctrl+F closed the bar in React, Vue, Svelte and
  vanilla; React did not search again on reopen; Angular's Ctrl+F opened a second, find-only bar
  with its own search outside the shared session, and its Ctrl+H only opened the replace bar.
  Angular's viewer now shows the one shared bar in both modes (`FindBarComponent` stays exported).
- `createSlideRailController` with `slideRailMenuEntries`: the thumbnails pane's Ctrl/Shift
  selection by slide id and its right-click menu (whole selection when right-clicked inside it,
  that slide alone outside it, refused on a read-only deck) with command routing. Five copies of
  this state went; the read-only gate had sat in three different places.
- Zoom: the zoom state, gestures and per-slide scroll reset were already shared
  (`viewer-zoom-store`, `attachViewportNavigation`, `resetViewportScroll`). React's Zoom to
  Selection computation moved to `zoom-to-selection` (`selectionZoomBounds`,
  `zoomToSelectionScale`, `selectionScrollTarget`). No binding offers a Zoom to Selection control
  (React's hook action is not wired to one); adding it is a feature for all five.

Non-test binding source after this work (from 104,367 / 92,113 / 80,113 / 66,446 / 65,996):
React 103,318, Angular 92,165, Vue 79,804, Svelte 66,398, vanilla 66,019.

Still per binding: the notes pane's draft. Each binding keeps the note's text and segments while
typing, debounces or not (React, Vue and Angular's rich editor debounce; Svelte, vanilla and
Angular's plain textarea save on change or blur), and decides when a slide switch or an echo of
its own save reseeds the editor. A shared draft controller (load per slide id with an echo guard,
debounced edit, flush on blur and before a slide switch, saving to the slide it was typed on) is
the next step; the debounce-for-all decision and each binding's contentEditable caret handling
make it a change of its own.
