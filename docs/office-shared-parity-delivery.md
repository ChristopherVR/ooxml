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

- [ ] Data-label visibility/content, per-series and per-point overrides, value-from-cells labels.
- [ ] Series/point fills, lines, gradients and transparency, preserving unrelated theme tokens.
- [ ] Rich title runs, legend entries, axis titles/fonts, gridlines and number/date/display units.
- [ ] Series structure, type switches and combinations, trendlines, error bars and markers.
- [ ] Layouts, 3D/chrome, data tables, stock lines/bars, protection and supported extensions.
- [ ] Neutral ChartEx reading, editing, generation and family switching.
- [ ] Make the PowerPoint model a compatibility adapter over the neutral models; remove superseded
      writer paths only after their admitted cases have regression coverage.
- [ ] Run the XML/model corpus harness with every difference classified, then native reopen and
      embedded-workbook refresh checks. Do not normalize away authored semantics to match a baseline.

## 5. Shared SmartArt style and colour authoring

- [ ] Neutral style/colour operations and definition writers preserve unknown XML.
- [ ] Share cached-drawing generation independently of native Office cache regeneration.
- [ ] Word and Excel mount shared Styles and Change Colors galleries with history and persistence.
- [ ] Cover per-node overrides, rich text, host theme changes and save/reopen.
- [ ] Keep Visio diagram conversion a separate host feature with explicit native-shape or embedded
      semantics; sharing diagram primitives is not native Visio SmartArt support.

## 6. Shared chart controls and ChartEx consumers

- [ ] Complete shared formatting intents and controls for effects, 3D, text, richer fills,
      secondary axes, trendlines, error bars, markers and axis titles.
- [ ] Wire the supported controls in PowerPoint, Word, Excel and embedded Visio charts.
- [ ] Enable Word/Excel ChartEx data-sheet editing, family options and structural edits.
- [ ] Synchronize workbook references/caches with safe external-source and hierarchy policies.
- [ ] Verify all applicable framework bindings, keyboard/history behavior and native save/reopen.

## Integration and evidence

- [ ] Reconcile `roadmap.md`, `agnostic-core-plan.md`, `structure-plan.md`, product capability
      pages and PowerPoint ribbon reviews with the delivered implementation.
- [ ] Run the repository CI planner's verification from the isolated integration worktree and
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

The integrated foundations plus ChartEx storage passed 1,085 tests across 145 files and both core
compiler projects. ChartEx UI integration and native acceptance are still in progress.

### Further neutral chart saves, 2026-10-26

Shared operations now preserve styled title runs, axis source-linked formats and date/label/display
units, hidden-label content and value-from-cells extension siblings. PowerPoint group spacing maps
to the neutral schema-validated operation. Formatting-only changes leave embedded workbook cells
alone; data edits retain blank cells. The title/axis increment passed the 300-case saved-part corpus
plus 41 focused tests. Label-range and group edits have focused preservation regressions; expanded
corpus and native acceptance remain required before closing priority 4.
