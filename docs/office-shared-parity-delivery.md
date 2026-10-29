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

These checks cover supported source semantics and operations, not native pixel fidelity. Wider
3D/XY combination switches, raw geographic/Pareto creation, arbitrary ChartEx custom label text or
manual layout, and full removal of the PowerPoint source-tree model remain outside this delivery.

## 5. Shared SmartArt style and colour authoring

- [x] Neutral style/colour operations and definition writers preserve unknown XML.
- [x] Share cached-drawing generation independently of native Office cache regeneration.
- [x] Word and Excel mount shared Styles and Change Colors galleries with history and persistence.
- [x] Cover per-node overrides, rich text, host theme changes and save/reopen.
- [x] Keep Visio diagram conversion a separate host feature with explicit native-shape or embedded
      semantics; sharing diagram primitives is not native Visio SmartArt support.

## 6. Shared chart controls and ChartEx consumers

- [x] Add shared formatting intents and controls for effects, 3D, text, pattern fills,
      admitted secondary axes, trendlines, error bars, markers and axis titles.
- [x] Wire the supported controls in PowerPoint, Word, Excel and embedded Visio charts.
- [x] Enable admitted Word/Excel ChartEx data-sheet, family-option and structural edits.
- [x] Synchronize workbook references/caches with safe external-source and hierarchy policies.
- [x] Verify applicable bindings, history and save/reopen, including native Word/Excel ChartEx cases.

These checks cover the admitted controls, not every Office formatting option. Picture media
relationships remain host-owned. Secondary-axis authoring keeps the first flattened series as
the primary rendering anchor; horizontal combinations, stacked multi-series splits and unknown
group-extension splits are refused. The pane edits the first trendline, disables ambiguous multiple
trendlines, and preserves custom error sources without offering a range editor. ChartEx exposes its
own supported controls and refuses classic-only advanced requests.

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
