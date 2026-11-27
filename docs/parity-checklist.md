# Office parity checklist

The single list of what is still missing for 1:1 parity with Microsoft Word, Excel, PowerPoint
and Visio. It was built on 2026-10-30 from a code review of `main` at `26453c645` (reading and
searching the source, not running it) and replaces no evidence: the per-product trackers below
keep the detail, the native measurements and the history of what landed.

**No product has reached Office parity.** Each product works within the scope its tracker
describes; every unticked line here is a known gap in that scope.

## How to use this list

- One line per gap, under its product and area. Severity: **P0** breaks or loses data in common
  files or blocks a core workflow, **P1** a common feature is missing, **P2** niche.
- The path after each item is where to start; it is the evidence the gap was found from.
- Tick an item in the commit that closes it, and link the test or native evidence. Ticked items
  are dropped at the next review, and the product tracker keeps the record.
- New gaps go here first, then into the product tracker's detail if they need it.
- Report gaps honestly in the product too: the Excel Feature Status dialog
  (`src/ui/src/xlsx/dialogs/help-dialogs.ts`), load warnings (`src/core/docx/parse-warnings.ts`,
  `data.warnings` in PowerPoint) and disabled ribbon commands with a reason.

Product trackers: [Excel](xlsx-parity-status.md), [Word](word-parity-review.md) (with the
viewer's [roadmap](../viewers/docx/docs/parity-roadmap.md) and
[outstanding work](../viewers/docx/docs/outstanding-work.md)), PowerPoint
([limitations](../viewers/pptx/docs/guide/limitations.md),
[ribbon parity](../viewers/pptx/docs/guide/ribbon-parity.md) and
[other tabs](../viewers/pptx/docs/guide/ribbon-parity-other-tabs.md),
[OpenXML conformance](../viewers/pptx/docs/architecture/openxml-conformance.md)), Visio
([review](visio-parity-review.md), [ledger](../viewers/visio/docs/parity.md)), and the shared
layer ([delivery](office-shared-parity-delivery.md), [chart gaps](chart-renderer-gap-report.md)).
Teams has its own list in [teams-parity.md](teams-parity.md).

## Shared layer and every product

- [ ] **P2** The opt-in cards look (`data-office-look="cards"`, the suite Simple view default) is
      uneven: Vanilla and Vue keep some bordered inner cards, not every field row becomes the
      two-column grid, and selects, submenus, the font picker and the canvas context menu keep
      their own radius and shadow. In Simple view Word, Excel and Visio also pick up the pill
      status groups and 12px menus.
- [ ] **P2** Japanese and Korean exist only for PowerPoint: the suite's shared language list
      (`apps/office-suite/src/settings.js`) and `SUPPORTED_OFFICE_LOCALES` do not offer them, and
      Word, Excel, Visio, Teams, the diagram strings, the shared options and ChartEx tables have
      no ja/ko. CJK chrome fonts apply only when the host sets `lang`.
- [ ] **P2** PowerPoint's Japanese and Korean animation names, newer SmartArt layout names, and
      bevel, material and 3-D camera presets still need review by native speakers.
- [x] **P1** OneDrive save recovery only offered choosing/copying versions: format-aware three-way
      mergers now combine supported independent DOCX/XLSX/PPTX/VSDX edits automatically and rebase
      against the latest revision. The suite reloads the acknowledged combined file. Evidence:
      `src/core/embedding/onedrive/merge-session.test.ts`, each format's `merge-snapshots.test.ts`
      and `scripts/test-support/onedrive-merge-checks.mjs`.
- [ ] **P1** Ambiguous structural or same-content overlaps in OneDrive saves require
      retained-version recovery. See [concurrent save merging](onedrive-integration.md#automatically-combine-concurrent-edits-during-save).
- [ ] **P1** OneDrive collaboration with Microsoft Office lacks a live operation stream.
      [Concurrent save merging](onedrive-integration.md#automatically-combine-concurrent-edits-during-save)
      combines compatible saved snapshots.
- [x] **P1** Fonts that are neither installed nor embedded were downloaded only in PowerPoint:
      Word, Excel and Visio now share its webfont fallback (`ooxml-core/text` `webfonts/`,
      `ooxml-ui/fonts/webfonts.ts`), with metric-compatible stacks wherever they measure and paint
      text and a bounded wait before showing a file. Off unless the host sets `googleFonts`; the
      suite turns it on from Settings > Fonts. Evidence: `docs/font-loading.md`, each product's
      `webfont-runs.test.ts` and `webfonts.test.ts`.
- [x] **P2** Word, Excel and Visio had no Trust Center toggle for font downloads: each now has
      PowerPoint's "Download fonts that are not installed or embedded" in File > Options > Trust
      Center (off by default; a host `googleFonts` value locks it). Evidence:
      `src/ui/src/fonts/font-download-policy.test.ts` and each product's Trust Center tests.
- [x] **P2** In Word, Excel and Visio a font picked while editing is requested immediately without
      waiting for the next load (`webfontRunsKey` in `src/ui/src/fonts/webfonts.ts`,
      `DocxWebfonts.refreshIfChanged`, `XlsxWebfonts.refresh`, `ViewerWebfonts.loadIfChanged`;
      evidence: `src/ui/src/docx/docx-webfonts.test.ts`, `src/ui/src/xlsx/webfonts.test.ts`,
      `src/ui/src/visio/viewer-webfonts.test.ts`).

| Capability    | Word         | Excel        | PowerPoint                 | Visio                 |
| ------------- | ------------ | ------------ | -------------------------- | --------------------- |
| Chart render  | 2-D and 3-D  | 2-D and 3-D  | yes, own renderer with 3-D | 2-D and 3-D, embedded |
| Chart editing | yes          | yes          | yes                        | partial               |
| ChartEx       | partial      | partial      | partial                    | partial               |
| SmartArt      | render, edit | render, edit | render, edit               | no (by design)        |
| Equations     | yes          | no           | yes                        | no                    |
| Proofing      | browser only | none         | browser only               | browser only          |

- [x] **P0** 3-D charts draw flat in Word, Excel and Visio: `src/core/chart/render/chart-3d-*.ts`
      now projects `c:view3D` (oblique for right-angle axes, perspective otherwise) and paints
      bar3D, line3D, area3D, pie3D and surface3D with depth ordering and face shading; PowerPoint's
      projection and shading code is shared with it.
- [ ] **P1** 3-D chart detail in the core painter: floor and wall formatting (`chart-3d-walls.ts`),
      bar `c:shape` (`cylinder`, `cone`, `coneToMax`, `pyramid`, `pyramidToMax`;
      `chart-3d-shapes.ts`), `pie3D` series/point explosion (`chart-3d-pie.ts`), `surface3D`
      wireframe and 2-D `surfaceChart` contour (`chart-3d-surface.ts`), and projected 3-D data
      labels (`chart-3d-labels.ts`; `chart-3d-detail.test.ts`) are implemented. Still open:
      pixel-level measurement against Office renders.
- [x] **P1** Pareto charts draw in the shared painter (bars in descending order, the cumulative
      line on a 0 to 100% secondary axis; `chart/pareto.ts`, `chart/ex/create-family.test.ts`), and
      Word, Excel and PowerPoint insert pareto charts as Office writes them (an owned `paretoLine`
      series and a percentage axis, observed from Excel 16's `AddChart2(-1, xlPareto)`).
- [x] **P1** Word and Excel ChartEx editing covers axis visibility, title, gridlines, number format
      and display units, series names, histogram and pareto bins, funnel gap width and region map
      projection, area and labels, with native re-save evidence (delivery ledger, 2026-11-16).
      Secondary axes, custom label text and manual label positions stay refused for ChartEx.
- [ ] **P1** Region maps: Word, Excel and PowerPoint insert them (`cx:geography` with culture,
      projection and area, no provider cache) and the shared painter draws them over bundled,
      approximate country outlines with a translated notice. A new map has no `cx:geoCache`, and
      Office 16 draws it empty until Office itself adds map data through its online service; this
      code never fetches geodata. Exact shapes need a cache Office wrote.
- [x] **P1** Visio ChartEx path: embedded Excel ChartEx workbooks (all 8 families, including
      standalone `xl/charts/chartEx*.xml`) are read, rendered via the shared SVG painter, and
      insertable/editable in Visio's Insert Chart dialog (`visio/prepare-embedded.ts`,
      `src/ui/src/visio/render-embedded-chart.ts`, `viewer-insert-chart-ex.ts`; verified in
      `edit-chart-object-chartex.test.ts`, `render-embedded-chart.test.ts` and
      `viewer-insert-chart.test.ts`).
- [ ] **P1** Pivot charts: field buttons and the link to PivotTable edits and refresh work for
      charts with one plot group over a supported PivotTable (`xlsx/pivot/chart-binding.ts`);
      open: imported charts with several plot groups or unsupported PivotTable sources, and
      per-role field-button visibility.
- [ ] **P1** Chart Design: Word and Excel now have Add Chart Element (Title, Legend, Data Labels),
      Switch Row/Column (`switchChartRowColumn` in `src/core/chart/edit/data-sheet.ts`,
      `switchXlsxChartRowColumnPatch` in `src/core/xlsx/edit/chart-design.ts`) and Select Data
      (`chart-design-ribbon.ts`, `xlsx/ribbon/tabs/contextual.ts`; evidence: `data-sheet.test.ts`,
      `chart-ui.test.ts`, `chart-design.test.ts`). Still open: Add Chart Element, Select Data and
      Change Chart Type in PowerPoint.
- [ ] **P2** Edit Data row removal in Word, PowerPoint and Visio matches Excel's chart (ranges,
      counts and every point, measured against hidden Excel), and the embedded workbook's cells
      below the row, inside and outside the chart's ranges, move up with it; removing a series
      shrinks the table and clears its old last column (`chart/workbook/workbook.test.ts`). Open:
      a moved formula keeps its references (`SUM(B2:B3)` stays as written), where Excel rewrites
      them for the deleted row; the chart area cannot use `xlsx/formula/shift-formula.ts` without
      a shared reference shifter.
- [ ] **P1** Equations: Excel inserts, edits and saves standalone Office Math text boxes with the
      editor Word uses (`src/ui/src/math/`); open: equations in grouped Excel text boxes, legacy
      OLE equation previews, and authoring equations in Visio (Visio previews and keeps equations
      embedded from Word; desktop Visio 16 disables its own equation command).
- [ ] **P1** SmartArt layouts the preview marks approximate (`src/core/diagram/layout-preview.ts`;
      an edited diagram's regenerated cache binds to Office's presentation points but its
      placement is the engine's, within 1% of Office in 27 of 52 admitted gallery diagrams),
      3-D rendering in Word and Excel (approximate: a lit SVG bevel, stacked extrusion copies and
      the Scene styles' camera as a per-shape affine projection; no hidden surfaces, material
      model or contour, and not compared against Office's pixels) and bent hierarchy
      connectors; the built-in catalogue has 176 layouts (`diagram/builtin/catalog.ts`), not yet
      checked against Office's gallery.
- [x] **P1** SmartArt Change Colors and SmartArt Styles offer Office's full built-in catalogue in
      Word, Excel and PowerPoint: 38 colour and 14 quick-style definitions (nine 3-D), captured
      from PowerPoint 16 through COM (`diagram/builtin/style-*.ts`). A pick writes Office's own
      definition and `uniqueId`; Word and Excel generated caches carry each label's bevel,
      extrusion, material, light rig and text 3-D as Office's cache does. 82 Word 16 and Excel 16
      files reopened natively with the expected ids (2026-11-16). Remaining: the gallery tiles
      are hints, not renderings; PowerPoint previews a pick's 3D at once but its colours and
      effects at the nearest model intensity and colour patch until reopen. Word and Excel show the gallery names in every locale from
      PowerPoint's translations, which are agent-authored and unverified against Office's own
      localized names.
- [ ] **P1** Proofing is the browser's `spellcheck` attribute everywhere: no grammar, thesaurus,
      custom dictionary, ignore list or proofing-language engine.
- [x] **P1** Excel Picture Format and Shape Format contextual tabs: styles from the shared
      DrawingML catalogues (42 theme styles, 35 presets, 28 picture styles), fill, outline,
      effects, size, Alt Text, Reset and Compress for one or several drawings; Excel 16 reopens
      them (`xlsx/edit/drawing-format.ts`, 2026-11-26).
- [ ] **P1** Excel draws Picture Style inner shadows, reflections, soft edges, compound borders
      and 3-D effects flat in the worksheet (they are saved).
- [ ] **P1** Picture Format and Shape Format contextual tabs are partial in Word Picture Format (Wrap Text, Crop and Format only; Arrange is shared with Shape Format)
      and in PowerPoint Shape Format (Arrange is done; Size and Alt Text are open).
- [x] **P1** Shared chart Format pane analysis: several trendlines per series (a section each, Add
      trendline) and custom error-bar sources as values or a range (`chart/edit/error-bar-source.ts`;
      Excel reads its workbook, Word and Visio only cells the chart caches). Stacked groups move
      their first or last series to the secondary axis and merge back (`chart/edit/secondary-axis.ts`);
      PowerPoint saves the move through the same operation, so PowerPoint 16 reads the moved series
      on axis group 2 (`chart-neutral-save-series-axis.ts`, 2026-11-16).
- [x] **P1** Chart picture fills are drawn by `src/core/chart/render` (stretched, tiled with
      `a:tile` scale, offset and alignment, stacked along bars for `c:pictureOptions`, with
      transparency) on series, points, the plot area and the chart area, and Word, Excel and Visio
      offer Picture or texture fill, saving pending pictures through `chart/picture-parts.ts`.
      Tiles mirror by `flip`, `stackScale` stacks one picture per `c:pictureStackUnit` and
      `a:srcRect` crops the picture; picture effects other than transparency are not painted.
- [x] **P1** PowerPoint's chart Format pane edits several trendlines per series (type, order,
      period, forecasts, equation and R-squared, Add trendline) and error bars (amount, direction,
      end caps and custom values) in all five bindings; custom error amounts are literal values
      (PowerPoint's model keeps no custom range).
- [ ] **P2** `c:userShapes` (shapes over a chart) is parsed but not drawn by the shared chart
      painter; PowerPoint projects them over the chart instead.
- [x] **P2** Equation ownership audited: all 11 non-test `latex-*` and `omml-*` modules in
      `src/ui/src/pptx/render` are re-export facades over `ooxml-core/math`; none contains duplicate
      conversion logic. The shared math suite and four facade regression files pass (192 and
      170 tests). See `docs/office-shared-parity-delivery.md`.

## Excel

### File format and preservation

- [ ] **P2** Linked pictures (`a:blip r:link` to a web address) are not drawn: Excel downloads
      them on open, the Excel packages never do
      ([powerpoint-external-content.md](powerpoint-external-content.md), observed 2026-11-19).
- [x] **P1** Every save writes `fullCalcOnLoad="1"` and `calcId="191029"`
      (`src/core/xlsx/write/workbook-part.ts`), so Excel recalculates each saved file and asks to
      save on close. Write it only when cached values may be stale.
      Closed: the source calcPr attributes are kept (calcId included), and `fullCalcOnLoad="1"` is
      written only for a formula without a cached result or after an edit that was not
      recalculated (`Workbook.valuesStale`, set by `edit/session-calc.ts`).
- [ ] **P1** `.xlsb` is detected and refused (`load/detect.ts`); `.ods` is not read or written.
- [x] **P1** x14 data validation uses the shared validation model and editor. Cross-sheet lists,
      both formula operands, messages, flags, reference updates, clearing and undo survive save;
      unknown rule markup and collection settings are retained. Evidence: the Excel-authored
      `excel-x14-validation.xlsx`, `xlsx/write/validation-extensions.test.ts`, and all six bindings
      in `e2e/xlsx/x14-validation.spec.ts`. Excel 16 reopened edited, cleared and newly authored rules.
- [x] **P1** Rich values and in-cell pictures (`xl/richData`, Place in Cell, linked data type
      cards, FIELDVALUE and dot access on cached fields, and IMAGE() through an opt-in host
      fetcher) are modelled, drawn and saved; Excel 16 reopens them (`xlsx/rich-values.ts`,
      2026-11-26).
- [ ] **P1** In-cell pictures, value metadata and their media are not shared in collaboration
      (Place in Cell is disabled while sharing), and resolved IMAGE() pictures are not saved.
- [ ] **P1** Connections, query tables and Power Query are carried verbatim only; external links
      use cached values and cannot be refreshed or edited.
- [x] **P1** Threaded comment edits preserve person records, message identities, resolved state,
      mentions and unknown markup. Mention ranges follow edits around their unchanged text;
      copies receive fresh identities. Evidence: `xlsx/write/threaded-comments.test.ts`,
      collaboration regressions and twelve browser cases across all six bindings. Excel 16
      reopened and resaved edited, reopened, duplicated and pasted threads without repair.
- [ ] **P2** Creating new @mentions in threaded comments has no editor or host people-directory
      integration. Imported mentions are preserved; editing their label removes the mention.
- [ ] **P1** Pivots outside the bounded profile are read-only: label and value filters, shared
      caches, calculated fields and items, date-group filters (`pivot/admission.ts`).
- [x] **P2** Phonetic runs (`rPh`) are dropped from shared strings (`read/shared-strings.ts`).
      Closed: `rPh` runs and `phoneticPr` are read onto the cell (`Cell.phonetic`) and written back
      with the shared string. The guide is not displayed, edited or carried by copy and paste or
      collaboration, and `phoneticPr` `fontId` is written as read.
- [x] **P1** Digital signatures are verified on load and shown by state (signed, signer not
      verified, invalid, not verified) in the status bar, a load notice and File > Info, translated;
      a signature that does not verify is never shown as signed, and saving removes signatures as
      Excel does (`src/ui/src/xlsx/signature-status.test.ts`). The editor does not sign workbooks.
- [ ] **P2** Named sheet views and ActiveX controls are not modelled.
- [ ] **P2** Collaboration does not share filters, pivots, panes and views, page setup,
      protection, images and drawings or document properties.

### Calculation

- [x] **P1** Text compares as Excel's does. The comparison operators, criteria (COUNTIF and the
      other IF and IFS functions), XMATCH, XLOOKUP, UNIQUE, SORT and SWITCH compare
      linguistically: case-insensitive, accent-, width- and kana-sensitive, "ß" equal to "ss",
      "æ" and "œ" to their letters, ligatures expanded, final sigma equal to sigma, hyphens and
      apostrophes only breaking ties, symbols before digits before letters. MATCH, VLOOKUP,
      HLOOKUP, wildcards and SEARCH compare each character by its simple case folding
      (`src/core/xlsx/formula/text-compare.ts`). Evidence: `excel-p4-cases.txt`, recorded from
      Excel 16 through hidden COM (`excel-cases.ps1 -Corpus excel-p4`), run by
      `excel-parity.test.ts`. Measured with an English locale only.
- [x] **P1** Exact lookups (MATCH, XMATCH, VLOOKUP, HLOOKUP, XLOOKUP) compare numbers exactly, so
      `MATCH(0.3,...)` skips `0.1+0.2` while `=`, COUNTIF and SWITCH still find them equal; the
      General display of `0.1+0.2` is `0.3`. Evidence: the same corpus.
- [x] **P2** In criteria, `~` is literal unless the criterion holds `*` or `?`; MATCH and SEARCH
      always read it as an escape, and a trailing `~` escapes nothing. A blank criterion cell
      matches zeros, not blanks, and DCOUNTA counts empty text. Evidence: the same corpus.
- [x] **P2** Text limits: `&` and CONCATENATE cut text at 32,767 characters, CONCAT, TEXTJOIN
      and ARRAYTOTEXT return `#CALC!` past it, TEXT refuses formats and results over 255
      characters, BASE pads to at most 255 digits, LOWER keeps capital sharp s. Evidence: the same
      corpus.
- [x] **P0** A small workbook can no longer hang or crash the editor when it opens: each
      formula has a work budget priced from measured cost (matrix products, LAMBDA calls, sorting,
      pattern matching, scans an array of lookups or criteria repeats; about three seconds), a text
      budget and an array budget, past which that formula alone is `#NUM!`
      (`src/core/xlsx/formula/values.ts`); a recalculation stops between formulas, keeping the
      previous values, at an optional time limit or when its results would hold too much
      (`recalc-limits.ts`). The Excel editor reads files on the page, calculates on the page as long as
      reading took, and leaves a slower calculation on open to a worker with a ten-second limit,
      receiving only the cells it wrote
      (`src/ui/src/xlsx/workbook-opener.ts`). Evidence: `hardening.test.ts`,
      `calc-results.test.ts`, `workbook-opener.test.ts`.
- [x] **P1** REGEXTEST, REGEXEXTRACT and REGEXREPLACE run on the engine's own matcher
      (`src/core/xlsx/formula/regex/`), so they accept nested quantifiers such as `(a+)+` as Excel
      does, and a pattern that would run past the formula's work budget is `#NUM!` instead of
      hanging. Syntax is still JavaScript's, not PCRE2's, and scans of a whole 30,000-character text
      are 3 to 10 times slower than the browser's native matcher (under 1.5 ms a call). Evidence:
      `regex/differential.test.ts` (3,000 random patterns against the browser's matcher).
- [ ] **P2** Engine limits Excel does not have: a formula's work budget makes MMULT of two
      arrays larger than about 1,400 by 1,400, and MDETERM or MINVERSE larger than about 840 by
      840, `#NUM!`; a LAMBDA recursing past the JavaScript stack (about 700 levels) is `#NUM!`
      where Excel still allows 1,000 (`excel-p4-recursion-cases.txt`). Only the calculation on
      open has a time limit: a recalculation after an edit of a workbook with many slow formulas
      (each within its budget) still runs to the end on the page.
- [ ] **P2** Data Sort, the AutoFilter value list and Find still compare text with the browser's
      accent-insensitive collation (`src/core/xlsx/edit/sort-keys.ts`, `filter-menu.ts`), not the
      linguistic rule formulas use; Excel's behaviour there has not been measured.
- [x] **P2** `INFO`, `ASC`, `DBCS`, `JIS`, `PHONETIC` and `EUROCONVERT` are implemented locally
      (`formula/functions/info.ts`, `text-east-asian.ts`, `euroconvert.ts`; unit tests in
      `info-east-asian-euro.test.ts`; regenerated `viewers/xlsx/docs/functions.md`). `INFO("directory")`
      returns `""` and `INFO("osversion")` returns `"Web"` without disclosing host paths or OS details.
- [x] **P2** FILTERXML evaluates XPath locally and spills selected values, with Excel's namespace,
      whitespace, direct-text and value conversion behavior. Evidence: 35 recorded Excel cases,
      core edit/undo/save regressions and `e2e/xlsx/filterxml.spec.ts` in all six bindings. Excel 16
      reopened 35-sheet outputs in both date systems without repair and recalculated matching values.
- [x] **P2** Every Insert Function description has French, German, Spanish and Chinese entries
      (`xlsx/ui/commands/i18n/functions/`), enforced over the whole catalogue by
      `xlsx/ui/localization.test.ts`.
- [x] **P2** External formula integrations (`CUBE*` (7), `WEBSERVICE`, `STOCKHISTORY`, `RTD`,
      `IMAGE`, `COPILOT`, `TRANSLATE` and `DETECTLANGUAGE`) are registered with typed host-supplied
      provider interfaces on `CalcEngineOptions` (`formula/providers.ts`, `functions/cube.ts`,
      `functions/integrations.ts`, plus `createMemoryCubeProvider` in `formula/cube-memory.ts`;
      `integrations.test.ts` and `viewers/xlsx/docs/functions.md`). Core makes zero network
      requests: when a host does not supply a provider, cached results survive initial loading and
      newly entered formulas return `#NAME?`. `TRIMRANGE` now also returns a trimmed `RefValue`
      for single-area references.
- [ ] **P2** `EDATE` (and likely `EOMONTH`) over a spill range (`EDATE(A2#, 1)`): the engine lifts
      the call over each element and spills dates, where the samples saw Excel show 0. The cause
      is not established (Excel's handling of ranges in these Analysis ToolPak arguments, or how
      the saved formula is flagged as dynamic); confirm with a file Excel saves before changing
      the lift (`xlsx/formula/functions/datetime.ts`).

### Editing and behaviour

- [x] **P0** AutoFilter drop-down in the grid: Sort, Clear Filter, Filter by Color, Text, Number
      and Date Filters (Custom AutoFilter and Top 10 dialogs, Above/Below Average, date periods),
      search and the value checklist. Value, date-group, custom, Top 10, dynamic and colour
      criteria are read, written and evaluated (`src/core/xlsx/filter-criteria.ts`).
- [x] **P1** AutoFilter remainders: drop-downs on table headers (with undoable table column
      filtering via `setFilterCriteria`), the date checklist as an expandable year/month/day tree,
      icon filters in the menu (`Filter by Cell Icon`, `iconFilterCriteria`), and conditional-format
      fill/font colours in the colour list (`src/core/xlsx/filter-menu.ts`,
      `src/ui/src/xlsx/grid/filter-buttons.ts`, `src/ui/src/xlsx/grid/filter-menu.ts`; evidence:
      `src/ui/src/xlsx/grid/filter-menu-remainders.test.ts`).
- [ ] **P1** Sort on cell colour, font colour or icon, custom lists, case sensitivity and
      left-to-right sorting, a table's own saved sort, Data > Reapply and File > Options > Advanced > Edit Custom Lists (add, edit/rename, delete, and import from cells;
      `src/ui/src/xlsx/backstage/options-custom-lists.ts`, `options-custom-lists.test.ts`) are done
      (`src/ui/src/xlsx/dialogs/sort-dialog.ts`, `src/core/xlsx/edit/sort*.ts`). Still open: user
      custom lists are kept in the browser's localStorage (session only when it is blocked), not
      in a roaming profile.
- [x] **P1** Formula auditing is implemented (`src/core/xlsx/audit/`: Trace Precedents/Dependents through
      names, tables, 3D references, `OFFSET` and `INDIRECT` targets with repeated levels, off-sheet
      markers and red error arrows, cleared by structural edits; Remove Arrows; Evaluate Formula
      with Step In/Out and step-by-step chosen-branch evaluation for `IF`, `IFERROR`, `IFNA`,
      `CHOOSE`, `IFS` and `SWITCH`; Error Checking with all nine rules including two-digit text
      years and list data validation, Ignore Error saved as `ignoredErrors` and the rules listed in
      File > Options > Formulas; a live Watch Window; `audit-parity.test.ts`). Background
      green-triangle indicators on the grid react to enabled rules and `ignoredErrors`
      (`src/ui/src/xlsx/grid/grid-view.ts`), and Options > Formulas rule toggles persist between
      sessions in localStorage (`src/ui/src/xlsx/error-rules.ts`, `error-rules.test.ts`).
- [x] **P1** Active cells, selections and the active tab are saved after navigation. Split panes
      retain separate selections; frozen panes share the sheet selection as Excel does. Disjoint
      ranges write `activeCellId`. Core regression: `xlsx/ui/selection-persistence.test.ts`;
      all six bindings: `e2e/xlsx/selection-save.spec.ts`. Excel 16 reopened three generated
      workbooks with the expected active tab, disjoint ranges, frozen selection and split pane.
- [x] **P1** Flash Fill (`Ctrl+E`), Subtotal, Consolidate and Advanced filter are implemented in
      core with undo/redo (`xlsx/edit/flash-fill.ts`, `subtotal.ts`, `consolidate.ts`,
      `advanced-filter.ts`) and wired on the Data ribbon tab with dialogs and all four locales
      (`ui/commands/data.ts`, `src/ui/src/xlsx/dialogs/{subtotal,consolidate,advanced-filter}.ts`;
      verified in `flash-fill.test.ts`, `subtotal.test.ts`, `consolidate.test.ts`,
      `advanced-filter.test.ts`, `data.test.ts` and `data-tools.test.ts`).
- [x] **P1** Split panes (only Freeze Panes exists).
      Closed: View > Split splits at the active cell (a cross in the middle for A1), giving two or
      four independently scrolling panes with draggable bars and double-click removal; Split and
      Freeze Panes replace each other, keeping the window where it was. `pane state="split"`
      (twentieths of a point, `activePane`, each pane's `selection`) is modelled as
      `SheetView.split`, and both panes' first cells are kept current as the user scrolls and
      written as `pane/@topLeftCell` and `sheetView/@topLeftCell`. Maths:
      `src/core/xlsx/layout/split.ts`; painting: `grid/split-bars.ts` and the quadrant painter;
      browser check: `e2e/xlsx/split-panes.spec.ts`.
- [x] **P1** Text to Columns refuses fixed-width splitting (`dialogs/text-to-columns.ts`).
      Closed: the dialog offers Fixed width with typed break positions and a preview; the split
      rules are `splitFixedWidth` in `src/core/xlsx/edit/text-columns.ts`. Column-data-format
      conversion per column (Excel's Column data format step) is not offered.
- [x] **P1** Changing the workbook theme: Page Layout > Themes, Colors and Fonts apply built-in
      Office presets and colour/font schemes undoably (`xlsx/theme-presets.ts`, `edit/theme.ts`,
      `ui/commands/page-layout.ts`, `write/theme.ts`; verified in `edit/theme.test.ts`,
      `write/theme.test.ts` and `ui/commands/page-layout.test.ts`). Theme Effects presets remain
      open.
- [ ] **P1** Form controls cannot be moved or resized and buttons run no macros; shapes have no
      effects, 3-D, pattern or picture fills, and gradients are linear only.
- [ ] **P1** Printing goes through the browser; Page Layout view has no copy and paste, per-view
      zoom or native printer metrics.
- [x] **P2** Review > New Comment (Shift+F2) edits an existing thread's root and replies in one
      undo step, retaining authors and creation timestamps. Core round-trip and collaboration
      regressions cover reply edits; `e2e/xlsx/comment-replies.spec.ts` checks undo, redo and
      save/reopen in all six bindings. Excel 16 reopened the edited native fixture without repair,
      with the expected reply text, author and root/reply timestamps. Metadata limitations are above.
- [x] **P2** Fill-handle drag autofill exists in `grid/grid-pointer.ts`. The browser regression
      `e2e/xlsx/fill-handle.spec.ts` drags a two-cell numeric series to five cells and checks one-step
      undo and redo in all six bindings.
- [x] **P2** The grid exposes the shown window to assistive technology as rows of column headers,
      row headers and grid cells whose row and column indices are the sheet's own, as Excel reports
      them (B2 is row 2, column 2; the headers take no index) (`grid/grid-a11y.ts`, the window and the
      cell description in `src/core/xlsx/layout/accessibility.ts`); the active cell is the focus
      sink's active descendant, and its label names the table column, formula and comment and
      follows edits made without moving the selection. Pictures are named by their alt text.
      `grid/grid-a11y.test.ts`, `e2e/xlsx/grid-a11y.spec.ts` and `e2e/a11y/tree.spec.ts` cover it,
      and the axe baseline no longer allows missing rows, row parents or unnamed images in Excel.
      Excel's own screen-reader output was not compared.

### Ribbon

- [ ] **P1** Missing tabs: Chart Format (a pane instead).
- [x] **P2** Draw tab: drawing, the eraser, the lasso, Ink to Shape and Ink Replay work and save
      as Excel ink (`xlsx/edit/ink.test.ts`, `e2e/xlsx/draw-tab.spec.ts`). The Ruler (Stencils)
      lies over the cells; it is dragged, turned with the mouse wheel or Shift and the arrow keys,
      and a stroke begun at its edge is drawn straight along it (`ink/ink-ruler.test.ts`); it does
      not measure lengths or show a protractor. Ink to Math opens the Ink Equation
      dialog with the host's recognizer (`setInkMathRecognizer`, none built in; disabled with its
      reason without one) and inserts an equation text box (`xlsx/commands/ink-math.test.ts`).
      Ink is shared in a shared workbook, last writer wins per drawing
      (`xlsx/collab/share-ink.test.ts`).
- [x] **P1** PivotTable Analyze and PivotTable Design contextual tabs: options, field settings,
      refresh, change source, subtotals, grand totals, report layout, blank rows, the PivotTable
      style gallery and style options; Excel 16 reopens and refreshes the results (2026-11-26).
- [ ] **P1** PivotTable definitions and caches are not mapped to Yjs, so pivot editing is disabled
      while a collaboration session is bound; the Field List is the settings dialog, not a pane.
- [ ] **P2** PivotTable value field number formats (`dataField/@numFmtId`) are read, written and
      set through the model API (`valueNumberFormat`, `numberFormat`), but the report's value
      cells do not take the format when the model lays them out, and Value Field Settings has no
      Number Format button (`xlsx/pivot/number-format.test.ts`).
- [x] **P1** Insert: Shapes (the shared preset gallery), Text Box, New Note (a legacy note) and
      Header & Footer (edited in Page Layout view), with undo and save round trips
      (`xlsx/edit/page-authoring.test.ts`, `src/ui/src/xlsx/commands/ribbon-authoring.test.ts`,
      `e2e/xlsx/ribbon-page-authoring.spec.ts`).
- [x] **P1** Insert: PivotChart (from a range, table or PivotTable), Recommended PivotTables
      (sampled previews) and Equation (an Office Math text box); desktop Excel 16 reopens and
      refreshes the PivotCharts and re-saves the equations without repair.
- [x] **P1** Insert > 3D Models embeds a self-contained GLB 2.0 with a PNG preview. Insert > Object
      embeds an existing DOCX, XLSX or PPTX as a worksheet icon. The original bytes survive
      save and reopen, and the embedded Office file can be downloaded from its icon.
- [ ] **P1** Insert still needs stock models, other model formats, interactive 3D controls,
      desktop OLE activation, Create New and link-to-file.
- [x] **P1** Insert places Add-ins after Illustrations and fits at 1600px and 1280px; PivotTable
      and Sparkline editing live on their contextual tabs. Review > Notes > New Note creates legacy
      notes, and the header Comments button toggles the threaded comments pane with New Comment
      and focus return on Escape, in six bindings (`e2e/xlsx/ribbon-insert-order.spec.ts`,
      `e2e/xlsx/comments-list.spec.ts`, 2026-11-27).
- [x] **P1** WordArt inserts and edits native worksheet shapes with fill, outline and shadow
      presets and DrawingML text transforms.
- [x] **P1** Insert Symbol replaces the editing selection or inserts at the cell's text caret.
- [x] **P1** Wide folded ribbon groups scroll within the phone screen. Save, Undo and Redo
      stay available in the compact Quick Access Toolbar.
- [x] **P1** Page Layout: Scale to Fit (width, height, scale), Print Titles, a sheet background
      (drawn behind the cells, never printed) and Arrange (order, Selection Pane, align, group of
      shapes, rotate, flip, hide) (`xlsx/edit/arrange-drawings.test.ts`,
      `src/ui/src/xlsx/commands/arrange.test.ts`).
- [ ] **P1** Page Layout > Arrange: grouping pictures, mixed drawings and nested groups, and
      rotation, flip and Selection Pane visibility for drawings other than shapes and pictures.
- [ ] **P1** Data: Get Data and its sources, Queries & Connections, Refresh All, Edit Links.
- [ ] **P2** Review: Spelling, Thesaurus, Translate, Check Accessibility, Notes; View: Window group
      and Macros; Help: Help, Training, Feedback; Backstage: Account, Close.

## Word

### File format and preservation

- [ ] **P2** Linked pictures (`a:blip r:link` to a web address) are not drawn: Word downloads
      them on open, the Word packages never do
      ([powerpoint-external-content.md](powerpoint-external-content.md), observed 2026-11-19).
- [x] **P0** Legacy `.doc` opens as main-body text only: tables, headers, footers, images, fields
      and formatting are dropped, and saving writes plain paragraphs. Closed: `readDocModel`
      (`@christophervr/ole2`) is converted to a WordprocessingML package that the normal reader
      opens (`src/core/docx/load/legacy-doc.ts`, `legacy-doc/`): styles, numbering, run and
      paragraph formatting, tables (spans, merges, borders, shading, nested), sections, headers and
      footers by variant, footnotes, endnotes, comments, pictures (inline and floating), text
      boxes, hyperlinks, fields and bookmarks. The result saves as a full `.docx`.
- [ ] **P1** Legacy `.doc` writes back only paragraph text edits: `save()` keeps the binary format
      (unchanged bytes, or text edits to plain paragraphs) and refuses anything else; the editor
      saves `.docx` through `saveDocx()` (the file is renamed on open and the load warnings say so).
      No full `.doc` writer. Not converted: revision marks,
      form-field data, the commented text range of a comment (it is anchored at its mark), drawing
      shapes other than pictures and text boxes (lines, rectangles, groups), embedded OLE data
      (the preview picture is kept), WordArt, line numbering, footnote numbering settings, macros,
      document protection, table-style definitions, hyperlink and field flags, `PICT` and
      Symbol-font character runs.
- [x] **P1** Styles: Create a Style and Modify Style for paragraph, character, linked and table
      styles with undo; edited style catalogues, latent styles and the theme are saved with
      unknown markup kept (`style-catalog-save.test.ts`, `theme-save.test.ts`,
      `src/ui/src/docx/style-commands.test.ts`). Desktop Word 16 opens a saved file without repair
      and lists all four new styles with their base styles and the edited theme colour and font.
- [x] **P2** A table style's paragraph formatting (`w:pPr`, including conditional regions) is
      applied in the editor, Print Layout and PDF between the document defaults and the paragraph
      style, with both halves of `overrideTableStyleFontSizeAndJustification` ([MS-DOCX]): the
      default paragraph style's left justification and its 11 or 12 pt size (`w:sz` and
      `w:szCs` independently) do not override the table style's (#353,
      `docx/paragraph-styles.ts`, `docx/run-formatting.ts`,
      `layout/table-style-font-size.test.ts`).
- [x] **P1** Table-level properties (width, alignment, indent, borders, style, look, grid widths)
      are editable in core with both fresh and source-preserving save/reload round-trip support
      (`docx/table-properties-edit.ts`, `write-table-level-properties.ts`; verified in
      `table-properties-edit.test.ts` and `table-fidelity.test.ts`).
- [ ] **P1** Table-structure revisions (`cellIns`, `cellDel`, `cellMerge`, `tblPrChange`,
      `trPrChange`, `tcPrChange`) are not modelled and block structure edits; `tblGridChange` and
      `numberingChange` have no code. Tracked row insertions and deletions (`w:trPr/w:ins`,
      `w:del`) are modelled, accepted and rejected (`docx/table-row-revisions.test.ts`) but not yet
      exercised through the editor.
- [ ] **P1** Fields computed today: PAGE, NUMPAGES, SECTIONPAGES, DATE, TIME, TOC, SEQ, REF,
      PAGEREF, STYLEREF, DOCPROPERTY, AUTHOR, TITLE, SUBJECT, KEYWORDS, COMMENTS, LASTSAVEDBY,
      REVNUM, CREATEDATE, SAVEDATE, PRINTDATE, IF, `=` (scalar/list arithmetic & functions; table
      and bookmark references keep their cached result), AUTONUM, AUTONUMLGL, AUTONUMOUT, LISTNUM
      and SYMBOL (`docx/computed-fields.ts`, `computed-fields-extra.test.ts`). MERGEFIELD,
      INCLUDETEXT, INCLUDEPICTURE, INDEX/XE, TA/TOA and CITATION/BIBLIOGRAPHY keep their cached
      result (`layout/page-fields.ts`); Update Fields skips the `\n \r \w \p` switches.
- [ ] **P2** A HYPERLINK written as a complex field (`w:fldChar` and `w:instrText`, as Word writes
      table of contents entries and cross-reference links) is kept on save but is not followed as
      a link, and its result edits as plain text (`docx/parse-warnings.ts`; the sample
      `samples/docx/services-agreement-form.docx`).
- [ ] **P2** Legacy form fields (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN) are filled, toggled and
      picked, but their entry and exit macros, calculated results and number or date formats are
      not applied (`docx/parse-warnings.ts`).
- [ ] **P2** DATE and TIME results name months and weekdays in the result run's `w:lang/@w:val`
      (through styles and document defaults; a month after a day number takes the contextual
      form, measured for Russian in Word 16), but AM/PM designators stay English, the `w:bidi`
      and `w:eastAsia` languages are not consulted, and the `\@` switch of other date fields
      (CREATEDATE, SAVEDATE, PRINTDATE, merge fields) still gives English names
      (`docx/field-date.ts`, `docx/field-format.ts`).
- [ ] **P2** A list number takes the paragraph mark's run properties only when its paragraph has
      no visible text; a numbered paragraph with text still formats its number from its first run
      (`docx/layout/adapter.ts`, `src/ui/src/docx/list-marker-display.ts`).
- [ ] **P1** Canvases (`wpc`), legacy VML text boxes (`v:textbox`)
      and legacy VML WordArt and gradient or picture fills are kept but not drawn.
      DrawingML ink (`w14:contentPart`) is drawn and edited on the Draw tab (`docx/ink.test.ts`);
      legacy VML ink (`o:ink`) is drawn read only from its shape path (its ISF strokes are not
      decoded), cannot be edited on the Draw tab, and is saved as read, moved or resized through
      its style when floating (`docx/vml-ink.test.ts`).
- [ ] **P1** Ruby (`w:ruby`) has no model and probably blocks editing its paragraph; vertical
      text (`tbRl`, `btLr`), `eastAsianLayout`, `fitText` and emphasis marks are not laid out.
- [x] **P1** Picture bullets (`numPicBullet`): resolved (VML `imagedata` and `w:drawing`), drawn as
      the editor marker at the font size or the authored shape size, measured by the picture width
      in page layout, kept byte-identical on save, and creatable through Bullets > Picture bullet...
      (new entry, media part and `numbering.xml.rels` relationship written on save).
- [x] **P1** Protection: forms, read-only, comments-only and tracked-changes-only protection are
      enforced, and Review > Restrict Editing starts and stops them (password checked with the
      ECMA-376 hash, written back on save).
- [ ] **P2** Restrict Editing leftovers: only "Everyone" exceptions (no per-user list); an
      exception that continues across paragraphs or sits between blocks, and column-range
      permissions (`w:colFirst`), are enforced but cannot be edited in the pane; the allowed-styles
      list of formatting restrictions is kept but not enforced; the Word 97 four-byte `w:hash`
      (no `cryptAlgorithmSid`) cannot be verified, so that protection cannot be stopped here.
      Word-written passwords (legacy key text, SHA-1/256/384/512 by sid) are checked against
      fixtures saved by Word 16 (SHA-512 only; the other sids are checked against an independent
      implementation).
- [x] **P1** Digital signatures: Word verifies them on load and shows the verified state (signed,
      signed with the signer not verified, invalid, not verified) in a status bar badge, a load
      notice and File > Info with each signer (`office-ui-signature-status`,
      `src/ui/src/docx/signature-status.test.ts`, `e2e/docx/signature-status.spec.ts`). Only a
      signature that verifies is shown as signed, and the notice says saving edits removes them.
      Unchanged saves keep the original bytes; edited saves remove signature origins, signatures,
      certificates, relationships and content types through `opc/signature`, including
      nonstandard part locations (`src/core/docx/signatures.test.ts`). A browser has no trust
      store, so a valid signature reads "signer not verified" there. The editor does not sign.
- [ ] **P1** ActiveX controls and OLE objects without a supported preview display as placeholders.
      OLE and ActiveX activation in the browser remains open.
- [x] **P2** Comment metadata edits retain source bodies; plain-text edits patch existing runs;
      paragraph and durable IDs, reply anchors, UTC dates and metadata parts stay aligned.
      Evidence: `docx/comment-preservation.test.ts`, `docx/comments.test.ts`, six-binding
      `e2e/docx/comments-pane.spec.ts`, and Word 16 reopen/save checks of five edited variants
      of `docx/__fixtures__/word-comments-native.docx` (2026-10-31).
- [x] **P2** Comment text includes significant spaces, tabs, soft breaks and special hyphens.
      Edits retain surviving inline characters and write new tabs as Word elements, with native
      Word reopen/save evidence (`docx/comment-text.test.ts`, `word-comment-whitespace-native.docx`)
      and display/save coverage in all six bindings (`e2e/docx/comments-pane.spec.ts`).
- [ ] **P2** The comments pane cannot edit existing comment text. Rich comment content
      (formatting, pictures and tables) is retained in source parts, but not rendered or authored.
- [ ] **P2** Embedded fonts are kept, not used for rendering; `altChunk` and art page borders
      are kept, not drawn.
- [x] **P2** Solid theme page backgrounds resolve their theme colour, logical mapping and
      tint/shade, taking precedence over stale cached RGB. Text edits preserve the theme reference;
      an explicit page colour replaces it. Evidence: `docx/page-background.test.ts`, the native
      Word 16 corpus in `docx/__fixtures__/word-theme-colors.json`, and six browser bindings in
      `e2e/docx/page-color.spec.ts` (2026-10-31).
- [ ] **P2** Word theme tint/shade quantization still differs from native cached colours by up
      to one RGB channel level in the 90-colour corpus (`docx/theme-native.test.ts`). HSL luminance
      and tint-over-shade precedence are implemented; exact native rounding remains.
- [ ] **P2** Non-solid page backgrounds (VML gradient, pattern and picture fills) are preserved
      but not painted (`docx/page-background.ts`).

### Editing and behaviour

- [ ] **P1** Mail merge: the headless engine (`ooxml-core/docx/mail-merge`,
      `docs/docx-mail-merge.md`) reads CSV, TSV, `.xlsx` and host-supplied recipients, filters and
      sorts them, evaluates MERGEFIELD, NEXT, NEXTIF, SKIPIF, MERGEREC, MERGESEQ, ASK, FILLIN and
      IF, previews a record, checks for errors and writes one document or one per record; desktop
      Word 16 opens a merged result without repair. Fields that span paragraphs and fields in
      DrawingML text boxes merge (`mail-merge/cross-paragraph.ts`), and local CSV and TSV
      recipients are read in slices up to 100,000 records, 255 fields and 32 MiB (2026-11-27).
      Open: legacy VML-only text boxes, fields that cross a table, recipient lists beyond the
      in-memory caps (the records are still held in memory once read), and locale-specific
      number and date results checked against Word.
- [x] **P1** Find & Replace supports wildcards, whole words, prefix and suffix, ignore options,
      Special codes (`src/core/docx/find-pattern.ts`, `src/ui/src/docx/search-panel.ts`), and
      formatting search & replacement (`FindFormatCriteria`, `ReplaceFormatPatch` for font,
      paragraph and style, including format-only search over maximal contiguous spans;
      `docx/find-format*.ts`, `find-format.test.ts`, plus search-panel UI pickers for Format >
      Font, Paragraph and Style in `src/ui/src/docx/search-format-controls.ts`,
      `search-format-match.ts`, `search-format.test.ts`).
- [ ] **P2** Find and Replace search the main document, headers and footers, footnotes, endnotes and
      text boxes in turn (`docx/find-stories.ts`, `src/ui/src/docx/search-stories.test.ts`,
      `e2e/docx/search.spec.ts`); a match in a header, footer or note opens its editor (Print Layout
      switches to Draft). Open: comments, text boxes inside headers, footers and notes, the Find In
      story picker, navigation pane results, the host Find dialog (body only), Replace All being one
      undo step per story instead of one, and the story order is unconfirmed against Word's dialog.
- [x] **P1** Read Mode (columns and paging), Web Layout, Outline (levels, collapse, promote and
      demote with heading styles, branch moves) and Focus, in every binding (`docx/view-modes.ts`,
      2026-11-26).
- [x] **P1** Read Mode shows a footnote or endnote in a callout from its reference (click, Enter
      or Space; Escape or Back to text returns to the reference), and Outline's promote, demote,
      Outline Level and branch moves work with Track Changes on (paragraph property changes and
      linked `w:moveFrom`/`w:moveTo` text and paragraph marks) and in collaboration
      (`view-read-notes.ts`, `view-outline-rules.ts`, `docx/ui/outline-tracked.ts`, 2026-11-27).
- [ ] **P2** Outline stays honest where it cannot follow Word: a tracked branch move is refused
      (with its reason) when the branch holds tables, pictures, fields or breaks, already carries
      tracked changes, or would leave or land without a following paragraph; in collaboration a
      level change that must define a new heading style is refused, since style catalogs are not
      shared with the room.
- [x] **P1** Index (Mark Entry, Insert and Update Index), citations and bibliography (Manage
      Sources in `b:Sources`, APA, MLA and Chicago), and Table of Authorities (2026-11-27).
- [ ] **P2** The Index's section layout, the native Index and TOA styles and tab stops, and the
      Bibliography gallery's nested content control differ from Word.
- [ ] **P1** Compare is core only (`docx/compare/`, `docs/docx-compare.md`): body text and tables
      become tracked changes at word or character level, matching Word 16's CompareDocuments on 8
      fixture pairs. Not compared: moves, headers and footers, notes, comments, text boxes,
      paragraph, table and section formatting; tables with column changes, merged or nested cells
      show as replaced. Combine Documents, the Compare/Combine UI and the accessibility checker
      remain open.
- [ ] **P1** Paste: Word's clipboard HTML pastes as real lists with Word's levels and formats, its
      paragraph styles where the document defines them, and Word-only run formatting
      (`docx/clipboard/`); a Paste Options button (or Ctrl) re-pastes as Keep Source Formatting,
      Merge Formatting or Keep Text Only, and Paste Special (Ctrl+Alt+V) offers HTML or unformatted
      text. Open: notes and fields paste as text, missing source styles are not created, the header,
      footer and note editors do not read Word HTML, no K/M/T keys or other Paste Special formats,
      Merge Formatting is approximate, and pasted Word pictures do not load.
- [ ] **P1** Pictures: crop by percent in Format Picture and all seven Wrap Text choices (Picture
      Format tab and right-click menu, inline and floating converted both ways), plus dropping or
      pasting picture files at the pointer (`docx/picture-crop-edit.ts`, `docx/drawing-wrap.ts`,
      `e2e/docx/picture-drop-wrap.spec.ts`). Open: crop handles and crop by length, wrap polygon
      editing (Tight and Through use a rectangle), text distances, wrap and crop for header, footer,
      note and VML pictures (absolute VML pictures float and save size, position, wrap and
      stacking edits since #345; crop, inline conversion, rotation, flips and alt text are disabled
      for them with a reason), and the right-click menu on Print Layout pages.
- [x] **P2** Footnotes and endnotes convert like Word's Convert Notes (all to endnotes, all to
      footnotes, swap) and Convert to Endnote/Footnote on a note reference, swapping the note styles
      when defined (`docx/note-convert.ts`, `e2e/docx/convert-notes.spec.ts`); picture watermarks
      with scale and washout (`docx/picture-watermark.ts`, `e2e/docx/picture-watermark.spec.ts`).
- [ ] **P2** Note conversion and watermark leftovers: converting from inside the note text, notes
      holding charts, shapes or floating pictures (left in place), and drawing picture watermarks in
      Draft view.
- [ ] **P2** Inspect Document (File > Info > Check for Issues) inspects and undoably removes
      comments and revisions, document properties and personal information (setting and honouring
      `w:removePersonalInformation` and `w:removeDateAndTime`), custom XML data, header and footer
      content and hidden text, matching Word 16's saved output
      (`docx/remove-document-information.test.ts`, `e2e/docx/document-inspector.spec.ts`). Open:
      Collapsed Headings and Invisible Content, hidden paragraph marks and style-inherited hidden
      text, emptied comment parts are kept where Word deletes them, the glossary is not scrubbed,
      and the privacy flag cannot be turned off.
- [ ] **P2** Content controls: picture, building-block gallery and repeating-section controls are
      kept but cannot be inserted, and repeating items cannot be added or removed
      (`docx/content-control-model.ts`). The writer writes `w15:repeatingSection` and
      `w15:repeatingSectionItem` for a block or row control the model gives that kind, but the
      model holds one control per block or row, so a section wrapping its items (nested `w:sdt`)
      cannot be built from the model.
- [x] **P2** Content controls shown as tags (`w15:appearance="tags"`) draw start and end tags for
      inline, paragraph, row and cell controls in the editor, Print Layout and PDF, with the
      advances, tagged line height and outlines measured from Word 16 through COM and its PDF
      export (`docx/layout/fixtures/word-layout-reports/README.md`). The editing caret sits inside
      the control after the start tag and before the end tag and keeps that side across a remote
      collaborator's edit (#349, `e2e/docx/content-control-tag-caret.spec.ts`). Not measured:
      tags under multiple line spacing and in right-to-left paragraphs, whose Print Layout and PDF
      follow the open right-to-left reordering gap.
- [ ] **P2** A table cell whose content is wrapped in tagged cell and row content controls is laid
      out narrower in Print Layout than in Word, so a long word breaks inside the cell and the end
      tags are clipped at the cell edge (`block-content-control-tags.docx` against Word's PDF).
- [x] **P2** Hyperlinks to http, https, mailto, tel and ftp targets open on Ctrl+click, as in Word;
      script and disguised schemes are refused (`docx/hyperlink.test.ts`).
- [x] **P1** AutoCorrect and AutoFormat as you type: smart quotes, automatic bulleted and numbered
      lists (`- `, `* `, `1. `, `1) `), em/en dashes, ellipsis, common symbols (`(c)`, `(r)`,
      `(tm)`, arrows), fractions (`1/2`, `1/4`, `3/4`) and common English typos, with single-step
      Undo back to the raw typed characters (`src/core/docx/autocorrect.ts`,
      `src/ui/src/docx/autocorrect-plugin.ts`; evidence: `autocorrect.test.ts`,
      `autocorrect-plugin.test.ts`). User-editable AutoCorrect dictionary UI remains open.
- [ ] **P1** Collaboration covers the body and comments; headers, footers and notes are
      incomplete, table commands are off during collaboration, and there are no mentions.
- [ ] **P2** Customize Ribbon can show and hide, not reorder; Options are not persisted.

### Layout and rendering

- [x] **P1** Hyphenation: `autoHyphenation`, `hyphenationZone`, `consecutiveHyphenLimit`,
      `doNotHyphenateCaps`, paragraph/style `suppressAutoHyphens` and explicit soft hyphens
      (`\u00AD`) are applied by `src/core/docx/layout` using the offline English/Latin hyphenator
      (`text/hyphenation.ts`, `docx/layout/hyphenate-line.ts`; verified in
      `text/hyphenation.test.ts`, `docx/hyphenation.test.ts`, `docx/layout/hyphenation.test.ts`).
      Non-English hyphenation dictionaries remain open.
- [ ] **P1** Compatibility mode: `docx/compatibility.ts` reads `w:compat` into one profile; mode 15
      versus 14 and earlier changes page-break spacing, footnote separators and columns, and the
      width AutoFit tables size against (measured in modes 12, 14 and no mode). Legacy tables are
      not outdented by the cell margin, and the other mode-dependent rules and most `w:compat`
      options (including `growAutofit` and `autofitToFirstFixedWidthCell`) are not applied.
- [ ] **P1** Floating tables (`w:tblpPr`) are positioned by their anchors, offsets and alignment and
      text wraps on both sides (`layout/flow-floating-table.ts`, checked against Word 16 in
      `__fixtures__/floating-table`). They stay on the page where they occur; tables that do not
      fit, hold footnotes or merge cells vertically fall back to the flow; narrow gaps stay empty;
      floating tables in cells, on multi-column pages and in modes before 15 are unmeasured.
- [x] **P1** Text frames (`w:framePr`) other than drop caps are parsed with style inheritance and
      laid out at their anchored box positions (`hAnchor`/`vAnchor`, `x`/`y`, `xAlign`/`yAlign`,
      `w`/`h`/`hRule`), grouping consecutive matching paragraphs into one frame and registering
      wrap exclusions (`around`, `notBeside`, `none`) (`docx/layout/flow-text-frame.ts`,
      `docx/layout/text-frame.test.ts`).
- [ ] **P2** East Asian layout: kinsoku (per-language default sets measured from Word 16,
      `w:strictFirstAndLastChars`, `w:noLineBreaksBefore`/`After`), `w:autoSpaceDE`/`DN` (a quarter
      em) and hanging punctuation in justified paragraphs (`w:overflowPunct`) are applied
      (`text/kinsoku.ts`, `text/east-asian.ts`, `layout/text-breaks.ts`). Open: punctuation
      compression (`w:characterSpacingControl`), `w:wordWrap`, justified CJK lines spread between
      characters, the font-dependent gap after Western text, and checks on a Word with East Asian
      editing languages enabled.
- [x] **P1** Thai, Lao, Khmer and Myanmar text breaks between grapheme clusters as Word 16 does
      (never inside a cluster; U+200B always breaks) (`text/complex-breaks.ts`). Measured on an
      install without a Thai editing language; dictionary word breaking is not modelled.
- [x] **P1** Fields LISTNUM, AUTONUM, AUTONUMLGL, AUTONUMOUT, DOCPROPERTY and SYMBOL are computed in
      layout, by Update Fields and on save, matching Word 16's output
      (`docx/__fixtures__/computed-fields`).
- [ ] **P2** Computed-field leftovers: DOCPROPERTY `\#` pictures, dates without `\@` and statistics
      keep their saved result; AUTONUM and LISTNUM count the main document only; LISTNUM resolves
      only NumberDefault, OutlineDefault and LegalDefault by name; SYMBOL `\h` is not applied;
      DOCPROPERTY shows the current value where Word keeps the stale result until F9.
- [ ] **P1** Right-to-left text: paragraph direction is applied but runs are not reordered by
      the bidirectional algorithm, and the Word UI has no right-to-left locale or mirrored chrome.
- [ ] **P1** Tight and through wrapping use the object's rectangle, not its wrap polygon, and only
      the larger side gets text (`layout/floats.ts`).
- [ ] **P1** Table AutoFit (`layout/table-autofit.ts`) sizes columns from content, preferred cell
      and table widths and merged cells, matching Word 16's saved grids in
      `__fixtures__/table-autofit`; default cell margins resolve through the table, its style chain
      and the default table style. `gridBefore`/`gridAfter`, `noWrap`, cell spacing, negative
      indents and preferred widths on merged cells are not measured.
- [ ] **P1** Merged cells are read-only on Print Layout pages.
- [x] **P1** Floating objects anchored in table cells are laid out in their cells on Print Layout
      pages as Word 16 does: from compatibility mode 15 always (`layoutInCell` is ignored), before
      it only with `layoutInCell="1"` (`"0"` places the object on the page). The cell gives every
      horizontal and vertical frame, the cell's text wraps around a wrapping object and its row
      grows to hold it; floating charts are positioned, wrapped and painted like inline charts, in
      the body and in cells; `character` and `line` frames are measured from the anchor
      (`docx/__fixtures__/cell-floats`, `layout/cell-floats.ts`). Not modelled: before mode 15 an
      object placed on the page does not push its table down, and such tables are not outdented
      by their cell margin as Word 2010 draws them.
- [x] **P1** Linked text boxes (`wps:txbx/@id` and `wps:linkedTxbx`) flow on Print Layout pages:
      text one box cannot fit continues in the next box of the chain, in `seq` order, reflowed at
      that box's width; the last box keeps the rest clipped, as Word 16 draws it. Saved byte for byte
      (`docx/layout/text-box-layout.test.ts`).
- [x] **P2** A wrapping floating object positioned from its paragraph or line moves its paragraph
      to the next page when it would cross the bottom margin in compatibility mode 15, and only
      when it would pass the paper edge in earlier modes; a `wrapNone` object (in front of or
      behind the text) stays with its anchor and may overlap the margin and the paper edge, as
      measured in Word 16 (#354, `layout/floats.ts` `floatExceedsBottomLimit`,
      `layout/floats.test.ts`, `e2e/docx/print-cell-floats.spec.ts`).
- [ ] **P2** Before mode 15 Word keeps a square-wrapped paragraph-relative object that passes the
      paper edge by up to about 24 pt on its page, rewrapping the anchor's text; the engine moves
      it to the next page as soon as it passes the edge (#354).
- [ ] **P2** A square-wrapped object positioned from its line is placed from the anchor's line
      after the text wraps around it, so a long anchor paragraph places it lower than Word, which
      uses the line before wrapping, and can move it to the next page early (#354).
- [x] **P2** `w:compat/w:doNotExpandShiftReturn` leaves a justified line that ends in a
      text-wrapping break (`w:br`) at its natural width, as the last line is, except in
      `jc="distribute"` paragraphs, which still expand it; a line ending in a page or column break
      is never expanded, with or without the setting, as measured in Word 16 (#358,
      `layout/paragraph-layout.test.ts`, `layout/word-layout-reports.test.ts`).
- [x] **P1** Text boxes are drawn on pages: a shape holding a `wps:txbx` (a `rect` text box, a line
      or connector) and a text box written with `txBox="1"` are drawn with their frame and their
      paragraphs, floating or inline, in the body and in headers and footers, and `txBox="1"`
      boxes keep their Text Box editing (#304, `docx/text-box-content.test.ts`,
      `docx/text-box-frame.test.ts`, `src/ui/src/docx/story-drawings.test.ts`).
- [x] **P2** Shape text takes the colour of its `wps:style` font reference (`a:fontRef`, resolved
      through the theme) where its runs set none, in SVG text and over text box paragraphs alike;
      without one, white on dark fills and black on light ones (`docx/text-box-frame.test.ts`).
- [x] **P2** Text box paragraphs on Print Layout pages are laid out by the page engine
      (`layout/text-box-layout.ts`), in the body and in headers and footers: lines break by the
      body's paragraph flow inside the preset's text rectangle less the `wps:bodyPr` insets and half
      the outline (measured against Word 16's PDF), overflow is clipped at that rectangle,
      `vert`/`vert270` turn the text, `spAutoFit` sizes the shape to its text and `wrap="none"` to its
      longest line for display (the saved extent is unchanged), `normAutofit` is drawn as no autofit
      and `numCol` in one column, both as Word draws them (`docx/layout/text-box-layout.test.ts`,
      `src/ui/src/docx/print-text-box.test.ts`).
- [ ] **P2** Text boxes on the editing surface and in Draft view's previews still flow by CSS: their
      line breaks can differ from Print Layout's, and autofit and linked boxes apply in Print Layout
      only. `eaVert`, `mongolianVert` and the WordArt vertical kinds are drawn as `vert`, without
      upright East Asian characters or stacked letters. Formatted text inside a shape (bold, sizes, borders, alignment other than
      centred) is read-only, so editing the shape keeps it.
- [ ] **P2** On the editing surface a floating picture or text box is floated left or right, or set
      on a line of its own: its exact position, its offsets and its place in front of or behind the
      text are shown on Print Layout pages only, so a cover picture behind the text or a caption
      box over a picture is drawn in the flow (`docx/parse-warnings.ts`; the samples
      `samples/docx/technical-report.docx` and `community-newsletter.docx`).
- [ ] **P1** A chart in a document cannot be edited (data, type or formatting), and the load note
      still calls it a labelled placeholder (`docx/parse-warnings.ts`).
- [x] **P2** Draft view's header, footer and note previews leave floating (`wp:anchor`) drawings
      out, as Word's Draft view does, and show inline pictures from the body's media cache
      (SVG originals and cropped or recoloured fills included) (`src/ui/src/docx/story-drawings.test.ts`).
- [x] **P1** Print Layout reserves space for tall headers and footers based on their in-flow
      content (#313). Paragraphs that move onto a page reflow around that page's floating drawings,
      including drawings anchored below them (#312). Print Layout and PDF export use the same
      story geometry for default and first-page footers, paragraph borders and indents, tab
      stops and pictures anchored inside table cells (#331, #332, #333, #342).
- [x] **P1** Single-spaced lines use the selected font's natural line height. The 300-line Word
      references agree on the first-page break and page count (#330).
- [x] **P2** All 44 page starts in the anonymized contract from #330 match between Print Layout in
      Edge on Windows and Word 16 COM, including the leading blanks and the split table row.
- [x] **P2** Word 16 COM and Windows Print Layout agree on the page and visible source position
      of all 2,238 body and table lines of the anonymized #330 contract, and on all 600 lines of
      its two single-spacing controls. Regressions cover Word's breaks at dashes, authored
      indentation, overlong words across formatting runs and modern justification.
- [ ] **P2** Word's justification near a line's fitting threshold is still approximated. Font
      substitutions can also shift line breaks. Literal U+00AD in a text run and the discretionary
      `w:softHyphen` element still share the same representation in Print Layout.
- [x] **P2** Overlong words in Print Layout break at grapheme boundaries (#316). Nearly fitting
      final lines in justified text boxes have a one-pixel fit tolerance (#314).
- [x] **P1** Auto-fit tables stay inside the text column when a cell holds an unbreakable token
      or many columns need more than their share; the token breaks inside its cell (#320). Header
      and footer tables draw their borders (#321), and Draft view keeps table widths inside the
      paper (#322). Checked against Word 16 with `src/core/docx/layout/fixtures/table-reports/`,
      where a "keep row together" row taller than a page is clipped on one page, as Word does (#323).
      A repeating header row that moves to the next page is drawn there once (#324).
- [x] **P1** Save as PDF (File > Export > Create a PDF, `exportPdf()`, `exportDocxPdf`) writes real
      text, vector borders, shading and rules, and pictures at their own resolution on the Print
      Layout pages, with headers, footers, page fields and footnotes, a page range, document
      properties and a report, without a print dialog (`src/ui/src/docx/pdf/`, #318).
- [x] **P1** Save as PDF keeps hyperlinks, links to bookmarks, `\h` cross-references and table of
      contents entries clickable, writes headings or Word bookmarks as the PDF's bookmarks, tags
      the PDF for accessibility (headings, paragraphs, tables, figures with alt text, links;
      headers and footers as artifacts), writes PDF/A-2b metadata and an sRGB output intent, and
      with markup shows tracked changes in author colours with change bars and comments as notes
      (`src/ui/src/docx/pdf/pdf-hooks.ts`, `export/pdf-structure.ts`, #318).
- [x] **P1** Save as PDF writes charts and drawing shapes as vector paths with real text (through
      the shared SVG to PDF converter), text boxes with the lines the page engine laid out,
      watermarks (text turned and half transparent, pictures washed out), page borders and line
      numbers where Word's PDF puts them, and floating objects in headers, footers and table cells
      (#318).
- [x] **P1** Save as PDF writes charts and shapes without a DOM (their SVG is read with the DOM-free
      XML parser when there is no `DOMParser`), draws shadow blur as a soft-masked raster under the
      vector shapes, draws WordArt warps from bent glyph outlines of the embedded fonts (with
      invisible text for selection) and outline, gradient and shadow text effects, tags lists as
      `L`/`LI`/`Lbl`/`LBody` with `ListNumbering`, makes links in headers and footers clickable on
      every page, shows formatting changes ("Formatted: ...") and comment balloons in a markup pane
      as Word's Print Markup does (`markupBalloons`), has an Options step like Word's Publish as PDF
      Options, and is validated with veraPDF (`scripts/pdfa-validate.mjs`): with the host's font
      files every corpus document passes PDF/A-2b and 2u, and the file declares 2u (#318).
- [x] **P1** Save as PDF: text watermarks are stretched to their VML shape; Word shape and picture
      shadows are read by the core and drawn in both Print Layout and PDF; chart SVG text drop
      shadows and WordArt glow and reflection are drawn; lists in footnotes, endnotes, headers,
      footers and text boxes are indexed and tagged as lists, and links in floating text boxes in
      headers and footers are clickable; table, row, cell and section property changes
      (`w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, `w:sectPrChange`) get formatting balloons;
      and `pdfALevel: 'A'` (PDF/A-2a) and `pdfUa: true` (PDF/UA-1) validate and declare conformance
      honestly (#318).
- [x] **P1** Opening a PDF converts it into an editable document, as Word's PDF Reflow does
      (`ooxml-core/docx/pdf-import`, `ooxml-core/pdf`; File > Open accepts `.pdf`, #319): flowing
      paragraphs with rejoined hyphenation, character formatting (font, size, bold, italic,
      underline, colour, super- and subscript), alignment, indents and spacing, size-ranked
      headings, running heads and feet as the header and footer with PAGE and NUMPAGES fields,
      pictures (JPEG as is, other samples as PNG), external links, page size and margins, document
      properties, passwords (RC4, AES-128, AES-256), permission flags and a per-page report.
      Checked against PDFs Word 16 exported and pypdf encrypted.
- [x] **P1** PDF conversion rebuilds ruled tables (merged cells, borders, shading, continued on
      the next page, rows a page break cut joined back), borderless tables including wrapped and
      empty cells, bulleted and numbered lists (also inside table cells) and two- or three-column
      regions (continuous `w:cols` sections), footnotes (also continued on the next page, or
      starting on the page after their mark), endnotes collected at the end of the document,
      internal links (to bookmarks), headings from tags and bookmarks, rotated text (as turned text
      boxes), vector drawings as floating DrawingML shapes and `wpg` groups, and decodes JBIG2 and
      JPEG 2000 pictures (#319).
- [x] **P1** PDF conversion rebuilds gradient fills (axial/radial shadings and shading patterns),
      tiling pattern fills, rectangularly clipped vector shapes, and text enclosed inside drawing
      shapes (as `wps:txbx`), extracts endnotes collected at the end of each section (`sectEnd`),
      joins table rows split between two sentences when empty cells are normal, applies embedded
      JPEG 2000 ICC Matrix/TRC profiles, decodes JBIG2 T.88 Amendment 2 extended generic templates
      (`EXTTEMPLATE=1`), and renders `wpg:wgp` groups as vector SVG shapes in the Word editor
      instead of placeholders (#319).
- [ ] **P2** PDF conversion still reports (without silently dropping) non-rectangular clipping
      paths, mesh shadings (Type 4-7), JBIG2 colour extensions, High Throughput JPEG 2000 and JPX
      Part 2 extensions; scanned pages stay pictures unless the host supplies OCR (#319).
- [ ] **P1** Pagination is approximate (no Word-rendered page corpus yet); caret placement at line
      and page boundaries is heuristic; no kerning or complex-script shaping. Single line heights
      come from the browser's
      font box without the font's line gap: Arial 11 pt lines are 16 px apart in Chromium against
      Word 16's 16.87 px (measured in text boxes), so a box or page can take one more line than Word.

### Ribbon

- [x] **P1** Mailings: Start Mail Merge (letters, directory), Select Recipients (local CSV, TSV
      or `.xlsx`, or the host's `mailMergeDataSource`), Edit Recipient List, merge fields and
      Rules, Match Fields, Preview Results, Check for Errors and Finish & Merge to a document or
      print, loaded on demand (`src/ui/src/docx/mailings/`, `e2e/docx`).
- [x] **P1** Mailings: Envelopes, Labels and Update Labels, Word's native `ADDRESSBLOCK` and
      `GREETINGLINE` fields (checked with Word 16 field updates), interactive ASK and FILL-IN
      prompts, and Send Email Messages through a host `emailProvider` (2026-11-26).
- [x] **P2** Labels: Label Options searches a catalogue of 39 sheet products (Avery US Letter,
      Online Labels, Avery A4, Avery Zweckform, HERMA) with their inkjet variants as aliases,
      loaded on demand (`docx/mail-merge/label-catalogue.ts`), and the engine takes any custom
      layout, vertical gaps included (2026-11-27).
- [ ] **P2** Labels: Word's catalogue lists far more vendors and products, and Label Options has
      no New Label (custom dimensions) or Details preview yet.
- [x] **P1** Design: Themes, Colors, Fonts, Paragraph Spacing and Effects galleries (shared
      theme presets in `drawingml/theme-presets.ts`), custom colours, fonts and spacing, and Set
      as Default through the host's preference store, each one undo step and saved
      (`docx/document-design.ts`, `e2e/docx/document-design.spec.ts`).
- [x] **P1** Design: the Word theme gallery with each theme's measured palette, fonts and format
      scheme, custom colour and font schemes, `.thmx` import and export, and the Effects gallery
      applying whole format schemes (`drawingml/theme-gallery.ts`, 2026-11-26).
- [ ] **P1** Design: nine legacy Word themes are omitted until their values are verified, and
      theme picture fills (Grunge Texture, Wood Type) are replaced with solid fills on import and
      not rendered, with 3-D styles.
- [x] **P1** Table Design and Layout tabs: table styles and style options, Cell Size and AutoFit,
      cell alignment and margins, Repeat Header Rows, Formula, Convert to Text, Select and Draw
      Borders (2026-11-27).
- [x] **P2** Word tables: Layout > Draw Table (a rectangle in the body draws a one-cell table; a
      line inside a cell splits it into columns, cutting a grid column when needed, or into rows)
      and the Eraser (erasing an inner line merges the two cells beside it when they form a
      rectangle, an outer line loses its border), each one undo step, Esc to stop
      (`docx/table-draw.ts`, `src/ui/src/docx/table-freeform.ts`, 2026-11-27).
- [ ] **P2** Word tables: Draw Table lines across several cells at once, dragging the Eraser along
      a line, a drawn table placed where it was drawn (it goes before or after the paragraph under
      the pointer, at the column's left edge), erasing outer lines of a table with merged cells, and
      Split Cells into arbitrary counts when the grid must be rebuilt.
- [x] **P1** Equation tab: Insert New Equation and Alt+=, the built-in gallery, Convert
      (Professional and Linear), Normal Text, Unicode and LaTeX input, symbol and structure
      galleries, and in-place editing of linear text that rebuilds the equation exactly.
- [ ] **P2** Equation: an equation whose linear text cannot hold it (mixed run formatting,
      tracked changes inside it, properties its notation cannot spell) opens read-only and the
      editor says which. A display paragraph with several equations is edited one equation at a
      time, the others kept byte for byte (`math/equation-parts.ts`). Ink Equation opens the Ink
      Equation dialog with the host's recognizer (`setInkMathRecognizer`, none built in; disabled
      with its reason without one; `docx/ink-equation.test.ts`).
- [ ] **P1** Missing tabs: Picture Format has only Wrap Text, Crop and Format.
- [x] **P2** Draw tab: drawing, the eraser, the lasso, Ink to Shape, Ink Replay and the clipboard
      work and save as Word ink (`docx/ink.test.ts`, `e2e/docx/draw-tab.spec.ts`). The Ruler
      (Stencils) lies on the pages as in PowerPoint and Excel (shared `office-ui-ink-canvas`),
      turned with the wheel or Shift and the arrow keys; it scrolls with the pages and does not
      measure lengths. Ink to Math opens the Ink Equation dialog
      with the host's recognizer, as Ink Equation does, and inserts at the caret; the selected ink
      on the page is not sent (write on the dialog's pad).
- [x] **P1** Layout and Shape Format: the Arrange group (Position, Wrap Text, Bring Forward, Send
      Backward, Selection Pane with rename, Align, Group, Rotate) and Text Direction, with canvas
      multi-selection (`docx/arrange-controller.ts`, 2026-11-26).
- [ ] **P1** Draft and Web Layout draw floating objects with centred or bottom `wp:positionV`
      at the anchor paragraph top, so Position presets require Print Layout.
- [x] **P1** Insert WordArt creates editable OOXML text boxes with fill, outline and shadow
      presets and DrawingML text transforms.
- [x] **P1** Wide folded ribbon groups scroll within the phone screen. Save, Undo and Redo
      stay available in the compact Quick Access Toolbar.
- [x] **P1** The page thumbnail rail scales to fit narrow editors and keeps the document visible
      in both suite layouts at 320×256 and 360×225.
- [x] **P2** Word and Excel share 36 WordArt styles, including six two-stop gradient presets,
      and a text transform picker. Gradient stops, the angle, preset names and adjustment guides
      survive native DOCX and XLSX saves. The editors preview the gradient and text path.
- [ ] **P2** Word and Excel WordArt still need multi-stop and themed gradient editing, 3-D effects,
      direct adjustment handles and glyph-envelope rendering for deformation presets.
- [x] **P1** Insert: Quick Parts (Document Property controls bound to the core and extended
      property parts, Field, Save Selection to Quick Part Gallery and the Building Blocks
      Organizer over the glossary part), Signature Line (VML signature line) and Online Video
      (a wp15 web video picture, on when the host sets `onlineVideoHosts`). Header and Footer
      offer designs built from content controls, plus Edit and Remove.
- [ ] **P2** Insert: Online Video has no search and plays only hosts the application allows;
      Signature Line cannot be signed (no digital signature UI).
- [x] **P1** Insert group order, command placement and label fit at 1600px and 1280px match Word
      in all six bindings (`e2e/docx/ribbon-insert-order.spec.ts`, 2026-11-27).
- [x] **P1** Insert > 3D Models embeds a self-contained GLB 2.0 with a PNG fallback at the caret.
      Insert > Object embeds an existing DOCX, XLSX or PPTX as a native inline file icon.
      The original bytes and preview survive save and reopen.
- [ ] **P1** Insert still needs stock models, other model formats, interactive 3D controls,
      Create New and link-to-file.
- [ ] **P2** View: Window group, Macros; Review: Thesaurus, Translate, Block Authors, Hide Ink;
      Home: Text Effects, Phonetic Guide, Character Border, Enclosed Characters; Backstage:
      Share, Close.

## PowerPoint

### Cloud storage integration

- [x] **P2** Optional shared host OneDrive connection resolves sharing links and saves native DOCX,
      XLSX, PPTX and VSDX through Graph and each editor's byte/storage contract. Sessions retain
      baseline/local/remote versions on content conflicts and support separate cloud-copy recovery;
      revision-only changes are retried once when the remote content is unchanged. The suite tracks
      pending cloud saves separately from local checkpoints and exposes explicit sharing scopes.
      Test transports are test-local; no public mock service or simulation UI ships.
      See [integration and package reference](onedrive-integration.md) and
      [acceptance matrix](onedrive-testing.md).
- [ ] **P1** Real OneDrive personal/work and SharePoint tenant acceptance, including consent,
      deployed-browser CORS, permission revocation, cross-tenant and policy-restricted links.
      The adapter requires host authentication and an item-specific permission resolver; unknown
      access is read-only. Mock tests do not close this acceptance gap.
- [ ] **P1** Microsoft Office live coauthoring and locks are not provided by the OneDrive storage
      adapter. Revision checking at upload-session creation is not a lock through final commit.

### File format and preservation

- [x] **P1** Linked pictures and media follow PowerPoint's default: nothing linked is downloaded
      when a deck opens, and a translated security bar offers Enable Content for that
      presentation (all five bindings and the suite; host `allowExternalContent` changes it).
      Enable Content redraws without reloading, keeps undo history, and is each collaborator's
      own choice.
      Evidence: [powerpoint-external-content.md](powerpoint-external-content.md),
      `e2e/pptx/allow-external-content.spec.ts`.
- [ ] **P2** Blocked linked content differs from PowerPoint in three ways
      ([powerpoint-external-content.md](powerpoint-external-content.md)): a blocked picture has no
      "PowerPoint has blocked automatic download of this picture" placeholder, intranet hosts are
      blocked too (PowerPoint downloads them), and after Enable Content an online video's frame
      loads with its slide (when PowerPoint first contacts the provider was not observed).
- [ ] **P1** 40 coverage-manifest constructs are `preserve: unassessed` and 30 `edit: unassessed`
      (`src/core/pptx/core/openxml/openxml-coverage*.ts`): add a dirty-save round trip with unknown
      elements on every part type, and ratchet the counts in CI.
- [ ] **P1** Chart save is not native-certified for every fresh template. New ChartEx parts now
      carry the chart style and colour style parts Office requires (without them Word and PowerPoint
      refuse the file): a fresh PowerPoint deck, Word document and Excel workbook with one chart of
      each of the eight families open in Office 16 and keep every family and value on Office's own
      re-save (`src/core/scripts/chartex-acceptance.ps1`, 2026-11-16).
- [ ] **P2** PowerPoint chart edits still saved only by the object-tree writers
      (`chart-neutral-save-plan.ts` leaves them unclaimed): cleared group subtype settings those
      writers keep, a cleared series number format, a category edit made in the same save as a
      whole-chart switch to XY, and any part the neutral parser cannot preserve or the shared
      writer refuses. Every field of the 2026-11-16 probe now has a neutral family: group subtype
      settings (`chart-neutral-save-subtype.ts`); `varyColors`, bar direction, scatter style and
      the `standard` grouping flag (`chart-neutral-save-group-flags.ts`); the gridline flag; a
      series' own number format; the auto-update flag, colour map override, colour style part,
      print settings, pivot formats and user shapes (`chart-neutral-save-part-settings.ts`); and
      Chart Filters hide, restore and drop (`chart-neutral-save-filtered.ts`). The object-tree
      writers had dropped `varyColors`, bar direction, scatter style, the gridline flag and the
      series number format on save. Print settings and pivot formats are still encoded by
      PowerPoint's element codecs before the shared writer places them, so those codecs and the
      object-tree writers are not deletable yet.
- [ ] **P2** Per-series chart combinations: a series of a 2-D column, bar, line, area or XY chart
      can become column, bar, line, area or scatter (`chart/edit/series-type.ts`): a series that
      turns between horizontal bars and a vertical family takes a secondary axis pair, a series of
      an XY chart that takes a category family turns the x axis into a category axis, and a series
      converted back rejoins the group it left, as Excel 16 saves each through `Series.ChartType`.
      Splits inside a stack, groups with unknown extensions and a turn on a chart that already has
      secondary axes keep the object-tree path; 3-D or bubble series types switch the whole chart,
      as Office does. Groups of one type that are not neighbours after a conversion stay separate.
- [x] `a:linkedTxbx` (Word-style linked text boxes) is not flowed: PowerPoint ignores it (measured
      with COM `Slide.Export`: the head keeps all text unclipped, successors stay empty), so the
      renderer matches by default (`setLinkedTextBoxFlow` opts in). The link is parsed and
      round-trips (`linkedTxbxId` / `linkedTxbxSeq`).
- [x] **P0** Copies never share an owned part: a duplicated or pasted chart, diagram, OLE object
      or media element, a duplicated slide, master or layout and an inserted master each save with
      their own chart, workbook, diagram, embedding, theme and layout parts and relationships, and
      unique master and layout ids (`element-part-binding.ts`, `master-layout-crud-*.ts`; checked by
      `findSharedOwnedParts` in `part-ownership-sweep.test.ts`; all 110 sweep outputs open in
      PowerPoint 16, 2026-11-20). Before, PowerPoint refused such files (0x80070570).
- [ ] **P2** Elements copied between a layout or master and a slide keep the relationship ids of
      the part they came from: save binds only slide elements (`element-part-binding.ts`).
- [ ] **P1** Collaboration shares package-local part names and relationship ids (`chartPartPath`,
      `r:id` in `rawXml`, `collaboration-sync.ts`). A peer saving an element another client saved
      first can cite a part or relationship only that client's package has.
- [ ] **P1** `.ppt` export turns video and non-WAV audio into pictures.
- [ ] **P1** Clipboard interop with desktop Office is partial. Done: HTML (`text/html`, CF_HTML)
      pastes into a text box being edited as formatted runs, bullets, numbering and links
      (Keep Source Formatting; Ctrl+Shift+V keeps text only), and copying shapes writes a PNG and
      their text to the system clipboard. Open: copied slides add no picture yet (no binding
      supplies `renderSlidePng`); pasting HTML onto the slide (not into a text box) does not create
      a text box; no RTF or Office native formats (`Art::GVML`, PowerPoint 14.0 Slides) are read
      or written; a rich paste can take two Undo steps.
- [x] **P1** Embedded fonts honour OS/2 `fsType` embedding permissions: restricted faces are not
      displayed but their parts are saved unchanged; restricted, preview and print, and
      bitmap-only faces are never newly embedded; the Embed Fonts panel shows the restriction in
      every binding (`core/text/truetype/embedding.ts`).
- [x] **P1** Digital signatures are verified in the browser (part digests, signature value and the
      signer certificate's key) and every binding shows the verified state through
      `signatureStatusView` (`src/ui/src/pptx/render/signature-status.ts`): a deck reads as signed
      only when its signatures verify, and invalid or unsupported signatures say so in the badge,
      the signatures panel and its rows (`e2e/pptx/signature-state.spec.ts`). There is no trust
      store in a browser, so a valid signature is "Signed, signer not verified" at best; the viewer
      cannot sign, and saving an edit removes the signatures.
- [ ] **P2** 3-D models: `am3d:preTrans` is parsed, not applied; `rot` handedness and
      orthographic scale are unverified.
- [ ] **P2** `p:smartTags` is parsed and its recognizer part is kept, but has no editable model;
      `p:revisionInfo` survival is unverified.

### Editing and slide show

- [x] **P2** Tab and Shift+Tab in a text box: Tab at the start of a bulleted or numbered
      paragraph (or across a multi-paragraph bullet selection) demotes its list level (up to 8),
      Shift+Tab promotes it (down to 0), and Tab elsewhere inserts a tab character while Shift+Tab
      keeps the caret in the text box, in all five bindings (`src/ui/src/pptx/render/inline-list-tab.ts`,
      `inline-list-controller.ts`, `inline-list-tab.test.ts` and per-binding regressions).
- [x] **P1** Table cells are edited as rich text with each run's markup kept (hyperlinks, fields
      until their own text is edited, theme colours and transforms, untyped run properties):
      Enter adds a paragraph, Tab/Shift+Tab and edge arrows move between cells, Tab in the last
      cell adds a row, Esc keeps the text, tab-separated paste or Ctrl+V into selected cells fills
      and grows the table, and Home-tab formatting applies to every run of a cell range or the
      whole table (all five bindings). Open: the ribbon formats whole cells, not the text selected
      inside a cell.
- [ ] **P2** Format pane: the Properties tab shows selection-aware collapsible sections in all five
      bindings, and Format Shape, Size and Position and Alt Text open it on their section. Open: a
      multi-selection shows only a section-level "Mixed" badge (no per-field mixed values) and
      edits apply to the primary element only; Vue shows the slide inspector for
      multi-selections; tables have no table-wide Font/Paragraph section; section reordering is
      visual only (tab order unchanged).
- [ ] **P2** Format Shape > Fill: the six kinds (no fill, solid, gradient, picture, pattern, slide
      background) are one shared section in all five bindings, and a slide background fill saves
      as `useBgFill`. Open: a shape switched to a slide background fill draws no fill until the
      deck is reopened (the canvas does not resolve the slide background onto it); no gradient
      presets, direction gallery or per-stop brightness; the picture fill has no tile offsets,
      scale, alignment or mirror fields.
- [x] **P1** Save as PDF writes vector pages at the slide size: shapes, gradients, pictures, inline
      SVG charts and selectable text with the deck's embedded fonts subset in; hidden slides are
      left out unless Options > Advanced > Save as PDF includes them; effects PDF cannot express
      (soft edges, blur, glow, shadows, 3-D, blend modes, patterns, complex scripts) are pictures
      of just that element with searchable text kept, a video that has not played prints as its
      poster frame, and the export lists what it approximated
      (`src/ui/src/pptx/export/pdf-vector-*.ts`). Text no available font has (Chinese, Japanese
      or Korean in a deck that embeds no font for it) is a picture of the run with an invisible
      text layer in a glyphless font, counted in the notice, not left out; vertical text keeps
      the same layer, one character at a time. Open: fonts the page cannot read fall back to
      standard PDF fonts, so such CJK text is a picture rather than real glyphs (no system or web
      font is read or downloaded for the PDF); embedding a deck's own fonts is unit-tested only.
- [x] **P1** Save as PDF says what it approximated: one translated notice counts elements and
      slides saved as pictures, substituted fonts, and slides with content left out or effects
      simplified, with a Details list; all five bindings emit `pdfExportReport`, and a later clean
      export dismisses any earlier notice (`export/pdf-export-notice.ts`, `export/pdf-export-run.ts`,
      `web-components/pdf-export-dialog.ts`). Export > PDF first offers Range (All, Current slide,
      Slides from-to) and Include hidden slides.
- [x] **P1** Save as PDF (vector, the default) is fast on effect-heavy decks and its element
      pictures hold only their element. Each slide's styled clone is built once and every picture
      is cut from it (`export/pdf-vector-raster-batch.ts`); a picture shows only its target
      (shapes declaring `visibility: visible` used to leak into every picture as misplaced ghost
      copies), draws the stage at its own layout size (the canvas zoom used to shrink every
      picture's content to about 0.87 of its place, so pictures cropped the wrong region), and
      leaves the editor's action badges and link tooltips out (`render/export-clone.ts`). The
      14-slide test deck went from about 10.5 s to 0.7 s per slide locally (109 s to 8 s).
- [ ] **P2** Save as PDF options lack Selection and custom-show ranges, Publish what (handouts,
      notes pages, outline), document properties, accessibility structure tags, PDF/A and
      bitmap-text options; "Save slides as pictures" has no options step and writes every slide.
- [x] **P2** Table cells: explicit `bold: false` (`b="0"`) inside bold cells emits normal weight
      (`table-style.ts`), Home-tab highlight, character spacing, text shadow and change case apply
      to selected table cells (`table-cells-format.ts`, `table-cells-format.test.ts`), and `Esc`
      restores cell selection in Vanilla and Svelte as well as React, Vue and Angular.
- [x] **P1** Opening a deck paints a read-only preview of the opening slide before the rest
      parse, then fills thumbnails progressively (`progressive-slide-preview.ts`); browser timing
      benchmarks (`e2e/pptx/progressive-open.spec.ts`) and a large-deck work budget guard it
      (2026-11-26).
- [ ] **P1** Editing, search, save, undo, collaboration and slide show wait for the whole deck
      to parse after a progressive open. The rules for using a partly parsed deck (placeholders
      filled only by their own arrival, edits on parsed slides, queued whole-deck work and save,
      undo snapshots filled on restore, search coverage, slide show waiting at a placeholder) are
      the shared `createProgressiveDeckController`
      (`render/controller/progressive-deck-controller.ts`); the five bindings do not use it yet
      and still mount the editor only when `load()` resolves.
- [ ] **P2** The File page footer always names "Untitled Presentation.pptx", and the status bar
      reports an AutoRecover snapshot as "Saved" while the presentation has unsaved edits.
- [ ] **P1** Headless SVG export (`pptx/converter/SvgExporter.ts`, the CLI `export-svg` and the
      `export_slide_svg` tool) uses a simplified renderer that trails the editor (plain `<text>`,
      no warp, fewer effects); server output does not share the editor's paint.
- [ ] **P2** Automation is element-level MCP tools only: no Office.js-shaped object model
      (`RequestContext`/`load`/`sync`, text ranges, table borders and margins, adjustments, custom
      XML) for add-in style scripts.
- [ ] **P2** Customization hides or replaces regions and capabilities, not single toolbar controls,
      and dialogs have no headless draft/validate/apply hooks.
- [ ] **P1** Remove Background is authored (automatic guess in a worker, keep and remove marks)
      and round-trips here, and desktop PowerPoint 16 opens and draws the result, but PowerPoint
      drops the `a14:backgroundRemoval` state when it re-saves the file (also with an `hdphoto`
      relationship to a PNG original), so the removal is not editable there; HD Photo originals
      are not re-edited.
- [x] **P1** Summary Zoom creates a new overview slide before the first chosen slide, reuses
      existing sections, and starts new sections at the chosen slides.
- [x] **P1** Slide Zoom retains its destination and Section and Summary Zoom follow the beginning
      of their sections after slide reordering. A save/reopen test checks the original slide ID.
- [x] **P1** Zoom insertion offers Return to Zoom, duration and transparent thumbnails. Slide
      and Section Zoom accept a local cover image. The resulting options save in native PPTX.
- [x] **P1** The editor and live show refresh slide previews when destination artwork changes.
      Custom covers retain their embedded image.
- [x] **P1** Slide Zoom, Section Zoom and individual Summary Zoom tiles can change return
      behavior, duration and thumbnail transparency after insertion in the Format pane.
- [x] **P1** Browser saves refresh embedded Zoom thumbnails when their destination or format changes.
      Custom covers stay intact.
- [x] **P1** The Format pane can change or reset the cover image for a Slide or Section Zoom
      and for individual Summary Zoom tiles.
- [x] **P1** Insert > Cameo adds a live-feed picture with a saved poster and explicit local camera
      control on the editing and live presentation surfaces.
- [x] **P1** The Record tab inserts Cameo on the current slide or on all slides without one.
      Clear Timings removes timed advances on the chosen slides and preserves their effects.
- [x] **P1** The Record tab can record the camera and optional microphone into an embedded clip
      on the current slide. Reset to Cameo removes editor-recorded clips, keeping an existing
      Cameo or inserting a live placeholder when the slide lacks one.
- [ ] **P1** Slide-show narration, recording directly into a styled Cameo, native PowerPoint
      recording interoperability and Insert > Record Audio remain open.
- [ ] **P1** Animation: a by-paragraph build (`sequence: 'byParagraph'`) is written as one
      whole-shape effect with `p:bldP build="p"`, where PowerPoint writes one effect per paragraph
      (`p:txEl/p:pRg`), and such per-paragraph timing trees are not read back as one high-level
      animation (`pptx/core/services/animation-write-sequence-builders.ts`).
- [x] **P1** Animation: Effect Options directions for Split, Zoom, Blinds and the shape reveals,
      Color Typewriter, emphasis effects from the file's own behaviours, and Blink timing
      (2026-11-27).
- [x] **P1** Animation Painter (single use, locked with a double-click, Esc, one undo step per
      apply) and the ribbon's Start, Duration, Delay and Move Earlier/Later work in all five
      bindings through the shared controller (`render/controller/animation-painter-controller.ts`,
      `e2e/pptx/ribbon-animation-painter.spec.ts`).
- [x] **P1** Animation Painter and the Timing fields edit imported native effects, the Animation
      Pane chooses the effect Timing edits, and painting across slides keeps an embedded sound
      (2026-11-27).
- [x] **P2** Changing Start on an imported effect splits a click group (On Click in its middle
      moves the effect, its later With Previous siblings and the later wrappers to a new click) or
      merges one into the previous group (With or After Previous on its first effect), rebasing
      the moved offsets and recomputing After Previous starts, through save and reload
      (`native-animation-start-regroup.ts`, `native-animation-start-regroup.test.ts`, 2026-11-27).
- [ ] **P2** Start changes stay refused, with the reason in the Start field, for triggered
      (interactive) effects, nested `p:seq`/`p:excl` timing, wrappers or effects that start on a
      condition other than a delay, and After Previous after an effect whose length is not
      measurable (repeat until next click or end of slide).
- [x] **P1** Ink: Lasso, pen presets, Ink to Shape, Ink to Math, Ruler and Ink Replay
      (2026-11-27).
- [ ] **P2** Ink Replay's per-point timing is lost after PowerPoint desktop saves the file again.
- [x] **P1** Custom shows have editing operations (`editor/render/custom-show-operations`), the
      `manage_custom_shows` MCP and assistant tool, valid unsigned-int ids on save (with `p:showPr`
      and `ppaction://customshow` references remapped) and the Custom Shows dialog (New, Copy,
      Remove, two slide lists with Add, Remove, Move Up/Down, Show; the shared
      `createCustomShowsController`) in all five bindings, with unit tests and
      `e2e/pptx/custom-shows-dialog.spec.ts`.
- [x] **P1** `manage_sections` and `update_presentation_properties` persist through
      `saveDeckBytes`, including empty section lists, false slideshow flags and selecting an
      existing custom show with `showSlidesCustomShowId`. Deleting the selected show saves the
      switch back to all slides. Empty `present`, `browse` and `kiosk` elements also reload with
      their show type (`src/core/pptx/automation/__tests__/presentation-save.test.ts`).
- [ ] **P2** Thin strokes: widths round to whole device pixels and horizontal, vertical and
      plain-rectangle strokes snap to the device grid like PowerPoint's export, every edge of a
      rectangle on its own, positioned from the file's exact EMU rounded to PowerPoint's 1/576
      inch master unit (`src/ui/src/pptx/render/device-pixel-stroke.ts`, `device-pixel-exact.ts`);
      measured exact against PowerPoint at 1280 and 1920 px wide. Lines without an `a:ln/@cap`
      now end flat like PowerPoint's (they were round). Group children snap too, by their absolute
      position (ancestor group origins from the exact EMU; `device-pixel-group.ts`; inside a
      rotated or flipped group they are not snapped); a group's connectors in React's static
      renderer now share the interactive connector's width and snap. Checked against PowerPoint on
      a synthetic group deck: 21 differing pixels at 1280 px and 144 at 1920 px (exact-tie rows),
      all five demos identical. Open: PowerPoint's export at other widths is not a function of
      the position (measured with 110 probe lines per axis: 1024 px wide is anti-aliased, 1366 px
      wide leaves about 24% of lines on a different pixel under the best fit, at 16:9 and other
      aspect ratios alike, so the 768/540 vertical scale is not the cause); we keep the 1280 and
      1920 rule there. Rotated, rounded and gradient outlines, compound strands and CSS borders
      are not snapped.
- [x] **P1** Morph of elbow and curved connectors: the route is re-evaluated per frame from the
      interpolated box, signed flips and `adj` values at a constant stroke width, with no outline
      dissolve (`src/ui/src/pptx/render/morph-connector-carry.ts`); `curvedConnector3` to `5`
      follow the ECMA-376 definitions PowerPoint draws (a static export differs by under 0.05 %
      of the pixels, the earlier approximations by about 0.5 %).
- [x] **P2** Morph of a connector or line whose size changes a lot: PowerPoint grows each extent
      of the box ahead of its eased progress while the centre follows it exactly. Per axis,
      extent(p) = lerp + |delta| * hi / (hi + lo) * g * p * (1 - p), larger than the blend for
      grows and shrinks alike, with g = (1 - exp(-3.2 (r - 1))) * coupling(x), r = hi / lo (a 1.1x
      resize has none, 1.4x 70%, 2x 95%) and coupling = x^1.14 (x <= 1) or min(1, 1.1 / x^0.94)
      where x is the centre's travel over half the size change
      (`src/ui/src/pptx/render/morph-line-reach.ts`, fitted to 43 synthetic cases, 1.5 pt mean
      error against 40 pt; the earlier law was 4 pt off at 1.4x). Straight lines also follow it
      when rotated (PowerPoint blends the end points, so a pure turn shortens the chord and a
      turn while the centre moves is made up again; 2 to 3.5 pt on seven cases, one turn plus
      resize plus move case still 12 pt off), when their ends flip differently (the extent
      passes through zero and overshoots by hi * (r - 1) / (r - 1 + 0.56) * coupling, 1 to 5 pt
      on seven cases) and when two ends share a rotation. Freeform open paths and arcs do not
      overshoot in PowerPoint (measured linear). Elbow and curved connectors that turn or flip
      differently stay linear (not measured).
- [ ] **P2** Notes Master and Handout Master are placeholders; no Notes Page view; no Grayscale or
      Black and White view.
- [ ] **P2** Designer, Rehearse with Coach, Keep Slides Updated and Media
      Controls are disabled.
- [ ] **P2** Effect DAGs (`a:effectDag`) are painted as one effect list plus colour operations
      (`interpretDrawingEffectDag`, `src/core/drawingml/effect-dag.ts`): effects and colour
      operations inside `a:cont` and `a:blend` containers are found, `alphaModFix` and `alphaRepl`
      alpha multiplies, and the first of each list effect and colour operation is painted. Still
      approximated, and listed per shape in `effectDagApproximations`: sibling outputs are not
      composited separately, `a:alphaMod` masks are reduced to their alpha, non-`over` blends paint
      as `over`, `a:xfrm`/`a:relOff` are reported but not applied, and `a:effect` references,
      `a:fill`, `clrChange`, `clrRepl` and the alpha threshold operations are not painted. The
      committed decks hold no effect DAG, so this is unit-tested against the schema, not checked
      against PowerPoint output.
- [ ] **P2** A non-rectangular shape with a line effect or a 3-D bevel still loses its outer shadow
      and glow (a clip-path cuts them off); solid, unfilled, gradient, pattern and picture shapes
      (and their inner shadow) paint them as a shape-following filter
      (`src/ui/src/pptx/render/silhouette-effects.ts`; the fill of a gradient, pattern or picture
      is a clipped child layer, `silhouette-fill-layer.ts`). A morphing shape of that kind keeps
      its shadow and glow (its fill path morphs, `morph-shape-carry`) and now dissolves them the
      way PowerPoint does (`morph-effect-silhouette`): the arriving filter carries both ends'
      layers, scaled into its frame and summed, at `1 - p` and `p` of the morph's eased progress
      (a registered `--pptx-morph-p` property the carry keyframes animate; every painted copy of a
      slide has its own filter ids, `svg-id-scope`, so the CSS-driven opacity resolves in the copy
      that paints it). All five bindings agree exactly on a triangle to star shadow morph (1.79
      mean abs diff against PowerPoint at 50 %, from 2.31). Under a ghost only the incoming end
      fades in (the ghost carries the outgoing one); a pair that is not carried keeps full
      strength. Static silhouette
      shadows match PowerPoint within 0.4 to 0.9 mean abs diff (hollow ones included); hollow shadows
      at a 20 pt or larger blur read about 12 % too dark under the bottom edge only. In-place text
      editing in a text-bearing silhouette shape works in all five demos (click places the caret,
      typing edits); Angular's textarea editor ignores the click position, for plain shapes too.
- [ ] **P2** `a:path path="shape"` gradients: presets scale their outline about the focus and
      block arrows, other listed presets and every freeform paint the circle to the farthest corner
      (measured on 132 presets, 14.1 to 2.9 mean abs diff; the freeform SVG paint also draws true
      circles and `scaled="0"` angles). Still approximate: the ring presets (donut, frame, blockArc,
      noSmoking, chord, funnel, circular arrows, math divide and not-equal, tabs) and the
      multi-face presets (cube, can, bevel, magnetic disk, drum and tape), whose faces
      PowerPoint shades with their own fields.
- [ ] **P2** Preset geometry that still differs from PowerPoint 16 (measured with
      `Slide.Export` of every preset at a non-square aspect, default and adjusted, against
      `evaluatePresetShape`): the leader line of the line callouts whose COM default adjustments
      differ from the saved XML. `swooshArrow` (guide chain), the `actionButtonHelp` question
      mark (elliptical arcs) and the scroll curl swirl are closed in `fix(geometry)`. `leftCircularArrow` with a positive `adj2` (PowerPoint writes it modulo a turn) is
      pinned to the head angle 0 as the specification does. Closed in `fix(geometry)`:
      star4-star10, pentagon, heptagon, decagon, dodecagon, the math operators, the three tab
      shapes, braces and brackets, moon, heart, cloud, lightningBolt, wave, doubleWave, the arc
      family (visual arc angles), five multi-headed block arrows, smileyFace, can, cube, bevel,
      diagStripe, nonIsoscelesTrapezoid, circularArrow, leftCircularArrow,
      leftRightCircularArrow, ribbon, ribbon2, ellipseRibbon, ellipseRibbon2, leftRightRibbon,
      verticalScroll, horizontalScroll, funnel, irregularSeal2, flowChartMultidocument,
      flowChartMagneticDisk and the glyphs of the twelve `actionButton*` (flat face plus
      darken, darkenLess and lighten sub-paths; the binding-local glyph overlays are gone).
- [x] **P2** Sub-path shading (`lighten`, `lightenLess`, `darken`, `darkenLess`: can lid, cube,
      bevel, smiley eyes, ribbon folds, action-button glyphs) also applies under a gradient,
      pattern or picture fill, as a translucent white or black veil over the shaded sub-paths
      (`buildSubpathFillOverlay`, all five bindings). A `fill="none"` sub-path (the smiley's
      mouth) is still filled by the merged clip of a non-solid fill.
- [x] **P1** WordArt Transform: the closed-loop warps are drawn from PowerPoint's own geometry.
      Ring Inside and Ring Outside are two offset ellipses with the text run once round and the
      ink clipped to the ellipse of its band; Circle Pour is the box ellipse and its `adj2` scale
      (`text-warp-closed-paths.ts`); Circle starts at the left tip and runs clockwise, so a
      centred line sits on the right (`text-warp-arch.ts`). Mean absolute diff vs PowerPoint on
      the synthetic deck fell from 25.7 and 30.8 (rings), 24.3 (Circle Pour) and 16.9 (Circle)
      to about 6, 6.5, 10 and 4.6.
- [x] **P1** WordArt Transform, start angle and fit: the start angle of Arch Up, Arch Down,
      Button and Circle only picks the half of the ellipse (90 to 269 degrees the upper half,
      otherwise the lower, turned half a turn so Arch Up reads upside down; Circle centres on the
      left tip from 0 to 179 degrees) and text longer than the arc (a long line, or a start angle
      near the flip, which leaves a short sweep) is scaled to fill it exactly
      (`text-warp-arch.ts`, `text-warp-arch-fit.ts`, `warpParagraphLayouts` in all five
      bindings). Further lines lie one line pitch (1.2 em) apart on concentric ellipses, the line
      nearest the box centre on the base path, not 0.7 scaled copies. The wave and pour presets
      name their first guide `adj1`: it was dropped on load, so Wave 4 and Double Wave ignored
      their amplitude (mean abs diff of the Double Wave and Wave 4 deck 8.27 to 2.81).
- [ ] **P2** WordArt Transform, remaining gaps measured against PowerPoint 16: the arch frame
      offsets (`ARCH_FRAME_INSET`) fit Arial 32 pt and drift by up to 4 px for other sizes and
      fonts (the drift is not explained by ascent, cap height or text width: Times New Roman
      needs none at 24 and 48 pt, Arial 6 px at 48 pt); Circle with several lines sits about
      10 px further out than PowerPoint's; Button with two or three lines is a top arc, a flat
      middle line and a bottom arc in PowerPoint, drawn here as a stack; Double Wave 1 and Wave 4
      end the top and bottom ink independently either at the box edge or one ~9 px piece short
      (chaotic in the adjust: 6000 and 6500 full, 6250 short at both ends, not a function of the
      path's x range or length), Inflate, Deflate and Can Up keep a residual of 6 to 8 grey levels
      from heavier glyph weight in PowerPoint's warped text rather than placement (edges fit a
      linear ink-normalised map within 2.5 px), and Circle Pour reads its text length slightly
      short of the loop.

### Ribbon

- [x] **P1** Insert WordArt creates editable text with a preset from the WordArt Styles gallery.
- [x] **P1** The phone Insert sheet opens WordArt, Icons, Symbol, Slide Zoom, Section Zoom and
      Summary Zoom in all five bindings. Cancel returns to the sheet and a successful insert
      closes it. Shapes preserves the preset picked on the phone.
- [x] **P1** Insert Symbol uses the common Unicode gallery or a character code at an active
      text caret, or creates a text box when no text edit is active.
      The gallery also browses Unicode equivalents for the five bundled legacy font mappings.
      Browsing arbitrary installed fonts remains open.
- [x] **P1** PowerPoint View > Zoom sets a preset or a custom percentage and restores Fit.
      The dialog uses the editor's existing fit-relative zoom scale.
- [x] **P1** PowerPoint Insert > 3D Models accepts a local self-contained GLB 2.0 with a poster
      fallback. Insert > Object embeds an existing DOCX, XLSX or PPTX as a file icon.
- [ ] **P1** PowerPoint Insert still lacks stock 3D assets, other 3D formats, desktop OLE activation,
      Create New, link-to-file, live previews of newly inserted objects.
- [x] **P1** Insert Microsoft Forms response links in PowerPoint and Excel, and open the response
      form on demand during a PowerPoint slide show.
- [ ] **P1** Native embedded Forms objects and signed-in Forms management in PowerPoint, and
      Microsoft 365 workbook response synchronization in Excel.
- [x] **P1** In Word, Excel and PowerPoint, Get Add-ins installs from the host catalog and opens
      Microsoft Marketplace for browsing. My Add-ins opens installed tabs and removes catalog
      add-ins when enabled by the host.
- [x] **P1** Word, Excel and PowerPoint install add-in-only XML manifests, unified JSON manifests
      and .zip app packages after a consent step, and keep them in the suite's and demos' profile
      preferences (session memory elsewhere). When the host turns the runtime on, enforces a frame
      policy and approves publisher origins, supported custom ribbon commands open sandboxed task
      panes; Word and Excel get a text selection bridge. Hosts can resolve Marketplace links or
      asset ids to packages; the Store stays hidden unless a host links a marketplace.
- [x] **P1** Home: the header Comments button toggles a docked Comments task pane (the slide
      keeps at least 320px), Drawing > Arrange is a vertical menu with Order, Group, Position, Align
      and Rotate, and the Selection Pane docks as a task pane, in all five bindings (2026-11-27).
- [x] **P1** Get Add-ins opens the shared Office Add-ins dialog in Word, Excel, PowerPoint and
      Visio: My Add-ins, admin-managed entries and Upload My Add-in with a consent step. The Store
      pivot appears only when the host links a marketplace (2026-11-27).
- [ ] **P1** Still open: the CDN Office.js script (the bridge is a custom protocol), requirement
      sets beyond Selection and TextCoercion 1.1, a PowerPoint selection bridge, ExecuteFunction
      (function commands), commands placed on built-in Office tabs, identity and SSO, content
      add-ins, and full XSD and JSON Schema validation (see `docs/office-addins-plan.md`).
- [x] **P1** Video Format, Audio Format, Playback and SmartArt Format tabs (2026-11-27).
- [ ] **P2** SmartArt Format's shape and style galleries are disabled, with a reason, for diagrams
      without editable drawing shapes.
- [ ] **P1** Picture Format: Compress, Change Picture, Reset Picture (and size), Picture Border,
      Alt Text and Arrange work in all five bindings (`e2e/pptx/picture-format-commands.spec.ts`);
      Picture Layout converts the selected pictures into one of the built-in picture SmartArt
      layouts in the shared SmartArt picker, saved with the built-in layout definition and the
      pictures in its placeholders (`picture-layout.ts`, `picture-layout-dialog.ts`, lazy with the
      layout library, 2026-11-27). Open: formats other than PNG and JPEG for Change Picture,
      Compress and Reset; Picture Layout has no live hover preview on the slide and picks the
      layout in a dialog rather than a ribbon dropdown.
- [x] **P1** SmartArt Design: Text Pane, Promote, Demote, Right to Left, Move Up, Move Down and
      Convert (to shapes) are enabled and wired across all five bindings
      (`ribbon-galleries/smartart-commands.ts`, `smartart-layout.ts`, `smartart-commands.test.ts`
      and per-binding ribbon gallery regressions).
- [ ] **P2** Table Design lacks Shading, Borders, Effects and Draw Borders; Table Layout's Select,
      View Gridlines, Cell Margins, Table Size and Arrange are disabled.
- [x] **P2** Ribbon Display Options (Always show Ribbon, Show tabs only, Full-screen mode,
      Ctrl+F1, double click on a tab) and Options > General > "Collapse the ribbon
      automatically", in all five bindings on the shared `createRibbonDisplayController`.
- [x] **P2** Options > General > "Collapse the search box by default" collapses the title-bar
      search box to an expandable icon toggle (`office-ui-title-bar`, `buildTitleBarState`) in
      all five bindings and updates live when toggled (`title-bar.test.ts` and per-binding tests).
- [x] **P2** Ribbon Display Options includes the fourth menu entry (Show/Hide Quick Access
      Toolbar, synchronized with Options > Quick Access Toolbar in all six locales) and dismisses
      the peeked ribbon after executing a command as well as on outside press or Escape, in all
      five bindings (`src/ui/src/pptx/render/ribbon-display.ts`, `ribbon-display.test.ts` and
      per-binding regressions).
- [ ] **P2** Home: Align Text and Convert to SmartArt; Review: Thesaurus, Translate, Mark All as
      Read, Restrict Permission, Hide Ink; View: Macros is a placeholder.

## Visio

The target is current Microsoft 365 Visio (ribbon, tabs and behaviour, decided 2026-10-30); the
installed Visio 16 build is used to check that saved files open without repair.

### File format and model

- [x] **P0** Formula dependencies the evaluator cannot follow no longer refuse a recalculation
      (`recalculateVisioCells`, `follow` mode of `edit-recalculate-index.ts`): a computed `INDIRECT`
      and `EVALCELL` are resolved while evaluating (cycle-protected, treated as reached by any edit
      on their page), `DEPENDSON` names its dependencies, a function the evaluator cannot run
      (unknown functions, `SetAtRef`, `Font`, `ThemeVal`, `ContainerSheetRef`...) keeps the cached
      value and is reported as `FORMULA_STALE`, and cells of grouped, instanced or connected shapes
      are read as inputs when the edit does not reach them. Still refused: a formula that does
      not parse, drawn `Polyline`/`NURBS` data (a stale outline would be drawn), an edit that itself
      reaches a grouped or inherited cell, a read of a cell that only the master defines
      (`F="Inh"` or no local cell), and an `INDIRECT` target that does not exist.
- [ ] **P1** The static admission guards (`analysis.dynamic` in `edit-geometry-admission.ts`,
      `edit-style-admission.ts`, `edit-scope.ts`, `edit-page-size.ts` and others) still refuse an
      edit of a cell whose own formula is dynamic, and kept caches are not recomputed: Visio would
      recompute them, so they can be stale until it recalculates. Master-defined cells are not read
      through master inheritance.
- [x] **P1** `.vsdm`, `.vstx` and `.vstm` open as drawings and `.vssm` for its masters; an edit
      keeps the VBA project byte for byte (never run), saves a template as a drawing and removes a
      standard signature with a notice (`package-kind.ts`, `package-kinds.test.ts`).
- [x] **P1** A signed drawing shows its verified signature state (signed, signer not verified,
      invalid, not verified) in the status bar, a load notice and File > Info, translated, and is
      never called signed unless the signature verifies
      (`src/ui/src/visio/signature-status.test.ts`). The viewer does not sign drawings.
- [ ] **P1** Save as template (`.vsdx` -> `.vstx` and `.vsdm` -> `.vstm`, `visio/save-template.ts`,
      `EditVsdxOptions.saveAsTemplate`) and stripping non-standard signature relationships on edit
      (`package-kind.ts`, `opc/signature`; `package-kinds.test.ts`) are implemented. Still open:
      native Visio check of an edited `.vsdm` (unadvertised VBA in non-macro packages stays refused
      for macro-smuggling safety).
- [ ] **P1** Legacy `.vsd` is preview only (no groups, masters, styles, connections, layers or
      text formats; no conversion or editing); `.vdx` is not supported; `.vss` and `.vst` are
      refused.
- [x] **P1** New Stencil, Add to My Shapes, Edit Stencil (add, rename, delete with confirmation,
      undo) and Save Stencil write `.vssx` packages with master ShapeSheets, connection points
      and icons that Visio 16 opens for editing (`src/core/visio/stencil-editor.ts`, 2026-11-26).
- [ ] **P1** Stencil master editing is limited to default plain text, adding a selected instance
      as a master drops its local overrides, and built-in masters lack the native ShapeSheet
      behaviours (smart tags, actions, formulas) of Visio's own stencils.
- [x] **P1** Controls, Actions, User-defined and Scratch cells and tab stops are in the model
      (`controls`, `actions`, `userCells`, `scratch`, paragraph `tabs`). Control handles are drawn
      yellow on a selected shape and dragging one writes `Controls.Row_N.X/Y` as Visio does and
      recomputes the dependent geometry (matched to Visio 16 for the Basic Shapes Can,
      Parallelogram, Trapezoid, Chevron, Cube and Cone). Actions list in the shape context menu and
      run when they are `SETF(GetRef(cell), value)` on the shape's own cells. Still open: User and
      Scratch cells have no editor; actions that call other functions (RUNADDON, OPENFILE, DOCMD)
      are shown disabled with the reason; a Menu text is refreshed after a run only when it is an
      `IF`/`&` of strings; handles locked by XCon/YCon or driven by a formula such as
      `Width*0.5+User.Offset` cannot be moved; geometry updates on release, not during the drag;
      tab stops read the `Position<n>`/`Alignment<n>` cells that Visio 16 saves; the alignment
      codes (0 left, 1 centre, 2 right, 3 decimal), the 0.5 in default steps and the stop origin
      were measured against Visio 16's SVG export of injected Tabs rows, and the layout (left,
      centred and right-aligned lines; the PDF export too) and a Text dialog Tabs tab that writes
      the stops match them. Still open: justified, distributed and right-to-left lines with tabs,
      stops relative to a paragraph's left indent, `tp` rows above 0 (not seen in any Visio file)
      and what Visio writes for a single stop (it ignored the alignment of a one-pair row we
      injected, so the editor writes two pairs).
- [ ] **P1** Containers and lists: containers now write `msvStructureType`, `msvSDContainer*` rows and
      `DEPENDSON` relationships (container kind 4, member back reference kind 1), but those kinds are
      not yet confirmed in a Visio build, and lists are not modelled (`edit-container-members.ts`).
- [x] **P1** Data linking stores a data connection (`connections.xml`), the command and the primary key;
      refresh matches rows by the key column and reports shapes whose row vanished
      (`data-connections.ts`, `data-refresh.ts`). Visio does not yet refresh from the
      source itself: the connection names the file, not its folder.
- [x] **P1** Insert > Pictures and drops admit static SVG (an allowlist: no scripts, events or
      external references), BMP, TIFF (decoded for display) and bounded classic EMF
      (`visio/edit-picture-formats.test.ts`, `e2e/visio/picture-format-insertion.spec.ts`, the
      security corpus); desktop Visio 16 reopens the BMP, TIFF and EMF pictures.
- [x] **P1** Embedded OLE objects (Insert > Object in Visio: other Office files, packages, binary
      OLE containers) are drawn from the picture saved with them (an `image` relationship of the
      embedded part: EMF through the metafile path, PNG, JPEG, GIF, BMP or TIFF), as an icon or
      as content, whichever was stored, and every edit keeps the embedding and its preview byte for
      byte. An embedded Word, Excel, PowerPoint or Visio file opens on double-click or the shape
      menu's Open Object: the cancelable `file-command` event `openEmbeddedDocument` lets a host
      open it (the Office suite opens it in its editor), else it downloads. Binary OLE containers,
      macro-enabled documents, non-ZIP payloads and linked objects are refused with the reason
      (`visio/ole-object.ts`, `ole-object.test.ts`, security corpus `visio-ole-objects.ts`).
- [ ] **P1** OLE objects remainders: no Insert > Object in Visio, no in-place activation or
      Convert, WMF previews are reported and not drawn, an object without a saved preview is not
      drawn, edits to an opened embedded document do not flow back into the drawing, and desktop
      Visio 16 shows an inserted SVG blank (also with the documented fallback picture). EMF
      conversion covers classic vector records only, BMP uncompressed and bitfield scanlines, and
      TIFF single-image strips.
- [ ] **P1** Bevel and 3-D rotation are approximations, not Visio's renderer. Perspective
      foreshortening (camera distance fitted to Visio 16 renders), extrusion (flat-shaded side
      faces) and the contour are drawn (`solid-3d.ts`, `render-solid.ts`), and a themed top bevel
      or contour is read from the effect style's `a:sp3d` (`theme-solid.ts`). The 3-D Rotation
      gallery holds the Office preset cameras (values read from PowerPoint over COM, signs fitted
      to Visio 16 renders; `rotation-presets.ts`), bevels are a height profile per type lit by
      the shape's `BevelLightingType` (15 types fitted to Visio exports) and `BevelMaterialType`
      (flat and powder), text and pictures under perspective follow one affine fit of the
      projection, and the children of a turned group turn with it (the same fit;
      `bevel-shading.ts`, `render-bevel.ts`). Still open: the profiles are shaped by eye and
      corners are rounded where Visio mitres them, strength varies with zoom, the oblique
      gallery presets (no rotation cells) and Visio's own gallery values were not read, material 8
      and the lighting colour tints are not drawn, `BevelLightingAngle` changed nothing in the
      export and is not read, side-face shading is fitted to a few renders, a bottom bevel is read
      and reported (Visio 16 shows no change on the front), perspective text and group children are
      not projected point by point, and the shipped Visio themes carry no `a:bevelT`, so the
      themed bevel path has no real-file evidence. Text has no Visio font metrics.
- [x] **P2** Every built-in arrowhead style 1 to 45 (style 13 in the sample drawings) is drawn
      from glyphs and setbacks measured in Visio 16 SVG exports (`filled-arrow.ts`,
      `open-arrow.ts`); custom `USE()` line ends are still reported, and off a straight
      round-capped line a glyph is drawn without its setback (`src/ui/src/visio/arrowheads.ts`).
- [ ] **P2** Shape text is measured and wrapped with browser fonts, so a block can run past its
      saved text-box height where Visio fits it; the viewer reports it
      (`src/ui/src/visio/render-text.ts`).
- [x] **P2** `InfiniteLine` rows are drawn across the page (`infinite-line.ts`).
- [x] **P2** Periodic NURBS are evaluated and drawn as closed loops (`nurbs.ts`, `nurbs-flatten.ts`;
      degrees 1-8, rational weights, seam continuity and affine knot normalization verified against
      a Cox-de Boor oracle in `visio/nurbs-periodic.test.ts`).
- [x] **P2** Stencil previews draw `InfiniteLine` using the existing page clipper with each
      master's page bounds, including transformed groups. Missing page dimensions use a local-box
      thumbnail fallback. Evidence: `visio/master-infinite-line.test.ts`, the synthetic Visio 16
      fixture generated by `scripts/record-visio-infinite-line-masters.ps1`, and all six browser
      bindings in `e2e/visio/stencil-infinite-line.spec.ts` (2026-10-31).
- [ ] **P2** Visio validation rule-set parts (`validation.xml`) are not modelled; Check Diagram is
      the viewer's own.

### Editing and behaviour

- [x] **P0** Containers do nothing after insertion: members move with the container, drops and
      moves change membership, and it grows to fit (`edit-container-fit.ts`). Open: swimlanes, lists
      and a native Visio check of the written relationships (see the P1 item above).
- [x] **P1** File > New offers Blank Drawing, Basic Flowchart, Cross-Functional Flowchart and Basic
      Organization Chart starters (`ui/templates.ts`); Visio 16 opens all three without repair.
- [x] **P1** Cross-functional flowcharts: lanes stack and resize together, insert before/after,
      delete with their shapes, separators, title and separator bars, Style 1-4 and rotated
      labels, with the contextual Cross-Functional Flowchart tab (`edit-swimlane-*.ts`,
      `ribbon-swimlane.ts`); Visio 16 opens edited files without repair.
- [x] **P1** Cross-functional flowcharts carry Visio's structure (frame container, Swimlane List,
      Phase List, `Relationships` kinds 1/4/2/5/11 confirmed through COM), Direction turns an
      existing chart, separators have phase labels, the page grows with the frame and deleting a
      lane is one undo step (`edit-swimlane-*.ts`, `ui/swimlane.ts`).
- [x] **P1** Lane edits (insert, delete, resize) work on a chart Visio built from its own template:
      the frame, lists and lanes are recognised by their masters, a new lane is a copy of its
      neighbour, and the cached values of the master instances are recomputed from the master's
      formulas (`cff-recalc.ts`, `edit-swimlane-clone.ts`). A vertical Cross-Functional Flowchart
      starter is on the New page, and shapes lying in a lane join its members on load, as Visio
      does (`swimlane-membership.ts`).
- [x] **P1** Charts the editor creates (File > New starters, Insert swimlane commands) are built
      from Visio's five hidden masters (CFF Container, Swimlane List, Phase List, Swimlane,
      Separator): the frame, lists, lanes and separators are instances of them, the page is
      related to the masters, and Visio 16 loads them as `CFF Container.50`, `Swimlane.10` and so
      on with the container, list and lane structure intact (`cff-author.ts`,
      `edit-swimlane-stamp.ts`). The masters are authored here: they carry Visio's names and
      structure rows but not its formulas, geometry or heading shapes, so the instances keep every
      cell they carry.
- [x] **P1** Phases and options of a horizontal chart Visio built: Insert Separator splits the
      phase it lands in (the new separator first in the Phase List, edge to edge, only the first
      phase extended over the lane headings), and Style, Show Title Bar, Show Separator Bar and
      Rotate Lane Labels write Visio's cells (`visCFFStyle`, `visShowTitle`, `visShowPhase` on the
      frame, `visRotateLabel` on the Swimlane List) and recalculate every instance from the
      masters' formulas. Checked cell by cell against files Visio 16 saved after the same changes
      (`cross-functional-native-*.vsdx`, `edit-swimlane-cff.ts`, `cff-recalc.ts`); charts the
      editor creates take the same commands, with the flowchart rows as plain numbers.
- [ ] **P1** Direction of a chart Visio built is refused (Visio swaps in "Swimlane (vertical)"
      and "Separator (vertical)" masters that are not copied here), as are phases, title bar and
      phase headings of a vertical chart Visio built (no Visio-made vertical chart to confirm the
      list layout). Hiding the phase headings or title bar was confirmed against Visio's formula
      cascade only, not its ribbon command, and measured-text caches (`visHeadingHeight` of a
      heading) are not recomputed. Visio's own add-in (Insert Swimlane, drop handling) acting on
      a chart the editor built is unverified: its Direction command opens a dialog that hangs an
      automated Visio.
- [x] **P1** Org chart positions carry Visio's org chart Shape Data (Department, Telephone, Name,
      Title, E-Mail), the Org Chart contextual tab adds subordinates, coworkers and assistants,
      deletes a position and lays the tree out (`ui/org-chart.ts`), and every New tile shows a
      thumbnail rendered from its template; Visio 16 reads the data and glue back.
- [x] **P1** Org chart positions are instances of Visio's Executive, Manager, Position and
      Assistant Belt masters joined by instances of the Organization Chart stencil's Dynamic
      connector (masters copied from ORGCH_M.VSTX with their UniqueIDs, Shape Data, user cells and
      photo; Visio 16 opens the files and reads the masters, Shape Data and glue); assistants sit
      beside the manager's line; Org Chart > Import builds a chart from a CSV or Excel sheet (Name,
      Title, Reports To; optional ID, Department, Telephone, E-Mail, Type); the text of a position
      is its Name and Title and follows them when the Shape Data changes; Delete on a position is
      Delete Position in one undo step (`ui/org-chart*.ts`, `viewer-org-chart.ts`).
- [x] **P2** Org chart types and glue: a Position Belt that gains its first subordinate (Org Chart > Subordinate) becomes a Manager Belt (the position is replaced by an instance of the Manager
      Belt master with its text, Shape Data and line), reporting lines are glued to the manager's
      Bottom and the subordinate's Top connection points and follow a moved position, and a Belt
      master instance can be resized (its text-height formulas keep the saved text extent and the
      text size pass measures it afterwards).
- [ ] **P2** Org chart remainders: a Manager Belt that loses its last subordinate does not turn
      back into a Position Belt, and a shape glued under a Position Belt does not promote it
      (neither could be observed through COM, where Visio's org chart add-in does not run; the
      Belt masters only show `User.ShapeType` and `msvReplaceClass`), no import from a
      directory, the Organization Chart Wizard's options or Compare/Update, an assistant's line
      still glues to the whole shape (it enters the assistant's side), the masters' theme-driven
      colours and line patterns are reported as warnings, and a page that already holds connectors
      cannot be grown for an imported chart.
- [x] **P1** Group resize scales nested groups and rotated or flipped members as Visio does
      (`W' = W(cos^2 sx + sin^2 sy)`, checked against Visio 16 in
      `edit-group-resize-native.test.ts`); multi-selection resize takes groups, rotated shapes,
      layered shapes on unlocked layers and shapes with glued connectors.
- [x] **P1** Connectors glued to group members follow the group, and group resize takes stencil
      (master) members and members on unlocked layers: glue reads a nested target through its
      group chain (PinX/PinY, LocPin, Angle, flips), so moving, resizing or turning the group
      reroutes the connector (ends checked against Visio 16 in `edit-group-glue-native.test.ts`);
      stencil members are written as local proportional formulas with refreshed `Inh` caches
      (`edit-group-stencil-resize-native.test.ts`); a multi-selection may hold stencil instances
      and connectors whose two ends are glued inside the selection (they reroute with the shapes).
- [ ] **P1** Still refused: a group member that is a connector or holds a connector end of its own,
      a group glued directly (its own PinX glue), a group holding an instance of a group master,
      stencil members whose pose cells carry formulas other than the group's own size, a selected
      connector with a free end or an end glued outside the selection, and a locked layer on any
      member. Not modelled: Visio also slides a dynamic end along a turned shape's side to line up
      with the other end (seen at 1.2 rad), and a selection resize of stencil instances is proved
      by the caller after the fact, not measured against Visio.
- [x] **P1** Dynamic connectors in files saved by Visio follow moved shapes, re-glue and reroute,
      including AutoConnect's PinY glue; end positions match Visio's through COM
      (`edit-connector-visio-native.test.ts`).
- [x] **P1** A copy that meets another dynamic connector is admitted with Visio's line jump,
      free-ended right-angle connectors reroute around a moved shape, and straight, curved,
      elbow and grouped connector masters copy and drop, checked against Visio 16 cells
      (`visio/__fixtures__/native-connector-copy-drop.json`).
- [ ] **P1** Visio's redistribution of parallel or coincident Dynamic connector copies, and its
      free-ended routing around complex obstacle layouts.
- [ ] **P1** A shape glued to a connector of a custom connector master (a 1-D master not built
      like Visio's Dynamic connector, as the sample generators write) cannot be moved: the
      connector cannot be rerouted, so `move-shape` refuses rather than leave it behind
      (`edit-instance-geometry.ts`, `edit-stencil-connector.ts` `proveStencilConnector`).
      Shapes glued with connectors this editor or Visio's Dynamic connector draws move.
- [x] **P1** Corner handles keep the proportions and Shift keeps them on every handle
      (`src/core/visio/ui/shape-resize.ts`; ledger entry of 2026-10-30).
- [x] **P1** Resize snaps the moving edge to other shapes' edges and centres and to shown guides,
      with Dynamic Grid alignment lines, outlines a shape whose width or height the frame matches,
      and Ctrl resizes about the centre (`src/core/visio/ui/resize-snap.ts`; ledger 2026-10-30).
- [x] **P1** Resizing a shape with `LockWidth`, `LockHeight` or `LockAspect` (local or
      style-inherited) constrains drags along the unlocked dimension or proportionally for
      `LockAspect` (`shape-resize.ts`, `resize-snap.ts`, `size-position.ts`), and core admission
      evaluates effective resize locks per dimension and aspect ratio rather than refusing
      unrelated axes (`edit-geometry-admission.ts`, `edit-geometry.ts`, `edit-resize-anchor.ts`,
      `edit-resize-locks.test.ts`).
- [ ] **P2** Ctrl and Alt during a resize are not verified against Visio's gestures (COM cannot
      drive the pointer); the size-match outline is our drawing, not Visio's.
- [x] **P1** Link Data (opens the External Data window, or imports first), Data Graphics Position,
      Configure and Auto Space (`data-graphic-config.ts`).
- [x] **P1** Data Visualizer (Data tab: Basic and Cross-Functional Flowchart from a local CSV or
      workbook with Process Step ID, Description, Next Step ID, Connector Label, Shape Type and
      Function columns, Refresh Diagram from the edited table), Data Graphic Fields (a pane of the
      selected shape's fields that show as text) and Advanced Data Graphics (the Edit Data Graphic
      item list: add, edit, delete) (`data-visualizer*.ts`, `data-graphic-items.ts`; ledger 2026-10-30).
- [x] **P1** Data graphics of a local shape are Visio data graphics: the shape becomes a group,
      each text callout, data bar or icon set an instance of a callout master that reads the
      shape's Shape Data by formula (`Sheet.N!Prop.Row`), the graphic a hidden data graphic master
      named by `User.visDGDisplayFormat`. Visio 16 reads the shape's `DataGraphic` back by that
      master's name, flags the callouts with `IsDataGraphicCallout` and recalculates them when a
      Shape Data cell changes through COM; our editor rewrites their caches on every Shape Data
      or data refresh edit and keeps their IDs. Position, Auto Space, Configure, the item list,
      resize, copy, delete and formatting work on such a shape (`data-graphic-masters.ts`,
      `edit-data-graphic*.ts`; ledger 2026-10-31, wave 6).
- [ ] **P1** Data graphics still fall back to separate marked shapes for an instance of a stencil
      master (so every Data Visualizer shape) and for ordinary groups: Visio's own structure there
      is a group holding an inheriting copy of the shape plus the callouts, not written yet. The
      callout masters are this package's own minimal ones with Visio's names (Text callout, Data
      bar 1, Status icons), not the DVCALL stencil's, so Visio's gallery of icon sets, data bar
      styles and callout shapes is not offered; the `msvDGPosition` drag handle, the Data
      Graphic layer and the document's last-applied graphic are not written; Visio's own Edit
      Data Graphic dialog has not been driven on our file. Data Visualizer does not connect to
      Visio's online template service. Closed 2026-10-31: Data Visualizer connector labels are
      the connector's own text, and a refresh keeps the shape and connector IDs of rows that
      remain and is one undo step.
- [x] **P1** PDF export writes vector paths, arrowheads, pictures and selectable text per page,
      with a per-page report of what degraded (`src/ui/src/export/pdf-vector.ts`,
      `visio/viewer-pdf.ts`).
- [x] **P1** PDF Export has a page range (all, current page, from/to over the foreground pages),
      shows the per-page report on the Export page, draws linear and radial gradients as PDF
      shadings, shadows as hard-edged offset copies, glow as a half-strength halo, and embeds
      host-supplied TrueType fonts as subset Type0/CIDFontType2 fonts with a ToUnicode CMap,
      and every embedded font, ToUnicode and page stream is Flate-compressed (`/FlateDecode`,
      `/Length1` the uncompressed size)
      (`visio-viewer.pdfFonts`; `core/text/truetype`, `export/pdf-embedded-font.ts`,
      `export/pdf-vector.ts`, `visio/viewer-pdf-options.ts`, `export/pdf-svg-gradient.ts`).
- [ ] **P1** PDF: no font is bundled or looked up, so text outside WinAnsi still becomes a
      picture page unless the host sets `pdfFonts` with a font that has the glyphs (CFF/OpenType
      fonts and fonts whose licence forbids embedding are refused and reported); shadow and glow
      blur, soft edges and reflections are not drawn; gradient transparency is averaged; no
      pattern fills, no PDF/A and no XPS.
- [x] **P2** Page tabs drag to reorder, background pages only among themselves
      (`chrome/tab-drag.ts`, shared with Excel's sheet tabs; `e2e/visio/page-tab-drag.spec.ts`).
- [ ] **P2** No ShapeSheet window or Developer tab; on-shape text editing is a plain text area.
- [ ] **P2** Stencil drops refuse picture masters and multi-root masters with rotated shapes or
      top-level connectors (`src/core/visio/ui/master-drop.ts`).

### Ribbon

- [x] **P1** The compact phone Quick Access Toolbar keeps Save, Undo and Redo reachable.
- [x] **P1** The contextual Shape Format tab exposes solid WordArt Text Fill using the same
      native text color edit as Home. Shape Styles, Arrange and Size are shared with Home.
- [x] **P1** Visio Shape Format offers a solid Text Outline and Text Effects for glyph shadow
      and glow. VSDX saves these controls in ShapeSheet User cells and the viewer draws them.
- [ ] **P1** Visio glyph outline, shadow and glow have no matching native Microsoft Visio
      character style. Gradient text fill, text deformation and additional effects are open.
- [x] **P1** Picture Format, Container Format and Chart Design contextual tabs replace Shape Format
      for a lone picture, container or chart; Chart Design opens the chart editor
      (`ribbon-contextual-formats.ts`).
- [ ] **P1** Picture Compression and Auto Balance are disabled (the editor does not re-encode or
      histogram a picture); the de/es/fr/zh-CN strings of the Picture and Container Format tabs are
      new and need a native review. A terminology pass (2026-10-31) aligned the command names with
      Microsoft's localised terms and the terms already used in each locale; it is not a native review.
- [x] **P1** Design > More Variants: Colors, Fonts, Effects and Connectors apply one scheme each and
      mark the current one; Visio 16 opens the result and keeps the four scheme cells.
- [ ] **P1** Embellishment writes the page's EmbellishmentIndex (Low, Medium, High as 1, 2, 3,
      the `embellishment` values of the theme's variants); Visio 16 resolves no shape cell from it,
      so there is nothing to draw, but what its Embellishment gallery does to the variant index is
      unverified. Themed arrowheads are read from the connector line style (Visio 16's own themes
      carry none). Visio's own schemes other than our six built-in themes cannot be picked; Visio
      does not name our scheme IDs.
- [x] **P1** Insert > Get Add-ins opens the shared Office Add-ins dialog; the Marketplace handoff
      is available when the host enables it (2026-11-27).
- [ ] **P1** Local add-in manifests cannot be installed in Visio; the controls say why.
- [ ] **P2** Disabled: CAD Drawing, Object, Thesaurus, Ink, Import Rules From,
      SharePoint Workflow, the Window group, Macros and Add-Ons.

## Review history

- 2026-10-30: first edition, from a fan-out code review of all four products and the shared
  layer. The stale claims it found were corrected in the product trackers in the same change.
- 2026-10-31: recorded completed work on signature cleanup, PowerPoint automation saves and
  equation ownership, Excel selection persistence, x14 validations, threaded comment editing and
  preservation, and FILTERXML. Added native and browser evidence for Word comment preservation,
  comment whitespace, solid theme backgrounds and Visio infinite-line stencil previews. Kept
  remaining work explicit, including new mention authoring, untranslated function descriptions,
  exact Word tint rounding, rich comments, non-solid backgrounds and periodic NURBS.

### Screen capture and slide creation

- [x] Word, Excel and PowerPoint insert a screenshot from the browser's shared-surface picker.
      Screen Clipping accepts a dragged rectangle or pixel coordinates and saves the result as
      an embedded picture.
- [x] PowerPoint Screen Recording can pause, resume and stop, optionally capture audio shared by
      the browser, and embed the resulting video. Select Area records a dragged rectangle or
      one entered as coordinates on the shared surface. Record Pointer passes a preference to
      the browser's capture picker. An in-browser dock replaces the setup dialog during recording.
      It shows elapsed recording time, collapses when unpinned and idle on desktop, and has a
      keyboard- and touch-accessible handle to reveal the recording controls.
- [x] PowerPoint Photo Album downloads a standalone presentation with a title slide, 1, 2 or 4
      pictures per slide, configurable picture order, and optional filename captions.
- [x] PowerPoint Reuse Slides inserts selected slides from a local OOXML presentation and keeps
      their source formats and related parts.
- [ ] Desktop window thumbnails, selecting an area directly on the desktop, reliable pointer
      control on every browser surface, a recording dock that follows other apps, and in-place
      Photo Album editing are open. Reuse Slides still needs destination-theme formatting and import
      between different slide sizes or OOXML conformance classes.

### Insert Icons

- [x] PowerPoint, Word, Excel and Visio offer a bundled, searchable gallery with color selection
      and batch insertion. Selection persists while searching.
- [x] PowerPoint, Word and Excel save icons as SVG with a PNG fallback. Visio saves native
      vector paths that resize and can be recolored with the existing Line control.
- [x] **P1** Graphics Fill recolors selected SVGs in PowerPoint, Word and Excel using a DrawingML
      color replacement effect. Visio uses its native Line control on inserted vector icons.
- [ ] **P1** The Microsoft 365 stock catalog, PowerPoint custom Graphics Fill colors, and themed
      gradients or additional graphics-specific formatting are open.

### Symbol-font fallback (#290)

- [x] PowerPoint's common Wingdings bullets and mapped Symbol/Wingdings text have
      Unicode substitution and an offline bundled symbol font fallback in the
      shared renderer. Evidence: `symbol-bullet-paragraph.test.ts`,
      `symbol-font-fallback.spec.ts`.
- [x] Word, Excel, PowerPoint and Visio browse and insert Unicode equivalents for Symbol,
      Wingdings, Wingdings 2, Wingdings 3 and Webdings. The dingbat mappings use the published
      WG2 N4384 source references. Word exposes these through Insert > Symbol > More Symbols.
- [ ] Exact symbol-font outlines, legacy symbols without a portable character, arbitrary
      installed-font browsing and a bundled font for supplementary-plane characters remain
      open. PowerPoint avoids expanding runs to two UTF-16 units when rendering stored text to
      keep editing offsets valid.
