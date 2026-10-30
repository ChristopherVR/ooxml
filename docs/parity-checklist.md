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

| Capability    | Word         | Excel        | PowerPoint                 | Visio              |
| ------------- | ------------ | ------------ | -------------------------- | ------------------ |
| Chart render  | 2-D only     | 2-D only     | yes, own renderer with 3-D | 2-D only, embedded |
| Chart editing | yes          | yes          | yes                        | partial            |
| ChartEx       | partial      | partial      | partial                    | no                 |
| SmartArt      | render, edit | render, edit | render, edit               | no (by design)     |
| Equations     | yes          | no           | yes                        | no                 |
| Proofing      | browser only | none         | browser only               | browser only       |

- [ ] **P0** 3-D charts draw flat in Word, Excel and Visio: the core renderer has no `view3D`,
      perspective or rotation; the 3-D meshes live only in `src/ui/src/pptx/render`
      (`chart-3d-bar-mesh.ts`). Move a projection into `src/core/chart/render`.
- [ ] **P1** Region map ChartEx (`cx:geography`) is modelled and saved but not drawn by
      `src/core/chart/render`; the Pareto cumulative line is unconfirmed there.
- [ ] **P1** Visio has no ChartEx path (`src/ui/src/visio/render-embedded-chart.ts`).
- [ ] **P1** Pivot charts: `pivotSource` and `pivotFmts` are only carried through; no field
      buttons and no link to PivotTable edits.
- [ ] **P1** Chart Design lacks Add Chart Element and Select Data in Word, Excel and PowerPoint,
      plus Switch Row/Column in Word and Excel and Change Chart Type in PowerPoint.
- [ ] **P1** Equations in Excel text boxes and in Visio (`ooxml-core/math` and
      `src/ui/src/docx/equation-*` are the template).
- [ ] **P1** SmartArt layouts the preview marks approximate (`src/core/diagram/layout-preview.ts`),
      3-D styles and bent hierarchy connectors; the built-in catalogue has 176 layouts
      (`diagram/builtin/catalog.ts`), not yet checked against Office's gallery.
- [ ] **P1** Proofing is the browser's `spellcheck` attribute everywhere: no grammar, thesaurus,
      custom dictionary, ignore list or proofing-language engine.
- [ ] **P1** Picture Format and Shape Format contextual tabs: missing in Excel (both) and Word
      (Picture Format); partial in Word and PowerPoint Shape Format (no Arrange, Size or Alt Text).
- [ ] **P2** `c:userShapes` (shapes over a chart) is parsed but not drawn; surface charts have no
      dedicated painter.
- [ ] **P2** Possible duplicate equation code in `src/ui/src/pptx/render` (`latex-*`, `omml-*`):
      confirm each file is a facade over `ooxml-core/math` and delete real copies.

## Excel

### File format and preservation

- [x] **P1** Every save writes `fullCalcOnLoad="1"` and `calcId="191029"`
      (`src/core/xlsx/write/workbook-part.ts`), so Excel recalculates each saved file and asks to
      save on close. Write it only when cached values may be stale.
      Closed: the source calcPr attributes are kept (calcId included), and `fullCalcOnLoad="1"` is
      written only for a formula without a cached result or after an edit that was not
      recalculated (`Workbook.valuesStale`, set by `edit/session-calc.ts`).
- [ ] **P1** `.xlsb` is detected and refused (`load/detect.ts`); `.ods` is not read or written.
- [ ] **P1** x14 data validation (lists over another sheet) is kept in `extLst`, not modelled or
      editable.
- [ ] **P1** Rich values and in-cell pictures (`xl/richData`, IMAGE() results, Place in Cell,
      linked data types) have no model.
- [ ] **P1** Connections, query tables and Power Query are carried verbatim only; external links
      use cached values and cannot be refreshed or edited.
- [ ] **P1** Pivots outside the bounded profile are read-only: label and value filters, shared
      caches, calculated fields and items, date-group filters (`pivot/admission.ts`).
- [x] **P2** Phonetic runs (`rPh`) are dropped from shared strings (`read/shared-strings.ts`).
      Closed: `rPh` runs and `phoneticPr` are read onto the cell (`Cell.phonetic`) and written back
      with the shared string. The guide is not displayed, edited or carried by copy and paste or
      collaboration, and `phoneticPr` `fontId` is written as read.
- [ ] **P2** Named sheet views and ActiveX controls are not modelled.
- [ ] **P2** Collaboration does not share filters, pivots, panes and views, page setup,
      protection, images and drawings or document properties.

### Calculation

- [ ] **P2** INFO, ASC, DBCS, JIS, PHONETIC and EUROCONVERT are not implemented (no host service
      needed).
- [ ] **P2** CUBE\* (7), WEBSERVICE, FILTERXML, STOCKHISTORY, RTD, IMAGE, COPILOT, TRANSLATE and
      DETECTLANGUAGE are not implemented (need an OLAP, web or AI service; document them as
      unsupported).

### Editing and behaviour

- [ ] **P0** No AutoFilter drop-down in the grid: no value checklist, search, Text, Number or Date
      filters, Top 10 or Filter by Colour (`src/core/xlsx/ui/commands/sort-filter.ts`; the model's
      `AutoFilterColumn` holds a value list only). Custom, Top 10, dynamic, colour and icon
      criteria in a file are neither evaluated nor authored.
- [ ] **P1** Sort has no Sort On cell colour, font colour or icon, custom lists, case sensitivity
      or left-to-right sorting (`src/ui/src/xlsx/dialogs/sort-dialog.ts`).
- [ ] **P1** Formula auditing: Trace Precedents, Trace Dependents, Remove Arrows, Evaluate Formula,
      Error Checking and Watch Window.
- [ ] **P1** Flash Fill, Subtotal, Consolidate and Advanced filter.
- [ ] **P1** Split panes (only Freeze Panes exists).
- [x] **P1** Text to Columns refuses fixed-width splitting (`dialogs/text-to-columns.ts`).
      Closed: the dialog offers Fixed width with typed break positions and a preview; the split
      rules are `splitFixedWidth` in `src/core/xlsx/edit/text-columns.ts`. Column-data-format
      conversion per column (Excel's Column data format step) is not offered.
- [ ] **P1** Changing the workbook theme: Page Layout > Themes > Colors only reports that it is
      not supported (`src/core/xlsx/ui/commands/page-layout.ts`); no Themes, Fonts or Effects.
- [ ] **P1** Form controls cannot be moved or resized and buttons run no macros; shapes have no
      effects, 3-D, pattern or picture fills, and gradients are linear only.
- [ ] **P1** Printing goes through the browser; Page Layout view has no copy and paste, per-view
      zoom or native printer metrics.
- [ ] **P2** Editing threaded comment replies.
- [ ] **P2** Fill-handle drag autofill: confirm it exists (`grid/grid-pointer.ts`) with a browser
      test; only Fill > Series was found.
- [ ] **P2** Grid accessibility has not been audited (`grid/grid-a11y.ts`).

### Ribbon

- [ ] **P1** Missing tabs: Draw, Picture Format, Shape Format, Chart Format (a pane instead),
      PivotTable Analyze and PivotTable Design (pivot commands sit in Insert > PivotTable).
- [ ] **P1** Insert: Shapes, Icons, 3D Models, Screenshot, PivotChart, Recommended PivotTables,
      Equation, New Note, and the Text group (Text Box, Header & Footer, WordArt, Object).
- [ ] **P1** Page Layout: Scale to Fit, Print Titles, Background and the Arrange group.
- [ ] **P1** Data: Get Data and its sources, Queries & Connections, Refresh All, Edit Links.
- [ ] **P2** Review: Spelling, Thesaurus, Translate, Check Accessibility, Notes; View: Window group
      and Macros; Help: Help, Training, Feedback; Backstage: Account, Close.

## Word

### File format and preservation

- [ ] **P0** Legacy `.doc` opens as main-body text only: tables, headers, footers, images, fields
      and formatting are dropped, and saving writes plain paragraphs
      (`src/core/docx/load/legacy-doc.ts`).
- [ ] **P1** Styles: no new styles; the paragraph, character and table style catalogues and the
      theme are refused on save; latent styles are not modelled.
- [ ] **P1** Table-level properties (width, alignment, indent, borders, style, look, grid widths)
      are not editable.
- [ ] **P1** Table-structure revisions (`cellIns`, `cellDel`, `cellMerge`, `tblPrChange`,
      `trPrChange`, `tcPrChange`) are not modelled and block structure edits; `tblGridChange` and
      `numberingChange` have no code.
- [ ] **P1** Fields computed today: PAGE, NUMPAGES, SECTIONPAGES, DATE, TIME, TOC, SEQ, REF,
      PAGEREF, STYLEREF. IF, `=`, MERGEFIELD, INCLUDETEXT, INCLUDEPICTURE, DOCPROPERTY, AUTHOR,
      INDEX/XE, TA/TOA, CITATION/BIBLIOGRAPHY, LISTNUM and SYMBOL keep their cached result
      (`layout/page-fields.ts`); Update Fields skips the `\n \r \w \p` switches.
- [ ] **P1** Grouped drawings (`wpg`) and canvases (`wpc`), legacy VML text boxes (`v:textbox`),
      WordArt, ink and gradient or picture fills are kept but not drawn.
- [ ] **P1** Ruby (`w:ruby`) has no model and probably blocks editing its paragraph; vertical
      text (`tbRl`, `btLr`), `eastAsianLayout`, `fitText` and emphasis marks are not laid out.
- [ ] **P1** Picture bullets (`numPicBullet`) are not drawn or written.
- [ ] **P1** Protection: forms protection is enforced; read-only, comments-only and
      tracked-changes-only are kept but not enforced, and there is no Restrict Editing UI.
- [ ] **P1** Digital signatures: `opc/signature` is not used by docx, so signed documents are not
      detected before an edit.
- [ ] **P1** OLE objects and ActiveX controls are kept and drawn as placeholders only.
- [ ] **P2** Changing the comment list rewrites `comments.xml` from the model, losing comment
      formatting the model does not hold; `commentsIds` and `w16cid` are not maintained.
- [ ] **P2** Embedded fonts are kept, not used for rendering; `altChunk`, art page borders and
      theme page backgrounds are kept, not drawn.

### Editing and behaviour

- [ ] **P1** Mail merge: no data source or MERGEFIELD engine.
- [ ] **P1** Find & Replace is literal with case sensitivity only: no wildcards, whole word,
      formatting search or special characters (`src/ui/src/docx/search-controller.ts`).
- [ ] **P1** Only Draft and Print Layout views: no Read Mode, Web Layout, Outline or Focus.
- [ ] **P1** Index, citations and bibliography, Table of Authorities.
- [ ] **P1** Compare and Combine; accessibility checker.
- [ ] **P1** Paste options (Keep Source, Merge, Text Only) and Paste Special.
- [ ] **P1** No AutoCorrect or AutoFormat as you type (smart quotes, automatic lists).
- [ ] **P1** Collaboration covers the body and comments; headers, footers and notes are
      incomplete, table commands are off during collaboration, and there are no mentions.
- [ ] **P2** Customize Ribbon can show and hide, not reorder; Options are not persisted.

### Layout and rendering

- [ ] **P1** Hyphenation: `autoHyphenation` is read but never applied by `src/core/docx/layout`.
- [ ] **P1** Tight and through wrapping use the object's rectangle, not its wrap polygon, and only
      the larger side gets text (`layout/floats.ts`).
- [ ] **P1** Table autofit (`tblLayout`) is not applied; merged cells are read-only on Print Layout
      pages.
- [ ] **P1** Floating objects inside table cells, floating charts and floating text boxes are not
      positioned on pages; linked text boxes have no code.
- [ ] **P1** Pagination is approximate (no Word-rendered page corpus yet); caret placement at line
      and page boundaries is heuristic; no kerning, East Asian line breaking or complex-script
      shaping.

### Ribbon

- [ ] **P1** Mailings: all 16 commands are disabled (`ribbon-native-tabs.ts`).
- [ ] **P1** Design: Themes, Colors, Fonts, Paragraph Spacing, Effects and Set as Default are
      disabled.
- [ ] **P1** Table tab (Table Design and Layout merged) lacks table styles and style options, Cell
      Size and AutoFit, cell alignment and margins, Repeat Header Rows, Formula, Convert to Text,
      Select and Draw Borders.
- [ ] **P1** Missing tabs: Draw, Picture Format, Equation.
- [ ] **P1** Layout: the Arrange group (Position, Wrap Text, order, Selection Pane, Align, Group,
      Rotate) and Text Direction.
- [ ] **P1** Insert: WordArt, Quick Parts, Signature Line, Object, Icons, 3D Models, Screenshot,
      Video; Header and Footer offer blank designs only.
- [ ] **P2** View: Window group, Macros; Review: Thesaurus, Translate, Block Authors, Hide Ink;
      Home: Text Effects, Phonetic Guide, Character Border, Enclosed Characters; Backstage:
      Share, Account, Close.

## PowerPoint

### File format and preservation

- [ ] **P1** 40 coverage-manifest constructs are `preserve: unassessed` and 30 `edit: unassessed`
      (`src/core/pptx/core/openxml/openxml-coverage*.ts`): add a dirty-save round trip with unknown
      elements on every part type, and ratchet the counts in CI.
- [ ] **P1** Chart and ChartEx save is not native-certified for fresh templates (a six-slide deck
      failed to open in PowerPoint).
- [ ] **P1** `.ppt` export turns video and non-WAV audio into pictures.
- [ ] **P1** Clipboard is a suite slide-package format with no desktop Office interop.
- [ ] **P2** 3-D models: `am3d:preTrans` is parsed, not applied; `rot` handedness and
      orthographic scale are unverified.
- [ ] **P2** `p:smartTags` has no model; `p:revisionInfo` survival is unverified.

### Editing and slide show

- [ ] **P1** Remove Background: an existing `a14:backgroundRemoval` is drawn but cannot be
      authored (`src/ui/src/pptx/render/image-background-removal.ts`).
- [ ] **P1** Zoom: slide, section and summary zooms render and save but cannot be inserted.
- [ ] **P1** Recording: no narration or Cameo; Clear and Reset of recorded timings are disabled
      (`ribbon-record-commands.ts`); Record Audio has no command.
- [ ] **P1** Animation: Split, Zoom, Blinds and the shape reveals keep their default direction;
      Color Typewriter plays a substitute without a warning; emphasis effects use preset
      keyframes; Blink is approximate.
- [ ] **P1** Animation Painter and the ribbon's Start and Duration fields are disabled
      placeholders (timing is edited in the Animation Pane).
- [ ] **P1** Ink: Lasso, pen presets, Ink to Shape, Ink to Math, Ruler and Ink Replay (Pen,
      Highlighter, Eraser and Freeform work).
- [ ] **P1** Custom shows load and save but have no add or edit API, tool or UI.
- [ ] **P2** Notes Master and Handout Master are placeholders; no Notes Page view; no Grayscale or
      Black and White view.
- [ ] **P2** Designer, Rehearse with Coach, Screen Recording, Keep Slides Updated and Media
      Controls are disabled.

### Ribbon

- [ ] **P1** Insert: Symbol, WordArt, Object, Icons, 3D Models, Screenshot, Photo Album, Reuse
      Slides, Forms and add-ins are disabled, plus ten menu entries without a command
      (`ribbon-insert-layout.ts`).
- [ ] **P1** Missing tabs: Video Format, Audio Format, Playback and SmartArt Format.
- [ ] **P1** Picture Format lacks Remove Background, Compress, Change Picture, Reset Picture,
      Picture Border, Picture Layout, Alt Text and Arrange.
- [ ] **P1** SmartArt Design: Text Pane, Promote, Demote, Right to Left, Move Up, Move Down and
      Convert are disabled (`ribbon-galleries/smartart-commands.ts`).
- [ ] **P2** Table Design lacks Shading, Borders, Effects and Draw Borders; Table Layout's Select,
      View Gridlines, Cell Margins, Table Size and Arrange are disabled.
- [ ] **P2** Home: Align Text and Convert to SmartArt; Review: Thesaurus, Translate, Mark All as
      Read, Restrict Permission, Hide Ink; View: Zoom and Macros are placeholders.

## Visio

### File format and model

- [ ] **P0** Formula dependencies the evaluator cannot follow refuse the edit instead of
      recalculating: `INDIRECT` is always dynamic, and `EVALCELL`, unknown functions,
      `DEPENDSON`/`SheetRef` and inherited or grouped reads refuse
      (`src/core/visio/formula-analyze.ts`, `EDIT_*_DEPENDENCY`). This is the root of most edit
      refusals below.
- [ ] **P1** `.vsdm`, `.vstx` and `.vssm` cannot be opened as drawings; macro-enabled or signed
      packages refuse every edit (`parser.ts`, `edit-package.ts`).
- [ ] **P1** Legacy `.vsd` is preview only (no groups, masters, styles, connections, layers or
      text formats; no conversion or editing); `.vdx` is not supported; `.vss` and `.vst` are
      refused.
- [ ] **P1** No `.vssx` or `.vstx` writer; built-in stencils are outlines without master
      ShapeSheets, connection points or Shape Data.
- [ ] **P1** Controls (control handles), Actions, User-defined and Scratch cells, and tab stops are
      not in the model, so they cannot be shown or edited (they survive saving).
- [ ] **P1** Containers and lists: membership is a private User row; `DEPENDSON`/`SheetRef`
      relationships and `msvStructureType` rows are not written, so Visio may not treat the result
      as a container.
- [ ] **P1** Data linking stores no data connection; refresh matches rows by order, not by key.
- [ ] **P1** OLE objects are not supported; EMF admission is an allowlist; SVG, BMP and TIFF
      pictures are refused on insert.
- [ ] **P1** Bevel, 3-D rotation and full theme effects are not drawn; text has no Visio font
      metrics.
- [ ] **P2** `InfiniteLine` geometry is not rendered; periodic NURBS are missing.
- [ ] **P2** Visio validation rule-set parts (`validation.xml`) are not modelled; Check Diagram is
      the viewer's own.

### Editing and behaviour

- [ ] **P0** Containers do nothing after insertion: members do not move with the container, it
      does not grow to fit, and drops do not change membership (`edit-diagram-parts.ts`).
- [ ] **P1** Swimlanes and cross-functional flowcharts, org charts and a New-drawing template
      gallery (New offers only a blank drawing).
- [ ] **P1** Group edits exclude nested groups, rotated and stencil members, flips and glue;
      multi-selection resize excludes groups, rotation, layers and glue.
- [ ] **P1** Connectors in files saved by Visio are refused for geometry edits; a copy that meets
      another dynamic connector is refused.
- [ ] **P1** Resize ignores Shift, Ctrl and Alt (aspect, centre-out); no resize alignment or
      size-match hints.
- [ ] **P1** Link Data and Data Visualizer; Data Graphics Position, Configure and Auto Space.
- [ ] **P1** PDF export is a picture per page with no selectable text; no XPS.
- [ ] **P2** No ShapeSheet window or Developer tab; page tabs cannot be dragged to reorder (a
      Reorder Pages dialog exists); on-shape text editing is a plain text area.
- [ ] **P2** Stencil drops refuse non-Dynamic connector, picture and multi-root rotated masters
      (`src/core/visio/ui/master-drop.ts`).

### Ribbon

- [ ] **P1** No contextual tabs at all (Shape Format, Picture Format, Container, chart); chart
      editing has no ribbon button.
- [ ] **P1** Design > More Variants: Colors, Fonts, Effects, Connectors and Embellishment are
      disabled (`ribbon-themes.ts`).
- [ ] **P2** Disabled: Bevel, 3-D Rotation, CAD Drawing, Object, Thesaurus, Ink, Import Rules From,
      SharePoint Workflow, the Window group, Macros and Add-Ons.

## Review history

- 2026-10-30: first edition, from a fan-out code review of all four products and the shared
  layer. The stale claims it found were corrected in the product trackers in the same change.
