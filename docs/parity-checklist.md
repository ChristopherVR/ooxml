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

| Capability    | Word         | Excel        | PowerPoint                 | Visio                 |
| ------------- | ------------ | ------------ | -------------------------- | --------------------- |
| Chart render  | 2-D and 3-D  | 2-D and 3-D  | yes, own renderer with 3-D | 2-D and 3-D, embedded |
| Chart editing | yes          | yes          | yes                        | partial               |
| ChartEx       | partial      | partial      | partial                    | no                    |
| SmartArt      | render, edit | render, edit | render, edit               | no (by design)        |
| Equations     | yes          | no           | yes                        | no                    |
| Proofing      | browser only | none         | browser only               | browser only          |

- [x] **P0** 3-D charts draw flat in Word, Excel and Visio: `src/core/chart/render/chart-3d-*.ts`
      now projects `c:view3D` (oblique for right-angle axes, perspective otherwise) and paints
      bar3D, line3D, area3D, pie3D and surface3D with depth ordering and face shading; PowerPoint's
      projection and shading code is shared with it.
- [ ] **P1** 3-D chart detail in the core painter: data labels, floor and wall formatting, bar
      `c:shape` (cylinder, cone, pyramid), pie explosion, surface wireframe and contour modes, and
      measurement against Office renders.
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

- [x] **P0** AutoFilter drop-down in the grid: Sort, Clear Filter, Filter by Color, Text, Number
      and Date Filters (Custom AutoFilter and Top 10 dialogs, Above/Below Average, date periods),
      search and the value checklist. Value, date-group, custom, Top 10, dynamic and colour
      criteria are read, written and evaluated (`src/core/xlsx/filter-criteria.ts`).
- [ ] **P1** AutoFilter remainders: drop-downs on table headers (tables keep value lists only),
      the date checklist as a year/month/day tree, icon filters in the menu (icon criteria in a
      file are evaluated through conditional formats), colour list entries that come only from
      conditional formats.
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
      through master inheritance. Container `DEPENDSON`/`SheetRef` relationships are still not
      written.
- [x] **P1** `.vsdm`, `.vstx` and `.vstm` open as drawings and `.vssm` for its masters; an edit
      keeps the VBA project byte for byte (never run), saves a template as a drawing and removes a
      standard signature with a notice (`package-kind.ts`, `package-kinds.test.ts`).
- [ ] **P1** No Save as template; a signature relationship outside the standard layout, or a VBA
      project in a package that does not declare macros, still refuses an edit. No native check of
      an edited `.vsdm` yet.
- [ ] **P1** Legacy `.vsd` is preview only (no groups, masters, styles, connections, layers or
      text formats; no conversion or editing); `.vdx` is not supported; `.vss` and `.vst` are
      refused.
- [ ] **P1** No `.vssx` or `.vstx` writer; built-in stencils are outlines without master
      ShapeSheets, connection points or Shape Data.
- [ ] **P1** Controls (control handles), Actions, User-defined and Scratch cells, and tab stops are
      not in the model, so they cannot be shown or edited (they survive saving).
- [ ] **P1** Containers and lists: containers now write `msvStructureType`, `msvSDContainer*` rows and
      `DEPENDSON` relationships (container kind 4, member back reference kind 1), but those kinds are
      not yet confirmed in a Visio build, and lists are not modelled (`edit-container-members.ts`).
- [x] **P1** Data linking stores a data connection (`connections.xml`), the command and the primary key;
      refresh matches rows by the key column and reports shapes whose row vanished
      (`data-connections.ts`, `data-refresh.ts`). Visio does not yet refresh from the
      source itself: the connection names the file, not its folder.
- [ ] **P1** OLE objects are not supported; EMF admission is an allowlist; SVG, BMP and TIFF
      pictures are refused on insert.
- [ ] **P1** Bevel, 3-D rotation and full theme effects are not drawn; text has no Visio font
      metrics.
- [x] **P2** `InfiniteLine` rows are drawn across the page (`infinite-line.ts`).
- [ ] **P2** Periodic NURBS are missing; stencil previews still omit `InfiniteLine`.
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
- [ ] **P1** Charts the editor creates still use plain shapes with direct cells, not Visio's hidden
      CFF master instances (so Visio's add-in that inserts lanes and handles drops does not act on
      them); phases, options (title, style, labels) and Direction of a chart Visio built are
      refused.
- [x] **P1** Org chart positions carry Visio's org chart Shape Data (Department, Telephone, Name,
      Title, E-Mail), the Org Chart contextual tab adds subordinates, coworkers and assistants,
      deletes a position and lays the tree out (`ui/org-chart.ts`), and every New tile shows a
      thumbnail rendered from its template; Visio 16 reads the data and glue back.
- [ ] **P1** Org chart positions are plain rectangles, not Visio's Executive, Manager, Position and
      Assistant Belt masters, so Visio's own Org Chart tab does not recognise them; assistants are
      placed as subordinates; no import from a file or directory; the text on a position does not
      follow its Shape Data.
- [x] **P1** Group resize scales nested groups and rotated or flipped members as Visio does
      (`W' = W(cos^2 sx + sin^2 sy)`, checked against Visio 16 in
      `edit-group-resize-native.test.ts`); multi-selection resize takes groups, rotated shapes,
      layered shapes on unlocked layers and shapes with glued connectors.
- [ ] **P1** Still refused: resizing or moving a group whose members have glued connectors
      (connector glue reaches top-level shapes only), a group with stencil (master) or layered
      members, and a multi-selection that holds stencil instances or connectors.
- [x] **P1** Dynamic connectors in files saved by Visio follow moved shapes, re-glue and reroute,
      including AutoConnect's PinY glue; end positions match Visio's through COM
      (`edit-connector-visio-native.test.ts`).
- [ ] **P1** A copy that meets another dynamic connector is refused; Visio also reroutes unglued
      connectors whose old path crossed a moved shape, the core does not; non-Dynamic and
      grouped connector masters are refused.
- [x] **P1** Corner handles keep the proportions and Shift keeps them on every handle
      (`src/core/visio/ui/shape-resize.ts`; ledger entry of 2026-10-30).
- [x] **P1** Resize snaps the moving edge to other shapes' edges and centres and to shown guides,
      with Dynamic Grid alignment lines, outlines a shape whose width or height the frame matches,
      and Ctrl resizes about the centre (`src/core/visio/ui/resize-snap.ts`; ledger 2026-10-30).
- [ ] **P2** Ctrl and Alt during a resize are not verified against Visio's gestures (COM cannot
      drive the pointer); the size-match outline is our drawing, not Visio's.
- [x] **P1** Link Data (opens the External Data window, or imports first), Data Graphics Position,
      Configure and Auto Space (`data-graphic-config.ts`).
- [ ] **P1** Data Visualizer (a diagram generated from a table), Data Graphic Fields and the
      Advanced Data Graphics editor are not built; graphics remain separate marked shapes, not
      Visio data graphic masters.
- [x] **P1** PDF export writes vector paths, arrowheads, pictures and selectable text per page,
      with a per-page report of what degraded (`src/ui/src/export/pdf-vector.ts`,
      `visio/viewer-pdf.ts`).
- [x] **P1** PDF Export has a page range (all, current page, from/to over the foreground pages),
      shows the per-page report on the Export page, draws linear and radial gradients as PDF
      shadings, shadows as hard-edged offset copies, glow as a half-strength halo, and embeds
      host-supplied TrueType fonts as subset Type0/CIDFontType2 fonts with a ToUnicode CMap
      (`visio-viewer.pdfFonts`; `core/text/truetype`, `export/pdf-embedded-font.ts`,
      `visio/viewer-pdf-options.ts`, `visio/pdf-svg-gradient.ts`).
- [ ] **P1** PDF: no font is bundled or looked up, so text outside WinAnsi still becomes a
      picture page unless the host sets `pdfFonts` with a font that has the glyphs (CFF/OpenType
      fonts and fonts whose licence forbids embedding are refused and reported); shadow and glow
      blur, soft edges and reflections are not drawn; gradient transparency is averaged; no
      pattern fills, no PDF/A and no XPS.
- [x] **P2** Page tabs drag to reorder, background pages only among themselves
      (`chrome/tab-drag.ts`, shared with Excel's sheet tabs; `e2e/visio/page-tab-drag.spec.ts`).
- [ ] **P2** No ShapeSheet window or Developer tab; on-shape text editing is a plain text area.
- [ ] **P2** Stencil drops refuse non-Dynamic connector, picture and multi-root rotated masters
      (`src/core/visio/ui/master-drop.ts`).

### Ribbon

- [x] **P1** A contextual Shape Format tab shows while shapes are selected (Shape Styles, WordArt
      Styles disabled, Arrange, Size), reusing the Home commands (`ribbon-shape-format.ts`).
- [x] **P1** Picture Format, Container Format and Chart Design contextual tabs replace Shape Format
      for a lone picture, container or chart; Chart Design opens the chart editor
      (`ribbon-contextual-formats.ts`).
- [ ] **P1** Picture Compression and Auto Balance are disabled (the editor does not re-encode or
      histogram a picture); the de/es/fr/zh-CN strings of the Picture and Container Format tabs are
      new and need a native review.
- [x] **P1** Design > More Variants: Colors, Fonts, Effects and Connectors apply one scheme each and
      mark the current one; Visio 16 opens the result and keeps the four scheme cells.
- [ ] **P1** Embellishment writes the page's EmbellishmentIndex (Low, Medium, High) but its
      decorative geometry is not drawn, and the level-to-index mapping is not verified against
      Visio's gallery; theme connector arrowheads are not themed from the connector scheme; Visio's
      own schemes other than our six built-in themes cannot be picked; Visio does not name our
      scheme IDs.
- [ ] **P2** Disabled: Bevel, 3-D Rotation, CAD Drawing, Object, Thesaurus, Ink, Import Rules From,
      SharePoint Workflow, the Window group, Macros and Add-Ons.

## Review history

- 2026-10-30: first edition, from a fan-out code review of all four products and the shared
  layer. The stale claims it found were corrected in the product trackers in the same change.
