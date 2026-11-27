# Sample files

Twelve Word, Excel, PowerPoint and Visio files that exercise what the editors in this repository
can do, shown in the gallery at
[christophervr.github.io/ooxml/samples](https://christophervr.github.io/ooxml/samples/). Each card
there opens the file in its editor on the same site or downloads it.

Every file is written by a script under [`generators/`](generators), so it can be rebuilt byte
for byte and reviewed as source. The generators use `ooxml-core`'s own models, writers and
editing operations where they exist, and write the remaining OOXML parts by hand (per ECMA-376
and the [MS-VSDX] open specification) where the library has no writer yet.

| Folder          | What it holds                                                              |
| --------------- | -------------------------------------------------------------------------- |
| `docx/`         | The Word documents                                                         |
| `xlsx/`         | The Excel workbooks                                                        |
| `pptx/`         | The PowerPoint presentations                                               |
| `visio/`        | The Visio drawings                                                         |
| `thumbnails/`   | One 800 x 450 WebP picture per sample, for the gallery card                |
| `generators/`   | The scripts: one folder per product, shared helpers in `lib/`, `build.ts`  |
| `manifest.json` | Title, description, product, features, size, hash and thumbnail per sample |

## The samples

### Word

**[Technical report](docx/technical-report.docx)** (203 KB). A long, fully structured report on a tidal energy pilot: numbered headings, a table of contents, lists of figures and tables, captions, cross-references, notes, citations, equations, charts and tracked review.

- Cover page and five sections (roman and arabic page numbers, a landscape page)
- Styles with multilevel heading numbering
- Table of contents, list of figures and list of tables
- Numbered captions and cross-references (SEQ, REF, PAGEREF)
- Footnotes and endnotes
- Citations and a bibliography
- Display and inline equations
- Tables with header rows, banding and merged cells
- Native charts with embedded workbooks, and pictures
- Headers and footers with a different first page
- Comments with a reply and tracked changes

**[Community newsletter](docx/community-newsletter.docx)** (148 KB). A designed two-page newsletter for a community garden: a masthead, columns that change from one to two to three, a drop cap, a pull quote, wrapped pictures, a sidebar and an events calendar, all coloured from a custom theme.

- Custom theme: colours and fonts used through theme references
- Continuous section breaks switching between one, two and three columns
- Masthead and pull quote in DrawingML text boxes
- Pictures with square and tight wrapping at anchored positions
- Drop cap
- Sidebar with shading and borders
- Events calendar table with a banded table style
- Designed header and footer with page numbers

**[Services agreement form](docx/services-agreement-form.docx)** (22 KB). A fillable services agreement and client intake form for a design co-operative: every kind of content control, legacy form fields, a fee schedule with a calculated total, numbered legal clauses and a signature block.

- Content controls: plain and rich text, drop-down list, combo box, date picker, check boxes, picture, building block gallery
- Repeating section for fee lines and a locked group around the terms
- Titles, tags and placeholder text on every control
- Fields: DATE, DOCPROPERTY, FORMTEXT, FORMCHECKBOX, =SUM(ABOVE), REF, HYPERLINK, PAGE and NUMPAGES
- Custom document properties read by fields
- Numbered legal clauses with multilevel numbering
- Two-column section for the terms
- Signature block table

### Excel

**[Five-year financial model](xlsx/financial-model.xlsx)** (42 KB). A driver-based plan for Brindlemoor Bakehouse: validated assumptions with three scenarios, a revenue and cost build, and an income statement, balance sheet and cash flow that tie, summarised on a dashboard.

- Eight linked sheets with cross-sheet formulas and cached results
- Workbook and sheet scoped named ranges
- Data validation: lists, whole-number and decimal ranges, input messages, error alerts
- Conditional formatting: data bars, icon sets, colour scales and formula rules
- Custom number formats: thousands, parentheses for negatives, multiples, year labels
- Outlined rows and columns with summary rows, two levels deep
- Frozen panes, restyled cell styles and a custom theme
- Native charts: stacked columns with a line on a secondary axis, and a line chart
- Sparklines, Scenario Manager scenarios, internal and external hyperlinks

**[Sales analysis](xlsx/sales-analysis.xlsx)** (128 KB). A year of orders for Kestrelline Outfitters as an Excel table with calculated columns, explored through a PivotTable with slicers and a timeline, and summarised on a dashboard of charts and sparklines.

- Excel table of 960 rows with calculated columns, structured references and a totals row
- Lookup tables read with XLOOKUP
- PivotTable with row, column and value fields, a report filter and its cache
- Two slicers and a timeline connected to the PivotTable
- Combination chart: columns with a line on a secondary axis
- Doughnut, bar and stacked area charts
- Line and column sparklines with markers
- Conditional formatting: data bars, colour scales, icon sets, top and bottom rules
- A designed dashboard on a custom theme

**[Operations tracker](xlsx/operations-tracker.xlsx)** (339 KB). A maintenance tracker for Vale Orrin Logistics: a log of 4000 work orders, a live overview that spills from dynamic array formulas, a printable depot report, notes and threaded comments, and protected sheets with unlocked inputs.

- Log of 4,000 rows with an AutoFilter and a stored sort state
- Dynamic arrays that spill: FILTER, SORT, SORTBY, UNIQUE, XLOOKUP, LET, SEQUENCE, TAKE, DROP, HSTACK, TEXTJOIN
- Spill references (A1#) in totals and COUNTIFS, and a chart over spilled ranges
- Legacy notes and threaded comments with replies and people
- Sheet protection without a password, with unlocked input cells
- Print setup: print area, repeated titles, landscape, scaling, headers and footers, page breaks
- Conditional formatting for overdue and priority rows, data bars, icon sets
- Data validation: list, date and custom rules
- Internal, web and mail hyperlinks; an Excel table; named ranges

### PowerPoint

**[Product launch keynote](pptx/product-launch-keynote.pptx)** (547 KB). The launch keynote for Aerie, the home weather station from Nimbuswright: thirteen dark, cinematic slides that Morph a glowing sensor ring from the cover into the product reveal, then prove the claim with a chart, a specification table and a live storm warning.

- Custom dark theme, slide master and eight layouts
- Morph transitions with paired shapes, plus fade, push, wipe, split, cover and zoom
- Entrance, emphasis and exit animations with with-previous and after-previous timing
- Text built paragraph by paragraph
- Motion path animation
- Animation triggered by clicking a shape
- SmartArt process diagram
- Column and doughnut charts with embedded workbooks
- Table with designed formatting
- Gradient fills, soft shadows and custom geometry
- Generated pictures with transparency
- Speaker notes on every slide and named sections

**[How Port Averly moved](pptx/data-story-ridership.pptx)** (196 KB). Port Averly Transit's 2025 ridership review, told as a data story: six native charts with annotations, a network map and a plan for next year, on a light editorial theme.

- Custom theme, slide master and six layouts
- Native column, line, stacked bar, doughnut, combo and scatter charts with embedded workbooks
- Morph, fade, push, wipe, split and cover transitions
- Entrance, emphasis, exit and motion-path animations
- Text built paragraph by paragraph
- Animations triggered by clicking a shape
- SmartArt, a formatted table and a generated network map
- Gradient fills, soft shadows and custom geometry
- Speaker notes on every slide and named sections

**[Studio portfolio](pptx/studio-portfolio.pptx)** (327 KB). The portfolio of Thimblewick Studio, a small design studio: three case studies (a bakery identity, garden wayfinding and an outfitter’s field guide) told with large generated imagery, quiet typography and Morph transitions that carry each picture from thumbnail to full page.

- Custom theme, slide master and six layouts in Rubik and Noto Sans
- Morph transitions with named shape pairs, plus fade, push, wipe, cover and split
- Entrance, emphasis, exit and motion-path animations, with- and after-previous timing
- Text built paragraph by paragraph
- Interactive sequences triggered by clicking a picture
- SmartArt process diagram
- Doughnut chart with an embedded workbook
- Table with designed cell formatting
- Gradient fills, soft shadows and a custom-geometry studio mark
- Generated pictures with alternative text
- Speaker notes on every slide and named sections

### Visio

**[Order to delivery flowchart](visio/order-to-delivery-flowchart.vsdx)** (36 KB). A cross-functional flowchart of how Kestrelline Outfitters takes an order from checkout to the doorstep, with a second page for returns.

- Swimlanes with phases
- Masters and glued connectors
- Shape data linked to a recordset
- Text fields, layers and hyperlinks
- Background page with a title block

**[Platform architecture](visio/platform-architecture.vsdx)** (48 KB). The Nimbuswright platform behind the Aerie home weather stations: a logical view by tier and the network it is deployed on.

- Containers with members
- Group masters and glued connectors
- Layers per tier and per kind of flow
- Shape data linked to an inventory recordset
- Callouts, text fields and a title block

**[Institute organisation chart](visio/institute-org-chart.vsdx)** (33 KB). Who reports to whom at the Saltmarsh Institute for Coastal Engineering, with the timeline of its Greywharf tidal energy pilot on a second page.

- Organisation tree with glued connectors
- Group master with department colours
- Cards linked to a staff directory recordset
- Project timeline with milestones and dependencies
- Date and text fields from shape data

## Rules the samples follow

- **Fictitious content.** Every company, product, place and person is invented
  (`generators/lib/fiction.ts`). No real organisation, brand, logo or person appears.
- **Open fonts.** Rubik for headings, Noto Sans for text and Noto Serif for long-form reading.
  All three are served by the editors' optional webfont fallback (`docs/font-loading.md`). No
  Microsoft font is named anywhere in a sample.
- **Generated pictures.** Every picture is drawn by the script that needs it
  (`generators/lib/art.ts`, encoded with the core's PNG encoder). Nothing is downloaded.
- **Small.** Each file stays under 1.5 MB and all of them together under 15 MB.
- **Deterministic.** No clock, no randomness beyond a seeded generator, fixed ZIP dates. Running
  the build twice gives the same bytes.
- **Opens in Microsoft Office without repair.** Each file was opened in desktop Word, Excel,
  PowerPoint or Visio through the application's documented automation before it was committed.

## Rebuilding

```bash
bun samples/generators/build.ts                  # write every sample and manifest.json
bun samples/generators/build.ts --only xlsx      # one product (the manifest is left alone)
bun samples/generators/build.ts --sample <id>    # one file
bun samples/generators/build.ts --check          # compare the committed files with a fresh build
bun samples/generators/build.ts --determinism    # generate everything twice and compare
```

`scripts/samples.test.mjs` (part of `bun run test:scripts`) typechecks the generators against the
core sources they import (`samples/tsconfig.json`, also `bun run typecheck:samples`), runs every
generator twice and compares the bytes, and verifies that `manifest.json` describes exactly the
committed files. CI runs it for a change under `samples/` and for a change to a core area the
generators import, so a model or writer change that breaks a generator fails with that change.

It does not compare the committed files with a fresh build, so a writer change that only moves
bytes does not fail it. `--check` reports that difference on demand.

**Commit rebuilt files only when a sample really changes.** The files are binary: every rebuilt
copy stays in the repository's history at its full size. Rebuild and commit when a generator
changes what a sample contains, or when a core fix changes what a reader would see or what Office
accepts, and then open the rebuilt files in Office first. A difference `--check` reports that
changes nothing a reader sees (part order, an id, whitespace) is not a reason to commit them.

The thumbnails are pictures of what Office draws, so they are not part of the deterministic
build. `generators/thumbnails.ps1` renders them on Windows with desktop Office and Python
(PyMuPDF and Pillow); `generators/thumbnails.json` says which slide, page or sheet stands for
each sample. Rerun it when a sample's first slide, page or sheet changes.

## How the gallery uses them

`scripts/build-samples-page.mjs` renders `site/samples/template.html` with one card per manifest
entry and copies the files and thumbnails beside it; the suite build and the Pages build both call
it. "Open in Word" (or Excel, PowerPoint, Visio) links to the suite with `?sample=<id>`; the suite
then fetches the manifest and that one file from its own site and adds a copy to the library in
the browser (`apps/office-suite/src/samples.js`, disclosed in `PRIVACY.md`). The file must have
the size and SHA-256 the manifest records. The samples are never bundled into an editor, and the
desktop app is built without them (`bun run build:suite:desktop`).

## What the editors do not show yet

The samples use features on purpose that the editors only preserve, or draw differently from
Office. Those gaps are listed in [`docs/parity-checklist.md`](../docs/parity-checklist.md); a
sample keeps the feature so the gap stays visible. `e2e/a11y/samples-expected.json` names the
load notice each gap causes and the checklist item behind it, and that list may only shrink.
