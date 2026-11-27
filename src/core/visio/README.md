# Visio area (`ooxml-core/visio`)

The `visio` area owns all Visio document logic: reading VSDX packages (and the read-only legacy VSD preview) into a DOM-free typed scene, the bounded ShapeSheet formula interpreter, inheritance, geometry, themes and text, the source-backed `editVsdx` transaction, blank drawing and template creation, stencils, the shape clipboard format, the DOM-free view logic the Visio UI paints (`visio/ui`) and the Yjs mapping for collaboration (`visio/collab`).

It is not a separate npm package: it ships inside `ooxml-core` and is re-exported whole by [`visio-core`](../../../viewers/visio/packages/core/README.md). The `<visio-viewer>` element is in `ooxml-ui/visio` (`src/ui/src/visio`), and the framework bindings, docs site and demos are in [`viewers/visio`](../../../viewers/visio/README.md). Visio is **Beta**: this is a supported subset, not Visio parity or a complete ShapeSheet engine, and plain-text editing and the shape clipboard are experimental. The [capability ledger](../../../viewers/visio/docs/parity.md) is the canonical list of what is admitted and the evidence behind it; the [verification record](../../../viewers/visio/docs/verification.md) holds every native Visio check.

## Entry points

| Subpath                   | Holds                                                                                                                                                                                                                                                                                                                   |
| ------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ooxml-core/visio`        | `parseVsdx`, `loadVisio`, `parseLegacyVsd`, `parseVsdxPages`/`mergeVisioPages`, `editVsdx` and the `VisioEdit` union, `createVsdx`, `convertVsdxToTemplate`, `readVisioStencil`, `loadVisioStencilEditor` (lazy), the clipboard codec, formula interpreter, text measuring, themes, stencils and the metafile adapters. |
| `ooxml-core/visio/ui`     | DOM-free view logic every binding shares: the element contract (`propertyKeys`, `eventKeys`), customization, document history, selection, arrange, resize, draw and text-replace planners, Size & Position, formatting state, org chart, data graphics, templates.                                                      |
| `ooxml-core/visio/collab` | The Yjs mapping of a VSDX package (pages, shapes and connects as keyed entries, other parts verbatim) and `bindDrawingSession`; see [collaboration area](../../../docs/collab-area.md).                                                                                                                                 |

```ts
import { parseVsdx, getVisioPageLayers } from 'ooxml-core/visio';

const drawing = await parseVsdx(bytes);
const layers = getVisioPageLayers(drawing, drawing.pages[0]!.id);
```

Scene coordinates are inches, y-up, with angles in radians. A shape's six-number affine matrix maps its local coordinates into its parent's. Geometry is inert SVG path data, never source markup; apply the page y-axis inversion only at the rendering boundary, and invert text and raster images to match. Background page layers come in painting order. Missing or unsupported features are returned as diagnostics; unsafe or malformed packages throw `VisioPackageError` with a `code`.

## Module map

Files sit flat in the area, grouped by prefix (tests next to the code):

| Prefix                                                                                                         | Responsibility                                                                                                        |
| -------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| `package*`, `parts`, `xml-*`, `memory-budget`, `heap-limit`                                                    | Bounded ZIP and XML reading, part resolution, the memory budget.                                                      |
| `parser`, `load`, `parse-incremental`, `legacy`, `shapes`, `sheet`, `style*`                                   | The scene: pages, shapes, masters, inheritance, styles; the VSD preview.                                              |
| `geometry`, `nurbs*`, `spline-geometry`, `complex-geometry`, `rounded-geometry`, `infinite-line`               | Geometry rows to SVG paths, bounded curve approximation, corner rounding.                                             |
| `theme*`, `line-*`, `fill-pattern*`, `*-gradient*`, `source-paint`, `effects`, `shadow`, `bevel-*`, `solid-3d` | Paint: theme resolution, line patterns and caps, gradients, effects, 3-D.                                             |
| `text*`, `paragraphs`, `tab-stops`, `plain-text`, `webfont-runs`                                               | Text runs, paragraphs, fields, measuring.                                                                             |
| `layers`, `visibility`, `shape-metadata`, `data-*`, `comments`, `print-*`, `page-*`                            | Layers, visibility, Shape Data and hyperlinks, data record sets, comments, print and page settings.                   |
| `media`, `picture-*`, `raster-*`, `emf-*`, `foreign-vector*`, `convert-metafile`, `prepare-*`                  | Pictures, EMF admission, the foreign-vector sanitizer, the metafile adapter, embedded charts.                         |
| `formula*`, `cell-formula`, `control-formula`, `action-formula`                                                | The ShapeSheet formula parser, analyser and evaluator.                                                                |
| `edit.ts`, `edit-*`                                                                                            | The `editVsdx` transaction and every command family (one `edit-<family>` module each, `*-commands.ts` for the types). |
| `create-document`, `save-template`, `stencil*`, `clipboard*`                                                   | Blank drawings, templates, stencils and the stencil editor, the clipboard codec.                                      |
| `connector-*`, `line-jumps`, `snap-glue`, `swimlane-*`, `container-relationships`                              | Connector routing and jumps, glue, swimlanes and containers.                                                          |
| `ui/`, `collab/`                                                                                               | The `visio/ui` and `visio/collab` subpaths.                                                                           |

## Rendering model

What the scene resolves, in brief (the ledger has the full table):

- Pages with background chains; groups, rotation, flips and pins; master cells, sections and rows with child inheritance and deletion; style chains and document defaults.
- Move/line rows, relative Beziers, arcs, ellipses, literal `POLYLINE`, bounded positive-weight NURBS and cached spline sequences (with approximation diagnostics), and saved corner radius on axis-aligned rectangles and orthogonal line chains ([connector rounding evidence](../../../docs/visio-connector-rounding.md)).
- Cached and themed colours, fills, line widths, caps and patterns, Quick Style selectors, DrawingML theme colours and variants, linear gradients with up to 32 stops, arrowheads.
- Text runs, fonts, paragraphs, bullets, margins and text backdrops; layer membership and visibility; Shape Data and inert hyperlink metadata; embedded PNG, GIF and JPEG with validated signatures.

Conventions consumers rely on:

- Paragraph offsets are UTF-16 offsets into `plainText`, excluding newlines. Exact line spacing is in inches; multiple spacing scales the largest font on a line. Text layout and raster decoding are the viewer's job.
- `style.lineCap` maps directly to SVG `stroke-linecap`. `style.lineDash` holds dash/gap lengths in stroke-width multiples (multiply by the real `lineWidth`, no minimum floor); zero-length square-cap dots use `lineDashDotLength`. `visioLineDashLengths` (in `visio/ui`) applies both rules.
- `visioSvgStrokeStyle` (in `visio/ui`) applies the native SVG export policy: an exactly zero saved weight draws at 0.75 pt.
- Suppress line properties, arrowheads included, when geometry has `stroke: false` or the effective `linePattern` is 0.
- Only cached values, supported theme records and numeric literals are used; no formula runs as code, and no macro, add-on, hyperlink, external data or network resource is executed or fetched. Shape Data dates stay serial days.

Unsupported: theme effects and transforms beyond the resolved set, radial and non-horizontal saved gradients (solid fallback plus diagnostic), hatch rendering, layer colour overrides, nonliteral splines, most foreign and vector records, OLE objects, APNG, uncommon JPEG coding and arbitrary image clipping. Native visual equivalence is not established.

## Editing contract

`editVsdx(input, edits, options)` applies an atomic, source-backed batch of `VisioEdit` commands and returns `{ bytes, changedParts, diagnostics }`. It edits the source XML; it never serializes the normalized scene.

- **Atomic.** All commands are checked before anything is written; any refusal fails the whole call. Master and stencil drops run first, in order, then the other commands as one transaction, so a dropped shape can be styled, sized, typed into and connected in the same call.
- **Preservation.** Untouched part payloads keep their exact uncompressed bytes, unknown parts, relationships and content types included. An edited XML part keeps unrelated content semantically, not lexically; ZIP headers and compression are not preserved. Changed saves use STORE compression, so files may grow. A no-op returns an independent copy. UTF-16 source XML becomes UTF-8 only in edited parts.
- **Admission.** VSDX and VSDM drawings are edited (stencils through the stencil editor). Digital signatures are removed before an edit, as Visio does; a VBA project is carried byte for byte and never run, and one in a package that is not macro-enabled refuses the edit.
- **Protection and dependencies.** Effective locks (`LockMoveX`/`Y`, `LockWidth`/`Height`/`Aspect`, `LockDelete`, `LockTextEdit`, `LockReplace`), `GUARD`, redirection and locked layers are honoured, never bypassed. Affected formula caches are recomputed through the bounded interpreter or the edit is refused.
- **Ownership.** Calls own their input bytes and command values before the first asynchronous step; each call has its own bounded lifetime, and no mutable DOM is exposed. Run editing in a disposable worker for responsiveness.
- **Limits.** Package limits apply, plus 1,000 commands, one million replacement UTF-16 code units and an output ceiling equal to `maxInputBytes` per call. A drawing whose dependency index would exceed 100,000 ShapeSheet cells (about 3,800 shapes) is too large to edit: `parseVsdx` reports `edit-size-limit` first.

### Command families

| Family               | Commands                                                                                                                                                                                                                                                                      |
| -------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Text                 | `replace-plain-text`, `replace-text-ranges`, `insert-text-field`, `update-text-field`                                                                                                                                                                                         |
| Formatting and order | `format-text`, `format-shape`, `reorder-shape`                                                                                                                                                                                                                                |
| Geometry             | `create-rectangle`, `create-ellipse`, `create-line`, `create-path`, `create-text-box`, `move-shape`, `resize-shape`, `rotate-shape`, `flip-shape`, `move-line-endpoint`, `delete-shape`                                                                                       |
| Copy and structure   | `duplicate-shapes`, `paste-shapes`, `group-shapes`, `ungroup-shape`, `change-shape`, `assign-layers`, `set-layer-properties`                                                                                                                                                  |
| Masters and stencils | `insert-master-instance`, `drop-stencil-master`                                                                                                                                                                                                                               |
| Connectors           | `glue-connector`, `set-connector-route`, `add-connection-point`, `delete-connection-point`, `set-snap-glue`                                                                                                                                                                   |
| Diagram parts        | `insert-container`, `insert-callout`, `set-container-members`, `set-container-options`, `fit-container`, `reconcile-containers`, `create-swimlanes`, `insert-swimlane`, `delete-swimlane`, `resize-swimlane`, `insert-separator`, `set-swimlane-options`, `create-subprocess` |
| Pages                | `insert-page`, `delete-page`, `rename-page`, `reorder-page`, `set-page-size`, `set-page-setup`, `set-page-properties`, `set-page-layout`, `set-page-theme`, `set-page-decoration`, `set-page-alt-text`, `create-guide`, `move-guide`, `delete-guide`                          |
| Metadata and data    | `set-shape-data`, `set-shape-hyperlink`, `set-shape-screentip`, `set-shape-alt-text`, `import-data-recordset`, `refresh-data-recordset`, `delete-data-recordset`, `link-data-rows`, `unlink-data-rows`, `set-data-graphic`, `set-control`, `run-action`                       |
| Objects and comments | `insert-picture`, `crop-picture`, `set-picture-adjust`, `insert-chart-object`, `replace-chart-object`, `add-comment`, `edit-comment`, `delete-comment`                                                                                                                        |

The types are in `edit-*-commands.ts` and `edit.ts`. The rest of this section keeps the admission rules that are not written down anywhere else; the ledger lists the admitted set per family.

#### Blank drawings

`createVsdx({ width?, height?, stencils? })` writes a source-backed, editable one-page VSDX in physical inches (default Letter, 1:1 scale) with explicit Calibri 12 pt text and black-line, white-fill defaults, independent of installed templates. `stencils` writes Stencil windows for Visio's `BASIC`, `BASFLO` and `ARROWS` files. Invalid dimensions fail before package creation.

#### Text

- `replace-plain-text` edits one local plain `Text` element per command. Rich text, unknown markup and carriage returns are refused; one stored terminal paragraph marker is excluded, intentional trailing LF and blank paragraphs are kept. On a master instance (or a sub-shape of a group instance) it writes a plain local `Text`, as Visio 16 saves, marking inherited Field rows deleted when the master text has fields; an instance with its own formatted text accepts range edits only.
- Cells that read the text through `SHAPETEXT(TheText)` or a text function (`LEN`, `FIND`, `STRSAME` and so on) are recalculated with everything downstream; a bare `TheText` in arithmetic, text-valued cells and unevaluable formulas refuse the edit; a cell in a cycle keeps its cache with a `FORMULA_CYCLE` diagnostic. Master, document-sheet and theme cells are not yet in this graph.
- A shape that sizes itself from its text (`TEXTWIDTH(TheText)`, `TEXTHEIGHT(TheText, ...)`, the flowchart masters' Resize with Text) is measured again after text, font, text block and resize edits (`edit-text-size.ts`, `text-extent.ts`), from font advance widths recorded from Visio 16; `setVisioTextMeasurer` lets a UI measure other fonts. Unreliable cases (unknown fonts, bullets, tabs, a line that just fits) keep the saved size with `edit-text-size-kept`; a refreshed size reports `edit-text-size`.
- `insert-text-field` adds a `<fld>` and Field row (Type 0 string, 2 number, 5 date) at an offset or the end. It evaluates `NOW()`, the document date and property functions (`TITLE()`, `CREATOR()` and the rest), `PAGENAME()`, `PAGENUMBER()`, `PAGECOUNT()` and numeric formulas over the shape's own cells; anything else is refused. Format pictures are the `VISIO_FIELD_FORMATS` subset. `update-text-field` edits one live local Field row and its unique marker after comparing source text, cached field text and the row's `sourceKey`; field-row references, protected cells, nonstandard calendars, nested, grouped or inherited shapes and ambiguous markers are refused. Range edits treat a field as one atomic token.

#### Formatting and stacking order

- `format-text`: `fontFamily` (an existing document FaceName), `fontSize` (points), `bold`, `italic`, `underline`, `strikethrough`, `fontColor` (hex), `fontTransparency`, `textCase`, `textPosition`, `language` (an LCID), `letterSpacing`, `bullets`, `bulletStyle` (0-7), `bulletText`, `indentLeft`/`Right`/`First`, `spaceBefore`/`After`, `lineSpacing` (`multiple` writes a negative SpLine, `exact` points), `horizontalAlign`, `verticalAlign`, `margins`, `textBackground` (saved as Visio's `RGB()+1` formula), `textBackgroundTransparency`, `textBlock` (written as `Width*k`/`Height*k` formulas so it follows resizes) and `tabStops` (Visio 16's `Position<n>`/`Alignment<n>` pairs).
- `format-shape`: `fillColor` (hex or `none`), `lineColor`, `lineWeight` (points), `linePattern` (0-23), `fillPattern` (0-24), `fillBackgroundColor`, `lineTransparency` and `fillTransparency` (0-100, rounded to native half-percent steps). Pattern 0 hides the paint, 1 is solid; `fillColor: 'none'` with a nonzero pattern is rejected; active gradient changes need explicit paint replacement.
- `reorder-shape` takes `order: 'front' | 'back' | 'forward' | 'backward'` within display band zero.
- Formatting admits local leaf shapes and lines, preserving mixed-run markers and cached fields; proven local `F="Inh"` cells become explicit overrides; protected cells, error caches, unsupported inheritance and formulas depending on the changed formatting are refused. Pure theme formulas can be overridden; `THEMEGUARD` is distinct from `GUARD` and does not protect manual formatting.

#### Geometry

- Coordinates are drawing inches, bottom-left origin, at the rotation pin, bounded to one million inches; shape IDs are explicit positive integers. Existing top-level local 2-D shapes from any producer are admitted.
- `resize-shape` without `anchor` holds the pin. `anchor: { x, y }` (each `0`, `0.5` or `1`, local y up) fixes a local bounds point for local unlayered 2-D leaves with line geometry or canonical ellipses: rotation and flips are kept, unguarded local pins scale, guarded pins stay only when their values prove the same pose. Groups, foreign shapes, glued targets, affected text fields and dimension-dependent Angle/Flip formulas are refused.
- `create-text-box` uses the document's default text style with `FillPattern=0` and `LinePattern=0`; it models fixed boxes, not the Text Tool's automatic sizing.
- Local unglued straight lines admit `move-shape` (both endpoints move, midpoint and length formulas stay), Width-cell `resize-shape` with Height 0, `move-line-endpoint` (`endpoint: 'begin' | 'end'`, the other end fixed) and `delete-shape`. Static Width overrides, flips, coincident endpoints and broader 1-D formulas stay refused.
- Relative MoveTo/LineTo geometry scales with Width and Height; absolute coordinates need a supported dependency on the dimensions (zero coordinates excepted). Cycles, error caches, unknown affected dependencies and incomplete recalculation refuse the transaction.
- A batch of only `delete-shape` commands removes a closed set together: references inside the set are allowed, references from retained formulas, metadata or Connect rows refuse. Mixed batches keep the one-shape-at-a-time policy.

#### Copy and paste

- `duplicate-shapes` clones local unglued 2-D leaves with fresh IDs and names, keeping stacking order, mixed text and unknown XML, remapping references between the copies and proving affected caches. `visioDuplicateCommand` defaults to 0.33 drawing inches right and down.
- `captureVisioClipboard(bytes, pageId, shapeIds)` returns an owned `VisioClipboardSnapshot`; `serializeVisioClipboard` and `deserializeVisioClipboard` encode and validate bounded, versioned text and never touch a system clipboard. `paste-shapes` imports it through the same clone transaction after checking fonts, styles, theme, page settings and drawing ratio; source references must stay inside the copied selection. Limits: 1,000 shapes, 64 resource records, 8,388,608 UTF-16 code units and 100,000 XML nodes before any DOM is built. Resource import, cross-scale paste and native paste cascade are unsupported.

#### Masters and stencils

- `parseVsdx` lists visible masters as `document.masters` (`VisioMaster` with the master's shapes resolved alone, for a preview) and docked stencil file names as `document.stencils`; stencil files are never opened.
- `insert-master-instance` drops a master as Visio does: a shape naming the master with `PinX`/`PinY`, one `MasterShape` sub-shape per sub-shape of a group master, the page's relationship to the master part, Visio's naming (`Name.ID`), inherited size caches (`F="Inh"`), the master's layers joined by name, and an Auto Size page grown. A master with several top-level shapes drops as the `Type="Group"` shape Visio makes. The Dynamic connector drops with two free ends. Chained masters, pictures and multi-root masters with turned or flipped shapes are outside the scope. `visioMasterDropCommand` (`visio/ui`) builds the edit from a pointer position.
- `drop-stencil-master` drops a master of a built-in stencil (`VISIO_BUILT_IN_STENCILS`: Basic Shapes, Basic Flowchart Shapes, Miscellaneous Flowchart Shapes, Arrow Shapes), copying it into the document stencil the first time (found again by `UniqueID`, or reused by universal name when it is one droppable 2-D shape). These masters are this package's own geometry in the drawing's default styles, with no icon, connection points or Shape Data.
- `change-shape` with `masterId` (or a built-in `shape`) makes a one-shape 2-D instance an instance of another master, keeping position, local size, text, formatting, shape data, layers and whole-shape glue; the ID stays (Visio assigns a new one), so dynamic glue stays valid. Refused: glue to a connection point, glue when the size would change, local overrides of the old master's geometry, `LockReplace` and locked layers.

#### Master instances

- **Pin-only moves.** A master-linked top-level 2-D instance moves when all six transform caches are proven lengths, both pins are unguarded direct overrides and effective movement protection is proven inactive (Width, height, aspect and deletion locks may stay active). Only the two pin caches are written; any other cache reading a changed pin refuses the transaction. A formula the proof cannot follow (menus, add-on calls, container lookups) no longer refuses on its own, since it can only read another shape by naming its sheet.
- **Recalculated edits.** `resize-shape`, `rotate-shape`, `flip-shape`, `format-shape` and `format-text` accept a top-level instance of a one-shape master (`edit-instance-*.ts`), and `move-shape` takes this path when a connector is glued to it. A changed cell becomes a local value (`U="IN"` on new lengths); every inherited cell that follows from it is evaluated over the instance's effective sheet and written as a refreshed `F="Inh"` cache, so the master's formula stays in effect. Arithmetic, comparisons, `IF`, `MIN`, `MAX`, `BOUND`, `AND`/`OR`/`NOT` and references inside the instance are evaluated; a drawn cell depending on anything else refuses with `EDIT_UNSUPPORTED_DEPENDENCY`, while uncomputable data-section caches are left as saved. A glued connector that `rerouteConnector` cannot handle refuses the edit with `UNSUPPORTED_INSTANCE_EDIT`, because Visio does not re-lay connectors on open.
- **Group masters** (`edit-instance-scope.ts`) take the same edits across the group and every sub-shape; master `Sheet.N!` references resolve to the instance's sheets. Group formatting skips guarded or unformattable sub-shapes, as Visio's ribbon does, and fails only when every part refuses; a sub-shape's ID targets it alone. Refused for groups: a sub-shape pin or size with no formula, shapes added to the instance on the page, pictures inside, and moving or resizing one sub-shape alone.
- **Whole-shape commands** (`edit-instance-shape.ts`): `delete-shape` removes the instance with its sub-shapes and releases a stencil connector glued to it (glue formulas become values, the Connect row goes); `duplicate-shapes` and `paste-shapes` copy a 2-D instance as an instance of the same master (a paste needs the same master under the same ID and refuses between drawings; stencil lines and connectors are not copied); `reorder-shape`, `assign-layers` (a local `LayerMember`) and `group-shapes` take unglued 2-D instances. A rotated or flipped group holding instances cannot be ungrouped.
- **Formula scope.** `formula-functions.ts` lists the ShapeSheet functions Visio documents. A function outside it refuses; so do references built from text (`INDIRECT`, `EVALCELL`, `EVALTEXT`, `REF`). Container lookups (`CONTAINERSHEETREF` and its family) refuse only when the shape being reordered, copied or grouped is itself a container or list.
- Refused for all instances: masters that inherit another master, 1-D masters, locally deleted sections or rows, `GUARD`ed sizes and pins, active locks, locked layers, and another page shape whose formula reads a changed cell.

#### Layers

`set-layer-properties` sets a layer's `visible`, `print` or `lock` flag (the Visible, Print and Lock cells of the page's Layer section). Omitted flags are kept, a missing cell is added, a formula-driven flag is refused, and the edit changes only `pages.xml` in its own transaction. Renaming, removing and recolouring layers and the Active, Snap and Glue flags are not written.

#### Planners in `visio/ui`

The UI helpers return ordinary atomic commands; source admission in the core stays authoritative. `visioArrangeCommands` aligns to the first selected shape and distributes by equal gaps between rotated boxes, ordered by centres with selection-order ties, as measured in Visio. `visioResizeDrag`, `visioSizePositionState` and `visioSizePositionCommand` produce fixed-pin size, move and rotation commands (an exact current value yields no command). `visioTextReplaceOccurrences`, `visioTextReplaceFindNext`, `visioTextReplacePlan` and `visioTextReplaceCommands` plan literal, case-sensitive Find and Replace over `selection`, `current-page` or `all-pages` scopes and reject stale snapshots; matched text the editor cannot change stays in the batch so the whole operation is refused rather than silently skipped. Unicode case folding, whole-word matching and special search codes are unsupported.

### ShapeSheet formulas

Numeric ShapeSheet interpretation uses a bounded AST, never JavaScript evaluation. Arithmetic, comparisons, `IF`, `GUARD`, `MODULUS`, `AND`/`OR`/`NOT`, `BITAND`/`BITOR`/`BITXOR` (integral inputs in 0..65535), `ATAN2` (y, x), `SIGN` (fuzz 1e-9), Width/Height scaling, local geometry and named-row references and static `Sheet.ID` references have dependency analysis. DL/DP/DT caches are internal inches and DA radians; dimensional comparisons need explicit compatible units (`Width > 3 IN`). Squared-length intermediates support distance and `SQRT`, but area caches cannot be written into length cells.

- `F="No Formula"` is Visio's locally blocked marker, not an expression: its cache stays usable and the marker is kept when a value is set. `F="Inh"` always needs real inheritance resolution ([Cell element](https://learn.microsoft.com/en-us/office/client-developer/visio/cell-element-geometry-sectionvisio-xml)).
- Affected caches are computed before writeback with `F` and `U` kept; implicit pin, dimension and text dependencies take part. Provably independent unsupported formulas may stay unchanged; opaque, unknown-function or dynamic dependencies fail safely, and a final check refuses formulas that would contradict the requested dimensions or pin.
- Master and page IDs have separate scopes; untouched active masters get a separate bounded independence proof that resolves template, instance override and style cells in their real sheet scopes, allowing only direct single-reference `SETATREF` and field-free local `TheText` measurement forms ([SETATREF](https://learn.microsoft.com/en-us/office/client-developer/visio/setatref-function), [TEXTWIDTH](https://learn.microsoft.com/en-us/office/client-developer/visio/textwidth-function), [TEXTHEIGHT](https://learn.microsoft.com/en-us/office/client-developer/visio/textheight-function)). Theme reads record their selector dependencies; affected theme evaluation is unsupported.
- Limits: 8,192 formula characters, 1,024 AST nodes, depth 64, 100,000 indexed cells, 10,000 affected cells and 100,000 evaluation steps per recalculation; master proof construction 100,000 steps and 10,000 bindings, traversal 100,000 steps at depth 64; move preparation 100,000 work units. These are cooperative bounds within the package deadline.

### Exclusions and next expansions

| Exclusion                                                                  | Reason                                                                             | Next expansion                                                                  |
| -------------------------------------------------------------------------- | ---------------------------------------------------------------------------------- | ------------------------------------------------------------------------------- |
| Foreign shapes                                                             | Nested transforms and resource semantics are not proven                            | Foreign-shape resource semantics                                                |
| Glue/Connects participation and broader 1-D editing                        | Proven straight-line translation does not establish connector routing or glue      | Endpoint editing and glued connector routing                                    |
| Non-page affected or unknown dependencies                                  | Page metadata, document, master and style scopes could otherwise keep stale caches | Scoped package-wide graph, starting with page metadata and pure theme functions |
| `GUARD`, `SETATREF` and referenced transform formulas                      | Direct overwrites would bypass protection or redirection, or discard semantics     | Verified redirection commands; never bypass protection                          |
| Protected cells and inherited or ambiguous protection                      | LockMoveX/Y, LockWidth/Height/Aspect/Delete must be honoured                       | Proven effective protection resolution                                          |
| Absolute geometry lacking dimension-dependent formulas; nonlinear geometry | Scaling cached coordinates can distort shape semantics                             | More row evaluators and explicit scaling proofs                                 |
| Deletion referenced outside a removed set                                  | Retained formulas, Connects or metadata could dangle                               | Validated dependency-removal and connector-healing transactions                 |

Affected master caches, container membership changes and general connector routing still need explicit supported recalculation before the editing scope can grow. The hash-pinned public corpus results (which drawings and shapes admit each edit, and why the rest refuse) are in `edit-geometry-corpus.test.ts`, `edit-master-move-corpus.test.ts` and the [corpus setup](../../../viewers/visio/docs/corpus.md).

## Bounded processing

Default package limits: 32 MiB input, 2,048 entries, 16 MiB per inflated entry, 128 MiB total inflated bytes, a 200:1 compression ratio, 16 MiB of XML characters, nesting depth 128, 1,500,000 XML nodes per part and two million across the package, a memory budget, and a ten-second cooperative deadline (`maxRuntimeMs`).

**Memory.** One read (`parseVsdx`) or edit (`editVsdx`) has one budget, shared by every package it opens (an edit reads its own output again, and some edits re-read the drawing for text sizes or Auto Size) and by the model a read builds. Each XML part is costed while it is scanned, before its DOM exists, at the heap the shared XML parser was measured to keep (800 bytes an element, 240 an attribute, 140 any other node, one a character); an edit charges each part twice. The model costs 3 KB a shape, 16 bytes a geometry command, 200 bytes a text run or paragraph and 2 bytes a character. A settled step gives back what it charged, except the model a read returns. A refusal reports `LIMIT_XML_MEMORY` with the memory needed.

The budget is `maxXmlHeapBytes`. By default it is the heap limit the host reports (`v8.getHeapStatistics().heap_size_limit` in Node and Deno, `performance.memory.jsHeapSizeLimit` in Chromium) less 256 MiB, divided by 1.6, the most heap measured per charged byte; where nothing is reported (Firefox, Safari, web workers, Bun) a 1 GiB heap is assumed. Past 64 MiB a read also takes at most 4,096 bytes per input byte (`maxXmlHeapPerInputByte`). Hosts that hold other large data, or run several reads at once, should set it lower.

| Node heap (`--max-old-space-size`) | `heap_size_limit` | Budget  |
| ---------------------------------- | ----------------- | ------- |
| 512 MiB                            | 704 MiB           | 280 MiB |
| 768 MiB                            | 960 MiB           | 440 MiB |
| 1 GiB                              | 1,216 MiB         | 600 MiB |
| Node 24's default on a 64-bit host | about 4.2 GiB     | 2.5 GiB |
| not reported (browsers' workers)   | 1 GiB assumed     | 480 MiB |

A Visio-saved page of 10,000 Basic Shapes rectangles (211 KB, 1.28 million XML nodes) is charged 466 MiB, so it needs a heap limit of about 1,000 MiB (`--max-old-space-size=820` in Node); 10,000 plain generated shapes are charged 367 MiB. An edit is charged up to four times a read. Measured in Node 24 at 512, 768 and 1,024 MiB heaps over 29 dense, budget-filling, generated and Visio-saved files (255 runs), no run exhausted the heap: each opened, edited or was refused with `LIMIT_XML_MEMORY`, with peaks at or below 385, 648 and 902 MB for reads and 424, 669 and 952 MB for edits. The deadline is checked between steps, not inside one part's DOM construction (the per-part node limit bounds that), so the viewer's parse worker adds a 15-second hard timeout. A part over the node limit reports its shape count ("The page has 18000 shapes, more than this viewer can open").

**Package checks.** ZIP validation checks local and central headers, CRCs, aliases, traversal, overlap and unsupported encryption or ZIP64. XML rejects DTD and entity declarations and malformed or ambiguous namespaces. Internal relationships must resolve to existing parts inside the package; external ones stay inert.

**Scene limits.** 25,000 shapes including expanded inherited shapes, depth 64, 500,000 geometry commands, five million text characters, 100,000 text runs, 50,000 paragraphs and 2,000 diagnostics. Rasters: 8 MiB each, 16 million pixels, 16,384 pixels a side. Curves: 64 KiB formulas, 256 controls, degrees 1-25, 2,048 generated segments and 2 million work units per curve, with `maxCurveWork` (20 million) across the package; NURBS flatten at 0.0001 inch with floating-point fidelity unverified. Diagnostic strings have a one-million-character aggregate cap. Metadata: 256-character IDs and attribute names, 1,024-character fonts, 4,096-character display and layer names (over-long metadata throws `METADATA_LIMIT`, document text is never truncated). Layer membership: 8,192 cached characters and 1,024 tokens per shape, 100,000 tokens and five million characters per document. Shape Data and hyperlinks: 100,000 rows and five million characters per document, 1,024 rows per shape, 8,192 characters per string.

Most limits are configurable through `ParseVsdxOptions` (`limits`, `images`, `metadata`, `maxShapes`, `maxCurveWork` and the scene maxima). Runtime checks cannot preempt a JavaScript inflater chunk already running; put parsing in a worker when hard responsiveness matters.

## Metafiles and foreign vectors

- `inspectVisioEmfAdmission` does bounded classic-record inspection and always returns `renderingEnabled: false`; no converter runs. VSDX EMF parts get specific diagnostics through an inspection-only adapter (32 unique parts, 32 MiB, 100,000 records). EMF+, images, text, unresolved palette colours and `INSIDEFRAME` pens stay unsupported.
- `sanitizeVisioForeignVectorTree` turns a narrow SVG tree into inert numeric paths, literal paints, matrices and a closed clipping graph; `validateVisioForeignVector` validates it again after transport. Both return deeply frozen data and reject unsafe tags, URLs, fonts, styles, accessors, cycles and amplification.
- `convertVisioMetafile` is an experimental adapter for a trusted `emf-converter` function the host supplies, narrowed to stock-painted primitive shapes and lines (256 KiB, 512 records, 2,048 pixels; at most eight conversions and 1 MiB per document). `parseVsdx` accepts it as `metafileConverter`, for use only inside a disposable worker with a parent-owned deadline; it is off by default. Format conversion itself stays in `emf-converter`. See [bounded conversion](../../../docs/visio-metafile-conversion.md) and [legacy VSD](../../../docs/visio-legacy-vsd.md).

## Native evidence

Behaviour that claims Visio's output is checked against files Visio 16 saved, through its documented COM automation, never by inspecting its code. Each recorder in [`scripts/`](../../../scripts/) writes a directory that an optional test reads when the matching variable is set; without it the test is skipped. Examples:

| Recorder                                                                                                                                     | Variable                                                                                                                                     | Covers                                                          |
| -------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------- |
| `record-visio-new-drawing.ps1` (fixtures: `create-visio-new-drawing-fixtures.ts`)                                                            | `VISIO_NATIVE_NEW_DRAWING_DIR`                                                                                                               | Blank drawings at Letter, landscape Letter and A4               |
| `record-visio-formatting.ps1`                                                                                                                | `VISIO_NATIVE_FORMATTING_DIR`                                                                                                                | Formatting cells and stacking at three drawing scales           |
| `record-visio-line-patterns.ps1`                                                                                                             | `VISIO_NATIVE_LINE_PATTERNS_DIR`                                                                                                             | 432 cases: patterns 0-23, three caps, two weights, three scales |
| `record-visio-svg-zero-stroke.ps1`                                                                                                           | `VISIO_NATIVE_SVG_ZERO_STROKE_DIR`                                                                                                           | Zero-weight SVG strokes and arrows                              |
| `record-visio-text-box.ps1`                                                                                                                  | `VISIO_NATIVE_TEXT_BOX_DIR`, `VISIO_NATIVE_TEXT_BOX_CUSTOM_DIR`                                                                              | Fixed-size text boxes with default and custom styles            |
| `record-visio-text-replace.ps1`                                                                                                              | `VISIO_NATIVE_TEXT_REPLACE_DIR`                                                                                                              | Characters-range replacements                                   |
| `record-visio-size-position.ps1`, `record-visio-anchored-resize.ps1`                                                                         | `VISIO_NATIVE_SIZE_POSITION_DIR`, `VISIO_NATIVE_ANCHORED_RESIZE_DIR`                                                                         | Numeric Size & Position and anchored resize                     |
| `record-visio-batch-delete.ps1`                                                                                                              | `VISIO_NATIVE_BATCH_DELETE_DIR`                                                                                                              | Closed-set deletion and capture/delete/paste reopen             |
| `record-visio-instance-text.ps1`, `record-visio-instance-geometry.ps1`, `record-visio-group-instance.ps1`, `record-visio-instance-shape.ps1` | `VISIO_NATIVE_INSTANCE_TEXT_DIR`, `VISIO_NATIVE_INSTANCE_GEOMETRY_DIR`, `VISIO_NATIVE_GROUP_INSTANCE_DIR`, `VISIO_NATIVE_INSTANCE_SHAPE_DIR` | Master-instance text, geometry, group and whole-shape edits     |
| `record-visio-glyph-advances.ps1`, `record-visio-text-extent.ps1`                                                                            | `VISIO_NATIVE_TEXT_EXTENT`                                                                                                                   | Font advances and text-sized shapes                             |

Each recording proves the captured cases, not general visual equivalence, native window behaviour or every geometry and paint combination; the [verification record](../../../viewers/visio/docs/verification.md) has the case counts and results for every family.

## Development

```bash
bun run build:visio      # build only this area (from the repository root)
bun run test -- visio    # the core Vitest suite, filtered to this area
bun run typecheck        # the strict project, which includes this area
```

The Visio viewer's own checks (`viewers/visio`, `bun run check:core` and the rest) are listed in its [README](../../../viewers/visio/README.md#development). Synthetic fixture generators are not in the published package. Code moved here records its provenance in [`PROVENANCE.md`](../../../PROVENANCE.md).

## Further reading

- Viewer docs: [viewer API](../../../viewers/visio/docs/api.md), [architecture](../../../viewers/visio/docs/architecture.md), [collaboration](../../../viewers/visio/docs/collaboration.md), [capability ledger](../../../viewers/visio/docs/parity.md), [verification record](../../../viewers/visio/docs/verification.md).
- Repository docs: [Visio parity review](../../../docs/visio-parity-review.md), [connector rounding](../../../docs/visio-connector-rounding.md), [metafile conversion](../../../docs/visio-metafile-conversion.md), [legacy VSD](../../../docs/visio-legacy-vsd.md), [collaboration area](../../../docs/collab-area.md), [parity checklist](../../../docs/parity-checklist.md).
