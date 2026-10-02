# Visio scenes and experimental text saving

`ooxml-core/visio` imports VSDX drawing packages into a DOM-independent typed scene.
UI components live in the separate viewer. This is a supported subset, not Visio
parity or a ShapeSheet calculation engine. A separate experimental plain-text
save API preserves untouched package payloads under the limitations documented below.

```ts
import { parseVsdx, getVisioPageLayers } from 'ooxml-core/visio';

const drawing = await parseVsdx(bytes);
const layers = getVisioPageLayers(drawing, drawing.pages[0]!.id);
```

Coordinates are inches, y-up; angles are radians. A shape's six-number affine
matrix maps its local coordinates into its parent's coordinate system. Geometry
is inert SVG path data, never source markup. Apply the page y-axis inversion only
at the rendering boundary, and invert text and raster images appropriately.
Background page layers are returned in painting order.

## Implemented

- Page identity, dimensions, order, background chains, and connector attachment metadata
- Group transforms, rotation, flips, pin/local-pin positions, and cached connector endpoints
- Master cells, sections, rows, child inheritance and deletion; style chains and document defaults
- Absolute/relative move and line rows; relative quadratic/cubic Beziers; circular arcs,
  three-point elliptical arcs, and ellipses with orthogonal axes
- Saved corner radius on explicitly closed, axis-aligned rectangular line geometry;
  exact circular SVG arcs with radius clamped to half the shortest edge
- Literal `POLYLINE` coordinate lists; bounded approximation of literal, clamped,
  positive-weight NURBS. Cached contiguous `SplineStart`/`SplineKnot` sequences use
  the same sampler with implicit unit weights. Both produce explicit approximation diagnostics
- Cached solid colors/fills, line widths/pattern codes, arrow codes/sizes, and opacity
- Cached/themed normalized line caps; legacy square/extended mappings are marked
  as inferred compatibility until verified against native geometry
- Cached line patterns 2-23 normalized to bounded dash/gap sequences, with inferred
  spacing diagnostics. Custom master patterns and themed dashes remain unresolved
- Saved Quick Style color selectors and internal DrawingML theme color/variant records;
  solid fills, linear gradient endpoints/stops, tint/shade/alpha color transforms,
  and themed line widths, with diagnostics for approximations
- Saved horizontal linear gradients with complete cached stops and per-stop opacity;
  explicit root selection resolves existing themed stop cells without inventing rows
- Explicit root-style selection for saved themed colors and scalar line/fill formats;
  root fill patterns and opacity honor zero Quick Style fill-matrix selection
- Character runs, fonts, bold/italic/underline, text box transforms and margins;
  paragraph offsets, alignment, indentation, spacing, direction, bullet metadata,
  and cached text backdrop colors/transparency
- Page layer membership and cached visibility. MS-VSDX path visibility rules apply;
  layer colors, printability and locking are retained as metadata
- Embedded PNG, GIF and common 8-bit JPEG, with verified format signatures,
  structural validation, dimensions and byte limits. Instances share byte arrays;
  consumers must treat these as immutable and release any created object URLs
- Inherited Shape Data rows with cached types, values, units, labels and errors;
  inert hyperlink metadata with separately validated external targets

## Explicit limits

Only saved cached values, supported theme records and numeric coordinate-list literals are used. No formulas,
macros, addons, hyperlinks, external data, or network resources are executed or
fetched. Unsupported theme effects/transforms, complex paint, arbitrary-path corner rounding,
layer color overrides, nonliteral/periodic
splines, most foreign/vector records, OLE objects, APNG, uncommon JPEG coding, and arbitrary image
clipping are unsupported. Resized masters with inherited absolute geometry are
warned because their formulas are not recalculated. Master roots that themselves
inherit another master are diagnosed. Missing/unsupported features are returned
as diagnostics; unsafe or malformed packages throw `VisioPackageError`.

Text layout and raster decoding remain viewer responsibilities. Paragraph offsets
are UTF-16 offsets into `plainText`, excluding newline delimiters. Exact line
spacing is in inches; multiple spacing scales the largest font on a line. Raster
placement is local y-up inches and should be clipped to the shape frame. Header
validation does not replace browser image decoding or guarantee pixel fidelity.
Text backdrops apply behind laid-out text. Supported theme gradients expose local
y-up inch endpoints and at most 32 sorted stops. Radial and non-horizontal saved gradients, incomplete saved
gradient overrides, and unsupported color transforms retain a solid fallback and
diagnostic. Selected theme effects and unsupported line styling are diagnosed
separately. Visio-specific ignored theme attributes do not alter normalized paint.
Root fill-format selection replaces only existing `Themed` scalar caches, preserves
literal caches and leaves missing root values unresolved. Only complete saved horizontal gradients (zero or pi radians modulo turns), with
shape rotation enabled and group gradients disabled, are normalized. Existing themed
root stop cells resolve at matching indices; literal caches remain authoritative.
The first ten active stop rows are used. Hatch rendering and arbitrary saved
gradients remain unsupported; native visual equivalence is unverified.
Normalized `style.lineCap` maps directly to SVG `stroke-linecap`. An absent effective
cell leaves it undefined; consumers retain their documented fallback. Themed caps
resolve independently of line width and use Office's flat default when the selected
theme line omits its cap. Cached caps 1/2 map to butt/square based on observed upstream
compatibility, with `inferred-line-cap` diagnostics; they are not an exact-fidelity claim.
Normalized `style.lineDash` contains alternating dash/gap lengths in stroke-width
multiples, unlike the scene's inch coordinates. Consumers multiply by the actual
`lineWidth` without a minimum-width floor. The cached built-in sequences have 2-6
entries, each at most 27, and retain `inferred-line-pattern` diagnostics because
MS-VSDX pictures the patterns without specifying exact numerical spacing. Cached
values override unused themed dashes. A requested themed pattern stays unresolved
with a solid fallback; custom pattern 254 and invalid enumerations are diagnosed.
Line properties, including arrowheads, must be suppressed when geometry has
`stroke: false` or the effective `linePattern` is 0.
Rectangle rounding preserves local coordinates and winding, and honors inherited
cached line styles. Open orthogonal line chains also support the saved radius when
every segment has enough room for adjacent tangent arcs, preserving endpoints.
Short-segment clamping, duplicates, reversals, curves and mixed subpaths remain
diagnosed and unchanged. See [connector rounding evidence](../../docs/visio-connector-rounding.md).
Negative radii are ignored; radii below output precision are diagnosed.
Shape Data dates remain serial days and durations remain days; formats are retained
without locale-dependent evaluation. Raw hyperlink addresses are untrusted metadata.
Relative/file targets, external addresses with subaddresses, and unsupported link
options remain unresolved. Nothing is navigated or fetched during parsing.

## Bounded processing

Default package limits: 32 MiB input, 2,048 entries, 16 MiB per inflated entry,
128 MiB total declared/actual inflated bytes, 200:1 compression ratio, 16 MiB XML
characters, 128 XML nesting depth, 250,000 XML nodes per part, one million XML nodes
across the package, and a ten-second cooperative
deadline. ZIP validation checks local and central headers, CRCs, aliases, traversal,
overlap and unsupported encryption/ZIP64. XML rejects DTD/entity declarations and
malformed or ambiguous namespace constructs. Internal relationships must resolve
inside the package to existing parts.

Scene limits: 25,000 shapes/expanded inherited shapes, depth 64, 500,000 geometry
commands, five million text characters, 100,000 text runs, 50,000 paragraphs and 2,000 diagnostics. Raster defaults:
8 MiB each, 16 million pixels, and a maximum dimension of 16,384 pixels. Literal
curve limits include 64 KiB formulas, 256 controls and 2,048 generated segments.
NURBS uses positive-weight control-hull flatness at 0.0001 inch, with floating-point
fidelity unverified. Weighted curve work is also bounded per curve and per package.
Diagnostic strings are bounded before deduplication, with a one-million-character
aggregate cap. Metadata limits include 256-character source IDs/attribute names,
128-character geometry row types, 1,024-character fonts/inherited IDs, and
4,096-character display names. Metadata limits do not truncate document text or
geometry formulas; excessive metadata throws `METADATA_LIMIT`.
Layer membership limits are 8,192 cached characters before trimming, 1,024 tokens
per shape before deduplication, 256 characters per normalized ID, and 100,000
processed membership tokens plus five million cached characters across all pages
and inherited instances. These include duplicate tokens and whitespace. Layer
names are limited to 4,096 characters. Each page uses one indexed layer lookup.
Shape Data/hyperlink budgets are separate: 100,000 rows and five million retained
characters per document, at most 1,024 rows per shape and 8,192 characters per
metadata string. `ParseVsdxOptions.metadata` configures the aggregate limits.

Selected package and scene limits are configurable via `ParseVsdxOptions`. Runtime checks cannot preempt an
already executing JavaScript inflater chunk; output retained by the reader is
bounded. Put parsing in an application worker when hard UI responsiveness is
required.

## Local development

Run `npm run build:visio` for just this area, or the repository's normal build for
all areas. Tests: `npx vitest run src/visio`. Strict typecheck remains the base
repository TypeScript project. Synthetic fixture generators are not shipped.

Cached spline sequences require a preceding endpoint and one or more contiguous knot
rows. They are bounded to 256 controls, degrees 1-25, 2,048 output segments and
2 million weighted refinement/subdivision work units per curve. The package default
`maxCurveWork` is 20 million aggregate units, checked during refinement and subdivision.
Maximum degree and control counts are independent ceilings and may not be combined
within the work budget. Only clamped, nondecreasing compact knot vectors are supported;
missing numeric caches, malformed order and unsupported nonclamped/periodic
sequences omit the complete geometry section with diagnostics. Approximation uses
positive-weight homogeneous knot insertion and Bezier control-hull subdivision
with a 0.0001-inch flatness threshold. This avoids fixed-sample aliasing, but
floating-point fidelity is not certified. Generated analytic
quadratic/cubic tests cover sampled chord error, but no genuine spline-row fixture
or native Visio comparison is yet available.

## Metafile inspection and bounded worker conversion

`inspectVisioEmfAdmission` performs bounded classic-record inspection and returns
`renderingEnabled: false` in every result. No converter runs. VSDX enhanced-metafile
parts receive specific diagnostics through an inspection-only package adapter,
limited to 32 unique parts, 32 MiB aggregate bytes and 100,000 inspected records.
Parts exceeding available encoded-byte budgets are rejected from validated ZIP
metadata before inflation. Concurrent duplicate requests share a reserved promise,
and distinct inspections are queued so aggregate budgets remain deterministic.
Original-equivalent WMF comments are recognized only in their documented position
with validated lengths/flags and whole-file checksum; the payload is never decoded.
All EMF+, images, text, unresolved palette colors and INSIDEFRAME pens remain
unsupported, as do characterized converter mapping/full-circle edge cases.

`sanitizeVisioForeignVectorTree` converts a narrow unknown SVG tree to inert numeric
paths, literal paints, matrices and a closed indexed clipping graph.
`validateVisioForeignVector` validates and copies that canonical scene again after
transport or raw host input. Both return deeply frozen independent data and reject
unsafe tags, URLs, fonts, styles, sparse arrays, accessors, cycles, invalid paths,
coordinate/transform amplification and expanded clip work. These helpers are not a
metafile renderer or a guarantee of converter completeness. The private viewer renders the supported primitive subset only through its
deadline-controlled parser worker. The non-worker fallback remains converter-free.

`convertVisioMetafile` is an experimental adapter for a trusted
`emf-converter` browser-package function supplied by the host. It copies and
reinspects input, narrows admission further to stock-painted primitive shapes
and lines, fixes non-raster conversion options, and validates the resulting tree.
Its 256 KiB, 512-record, 2048-pixel limits do not establish a hard heap quota.
Use it only inside a disposable worker with a parent-owned deadline; it neither
loads Node codecs nor enables conversion by default. `parseVsdx` accepts an optional
trusted `metafileConverter` callback for worker-only use. Per-document admission
allows at most eight unique conversions and 1 MiB aggregate input. See
[bounded conversion](../../docs/visio-metafile-conversion.md) for the exact subset. All format conversion remains in
`emf-converter`; the adapter owns admission and the inert output boundary only.

## Experimental plain-text save

`editVsdx(input, edits, options)` applies an atomic, source-backed batch of
`{ type: 'replace-plain-text', pageId, shapeId, text }` commands and returns
`{ bytes, changedParts, diagnostics }`. It edits one existing local plain-text
`Text` element per command. It does not serialize the normalized rendering model.

Untouched package part payloads retain their exact uncompressed bytes, including
unknown parts, relationships and content types. An edited XML part preserves
unrelated content semantically, not lexically. ZIP headers, comments, directory
entries, extra fields and compression representation are not preservation targets.
A no-op returns an independent copy of the original archive. Changed saves use
STORE compression to avoid introducing compression ratios outside reader limits;
files may grow and exceeding the output limit fails without returning a result.

Only VSDX drawing packages are admitted. Signed/macro-indicated packages,
master-linked or deleted targets, missing/ambiguous IDs, rich text, fields and
unknown text markup are rejected. Replacement text must contain valid XML
characters and cannot contain carriage returns. Edited documents containing
carriage-return text or tab/newline/carriage-return attribute values are rejected
conservatively because this slice does not establish their serialization fidelity.
UTF-16 source XML becomes declared UTF-8 only in the edited part.

Every retained payload is inflated through the bounded reader and CRC checked.
External relationships remain inert. Unknown untouched payloads are preserved as
opaque bytes; this is not a safety certification of their content. Dirty output
is re-admitted with the remaining transaction deadline. Package limits still
apply, plus defaults of 1,000 commands, one million replacement UTF-16 code units
and an output ceiling equal to `maxInputBytes`. Output is streamed with an actual
byte limit. Runtime checks are cooperative, not hard CPU/heap quotas; applications
should run editing in a disposable worker for responsiveness.

Calls own input bytes and command values before their first asynchronous step.
Each call has a fresh bounded operation lifetime; do not keep a `VisioPackage`
reader alive as an interactive editing session. No public mutable DOM is exposed.

There is no geometry editing, rich-text editing, master override creation,
formula evaluation, dependent-cache refresh, undo UI or native save parity claim.
Text-dependent formulas can have stale caches after a save. Every changed result
reports `edit-caches-not-recalculated`. Automated evidence covers reopening in
this library and preservation/security invariants; reopening in Microsoft Visio
and native rendering fidelity remain unverified.
