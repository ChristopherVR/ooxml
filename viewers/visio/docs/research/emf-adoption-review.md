# Embedded EMF adoption review

Date: 2026-10-02. Scope: the installed local converter, synthetic regressions, and the existing local Visio corpus. No converter repository changes, dependency updates, production activation, uploads, browser launches or publication were made.

## Decision

**Do not enable unrestricted `emf-converter` conversion in the viewer.** A dedicated worker plus strict preflight and a fail-closed, renderer-neutral output boundary is a defensible direction for a small, explicitly documented subset. The worker alone is insufficient protection, and the installed converter cannot currently provide trustworthy completeness diagnostics.

The audit prototype admits a deliberately tiny classic-EMF vector subset. **It admits none of the five actual EMF media parts found in the corpus.** Do not advertise those files as supported based on this experiment. All five produce converter output without warnings; that establishes neither fidelity nor safe admission.

The most useful next candidate is a larger, independently checked classic-EMF subset for `60973.vsdx` media `image2.emf` and `image3.emf`. These have no EMF+ records, embedded bitmap drawing records, font creation or text output. They still need mapping, clipping, palette-relative colors, line styles and more geometry than the initial subset. Their public comments also require subtype validation. Do not simply remove or skip unsupported records.

## Exact implementation examined

Installed under the sibling core repository:

- `emf-converter` **3.5.1**, package metadata and shipped `LICENSE`: Apache-2.0.
- ESM bundle SHA-256: `99b64b6333561c9c851bf7b5c06be8f2935f97447b4eb8721c3b8157c9110b40`.
- CJS bundle SHA-256, used by the isolated reproductions: `2127569d3337d19f98e44344fb51af914a9dbb830f6ded8170983d779d9b8527`.
- Linux, Node **24.19.0**, optional `@napi-rs/canvas` **1.0.10** already installed. The upstream issue reports instead name Windows 11, Node 26.7.0 and canvas 1.0.9.
- Public APIs include `convertMetafileToSvgTree`, SVG string/data-URL serializers and PNG conversion. Prefer the tree API at a validation boundary. The package also exports React/JSX helpers; those are unnecessary for this viewer.
- Source references below identify original module/function comments in the installed ESM bundle; line numbers are specific to the hash above.

### Distribution and provenance

The package has no declared runtime dependencies; the optional native canvas peer adds a separate native dependency and transitive licensing surface. The canvas package declares MIT and ships its license. Do not bundle that Node-only backend into the browser viewer.

Before shipping any converter code, preserve the Apache license and applicable attribution, identify modifications, and include upstream NOTICE attribution if supplied by the selected distribution. No NOTICE file was present in this installed converter package. This is a distribution checklist, not a provenance certification. See [Apache-2.0, section 4](https://www.apache.org/licenses/LICENSE-2.0).

The corpus stays external. Apache POI `60973.vsdx` is pinned by the existing corpus manifest to commit `732120980140d5ed64b482c470e0b625cdb1ab15`, document SHA-256 `c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03`. [Pinned source fixture](https://github.com/apache/poi/blob/732120980140d5ed64b482c470e0b625cdb1ab15/test-data/diagram/60973.vsdx). Preserve the existing fixture provenance and notices; the converter's license does not license document content or fonts. No original document, extracted media or generated document image is included in these audit files.

## Reproduced fidelity defects

All measurements below are from new local synthetic buffers. Windows GDI+/Visio reference pixels were **not** independently produced. The GitHub issue bodies were read through the authorized connector after the web fetch returned a cache miss. Versions 4.8.7 and 4.3.5 were not installed or tested.

| Case                                                            | Installed 3.5.1 observation                                      | Diagnosis and limit                                                                                                                                                                                                                     |
| --------------------------------------------------------------- | ---------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Plain rectangle                                                 | 299 x 199 canvas, red box 298 x 198                              | Control for the mapping case; edge rasterization is not claimed exact.                                                                                                                                                                  |
| MM_TEXT with window 900 x 600 and viewport 300 x 200            | Red box 99 x 66, identical to the explicitly anisotropic control | `handleCoordinateRecord` activates mapping on extents irrespective of current map mode, lines 22969-23008.                                                                                                                              |
| DrawString control                                              | 1,274 dark pixels in the text area                               | Same Arial font request and fill as the driver-string controls. Font substitution may differ on another platform.                                                                                                                       |
| DrawDriverString, five glyphs                                   | Zero dark pixels                                                 | The handler aligns the position array to four bytes instead of reading immediately after the 16-bit glyph array. Unpadded DataSize is rejected; padding DataSize like issue #19 produces misread coordinates and still no visible text. |
| DrawDriverString, four glyphs                                   | 1,065 dark pixels                                                | Refutes the broad claim that every such record necessarily paints nothing. The handler paints one string using only the first glyph position.                                                                                           |
| Four driver glyphs with a +40 x transform matrix                | Identical text-node coordinates and pixel metrics to no matrix   | The handler ignores the optional matrix and driver flags. Raw glyph indices, vertical text and per-glyph advances are not faithfully handled.                                                                                           |
| Narrow clip, then replace by a full leaf region                 | Red box remains 75 x 199, zero red pixels to its right           | `parseEmfPlusRegionObject` rejects `RegionNodeCount === 0`, lines 25060-25073. A valid single terminal region has no child nodes.                                                                                                       |
| Full leaf region, no previous narrow clip                       | Full canvas red                                                  | Not proof that SetClipRegion succeeded: ignoring it produces the same result.                                                                                                                                                           |
| Diagnostic mutation: same leaf with the incorrect child count 1 | Full canvas red even after narrowing                             | Shows replacement works once the object parser accepts it. This intentionally inconsistent count is a diagnostic probe, not a conforming reference fixture.                                                                             |

The three failures return **zero console warnings** in these cases. Therefore checking only `null` or warnings would silently accept wrong output.

Source reports: [#19, driver strings](https://github.com/ChristopherVR/emf-converter/issues/19), [#20, region clip replacement](https://github.com/ChristopherVR/emf-converter/issues/20), [#21, mapping extents](https://github.com/ChristopherVR/emf-converter/issues/21).

Primary format/behavior references:

- [Microsoft DrawDriverString record](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-emfplus/b794b780-e9a8-4682-af65-9b614aecfbe6) specifies glyph values, positions and the optional matrix. The local source's alignment and single-string replay differ from that structure.
- [Microsoft MS-EMFPLUS specification, section 2.2.1.8, PDF page 63](https://winprotocoldoc.z19.web.core.windows.net/MS-EMFPLUS/%5BMS-EMFPLUS%5D.pdf) defines RegionNodeCount as child nodes and includes one root in addition to them. Rejecting zero excludes leaf regions.
- [Microsoft SetViewportExtEx](https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-setviewportextex) says fixed mapping modes, including MM_TEXT, ignore window/viewport extents. Issue #21's immediate reproduction is EMF; its WMF relevance comes from a WMF carrying an embedded EMF. No WMF conversion was exercised in this audit.

## Parsing, resource and output hazards

| Boundary                     | Concrete evidence in installed code                                                                                                                                                                                                           | Required treatment                                                                                                                                                                                              |
| ---------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Header validation            | `parseEmfHeader`, lines 9881-9915, checks length and type but not the EMF signature, declared byte/record count, version or EOF. A zero signature and a truncated-off EOF both still produced output without warnings.                        | Preflight exact framing, signature/version, complete EOF and declared counts before invoking the converter.                                                                                                     |
| Record-local bounds          | EMF replay checks outer record sizes. EMF+ loops do not prove `DataSize <= Size - 12`; comment lengths are not confined to their enclosing EMR_COMMENT. Several image/text helpers check against the whole DataView rather than their record. | Validate each record/subrecord span and every offset/count/product against its own payload. A caught RangeError is not a validator.                                                                             |
| Partial success              | Unknown EMR records warn and are skipped. Other short/invalid payloads return early; replay can finish successfully. `emfWarn` and `emfLog` are **empty functions** in this bundle, lines 798-801. Many omission paths are completely silent. | Maintain an independent exhaustive admission whitelist; reject the complete image on unsupported/invalid content. Require explicit converter completion status before broader adoption.                         |
| Limits not global            | Replay defaults are 200,000 EMF/WMF and 500,000 EMF+ records. EMF+ counts restart per comment. Predecode walkers use separate fixed 500,000 limits and do not consume caller `maxRecords`.                                                    | One shared aggregate budget including comments, continuations, nested streams, paths, objects and decoder work. Validate before conversion.                                                                     |
| Predecode before surface cap | `replayMetafile`, around 27575, calls image/texture predecode before creating a bounded output surface.                                                                                                                                       | Output width/height limits do not bound source image allocations. For the first subset, reject every image, texture and nested metafile.                                                                        |
| Bitmap allocation            | DIB and raw EMF+ bitmap paths permit dimensions up to 8192 each. One 8192-square RGBA buffer is 256 MiB, before duplicate buffers/canvas/cache. EMF+ stride/byte ranges may be checked against the whole input.                               | Enforce dimensions, pixels, stride, decompressed byte lengths and aggregate decoded bytes before allocation. Do not experimentally allocate maximum-size payloads in this shared process.                       |
| Compressed images            | `decodeCompressedBytesToRgba` checks dimensions after `decodeDeferredImageBytes`. The built-in PNG inflater materializes a decompressed buffer. Encoded PNG/JPEG/GIF/WebP bytes can also be embedded verbatim in SVG.                         | Parse bounded image headers and enforce decompressed limits before codecs; reject compressed input until this is implemented. SVG embedding is not a decoder-resource bypass.                                   |
| Nesting/continuations        | The package has depth-3 nested-metafile guards and a 64 MiB continuation cap, but caches and budgets are not globally shared. PNG nested conversion calls omit caller options.                                                                | Set initial nesting/continuation allowance to zero. Later share byte/record/pixel budgets across all descendants; do not reset budgets per stream.                                                              |
| Text/fonts                   | CSS font families are derived from file content. Default text is environment dependent. `loadSystemFonts` reads system directories if explicitly called; supplied fonts enable a TrueType interpreter.                                        | Initial subset excludes text and fonts. Never call `loadSystemFonts`, fetch document-named fonts, or enable unreviewed font bytecode. Later use controlled licensed fonts and expose substitution diagnostics.  |
| Output expansion             | SVG trees, accumulated paths, defs, data URLs and serializer arrays have no total output cap. Repeated drawing can expand a small input many times.                                                                                           | Apply operation/coordinate bounds before conversion and iterative node/depth/character validation before serializing or posting output. Reject, never truncate a drawing.                                       |
| Browser worker               | Worker termination stops JS and isolates responsiveness; it is not a browser process/native-memory quota. A worker may still have network APIs.                                                                                               | Dedicated worker, trusted static bundle, restrictive CSP and no network operations, plus structural allocation bounds. If a hard total-memory guarantee is required, a normal browser worker cannot provide it. |

The converter implements many classic-GDI and EMF+ drawing/state handlers: lines, Beziers, polygons, ellipses/arcs, paths, brushes/pens, clipping, text, images and raster operations. A switch case is **not** evidence of full record semantics. Examples above return “handled” despite losing information. The initial admitted record list below is intentionally much smaller than the converter's advertised coverage.

### External references, executable output and fonts

Inspection found no direct `fetch`, XHR, WebSocket, script evaluation or HTML insertion in the bundled converter. Browser image decoding uses byte-backed Blob/createImageBitmap; the optional Node decoder receives a Uint8Array. Known embedded image types are sniffed into data payloads. Nested metafiles can become nested SVG data URLs. This is a limited code review, not a complete dependency/security certification.

The SVG string serializer escapes attribute values and text, but accepts arbitrary node tag/attribute names in its public tree input. Escaping alone is not a DOM security contract. A document-driven font family can also select a host page's web font and induce a font fetch when rendered there. The prototype avoids both concerns by rejecting text, all URLs, image nodes, arbitrary CSS, links, use/foreignObject, filters, animation, scripts and event attributes. Do not insert converter-produced markup with `innerHTML` or feed its tree straight to React/createElement.

## Actual corpus coverage

The runner checks both curated VSDX fixture directories, not their security/fuzz directories or duplicate upstream checkouts. It finds **five embedded media EMFs**, all in Apache POI `60973.vsdx`, plus **17 separately extracted thumbnail previews**. A thumbnail is not page media and is not an independent Visio-export oracle after conversion by this same package.

| Media part | Input bytes | EMR / EMF+ records | SVG nodes / attribute+text chars | Initial subset                        |
| ---------- | ----------: | -----------------: | -------------------------------: | ------------------------------------- |
| image1.emf |      18,028 |           311 / 20 |                     113 / 15,594 | Reject                                |
| image2.emf |       3,736 |            107 / 0 |                       38 / 2,713 | Reject; next classic-vector candidate |
| image3.emf |       4,216 |            100 / 0 |                       55 / 2,751 | Reject; next classic-vector candidate |
| image4.emf |       6,792 |           108 / 18 |                       22 / 4,537 | Reject                                |
| image5.emf |       8,760 |           179 / 20 |                       51 / 4,826 | Reject                                |

All five passed the prototype's framing checks and produced raw converter SVG/PNG with zero console warnings. These raw conversions are **audit-only**, on named reviewed local fixtures, and deliberately bypass subset eligibility. No such bypass should exist in production. The metrics above do not constitute a visual comparison.

None of these five media streams contains DrawDriverString or SetClipRegion. `image2`/`image3` explicitly set MM_ANISOTROPIC before extents, so the specific MM_TEXT regression is not their immediate blocker. Their public GDIC comments carry alternate WMF data. Future classic-only admission would have to validate and explicitly define handling of that comment subtype; blindly allowing all comments would reintroduce EMF+ and nesting hazards. They also use ROP2 13 (copy pen), the default palette, palette-relative colors, clipping, and styled pens. `image2` needs ARC/ARCTo; `image3` needs Polygon16.

Six thumbnail previews contain SetClipRegion: blue-box, qs-box, testfile4, testfile5, testfile6 and tdf136564-WhiteTextBackground. A record's presence alone does not prove the demonstrated defect changes a specific preview. Existing corpus warnings about stale/blank thumbnails remain applicable.

### Exact record inventory

Each entry is `record ID: count`, in numeric order. EMR_HEADER and EMR_EOF are included. EMF+ records are additionally counted inside their EMR_COMMENT parent. IDs are retained so the inventory can be mechanically checked without guessing a handler name. The evidence JSON also records media hashes and per-record diagnostic offsets.

#### visio/media/image1.emf

SHA-256: `6b8c0738e26d69f2e4d882d30a4294951d53c485cf0247bf9229d46ff45ce9f7`.

EMR: `1: 1`, `13: 2`, `14: 1`, `18: 2`, `19: 34`, `20: 1`, `21: 1`, `22: 1`, `24: 1`, `25: 1`, `27: 1`, `28: 1`, `33: 5`, `34: 5`, `35: 6`, `36: 6`, `37: 78`, `39: 30`, `40: 52`, `48: 1`, `58: 3`, `70: 12`, `76: 1`, `82: 1`, `91: 38`, `95: 21`, `98: 3`, `115: 2`.

EMF+: `0x4001: 1`, `0x4002: 1`, `0x4004: 3`, `0x4008: 2`, `0x401b: 1`, `0x401e: 1`, `0x4021: 3`, `0x4022: 1`, `0x4024: 1`, `0x402a: 4`, `0x402c: 1`, `0x4030: 1`.

#### visio/media/image2.emf

SHA-256: `87543078d0617e08a4649d6141fc443418903f1f951417490d0348c91c6e49b4`.

EMR: `1: 1`, `9: 2`, `10: 1`, `11: 1`, `14: 1`, `17: 1`, `18: 1`, `20: 2`, `22: 2`, `24: 2`, `25: 2`, `30: 1`, `33: 3`, `34: 3`, `37: 32`, `38: 10`, `39: 2`, `40: 12`, `42: 9`, `45: 8`, `47: 8`, `48: 1`, `70: 1`, `75: 1`.

EMF+: none.

#### visio/media/image3.emf

SHA-256: `8093b61d382b703f07e499f239377b6d0092230d69d3f779a209cec682a89ee0`.

EMR: `1: 1`, `9: 2`, `10: 1`, `11: 1`, `14: 1`, `17: 1`, `18: 1`, `20: 2`, `22: 2`, `24: 2`, `25: 2`, `27: 2`, `30: 3`, `33: 5`, `34: 5`, `37: 24`, `38: 5`, `39: 5`, `40: 10`, `42: 2`, `48: 1`, `54: 2`, `70: 1`, `75: 1`, `86: 18`.

EMF+: none.

#### visio/media/image4.emf

SHA-256: `284aa6995f05cb6c2bf53eb2cb890e5f2c63111448209a65cf4f139f8c081c30`.

EMR: `1: 1`, `9: 2`, `10: 2`, `11: 1`, `12: 1`, `14: 1`, `17: 1`, `19: 11`, `20: 1`, `22: 1`, `24: 1`, `25: 1`, `33: 4`, `34: 4`, `37: 23`, `38: 6`, `39: 5`, `40: 12`, `48: 1`, `70: 8`, `75: 1`, `76: 1`, `82: 1`, `87: 4`, `91: 11`, `98: 3`.

EMF+: `0x4001: 1`, `0x4002: 1`, `0x4004: 2`, `0x4008: 2`, `0x401b: 1`, `0x401e: 1`, `0x4021: 3`, `0x4022: 1`, `0x4024: 1`, `0x402a: 3`, `0x402c: 1`, `0x4030: 1`.

#### visio/media/image5.emf

SHA-256: `9b4bdbfd51f136d985b8d6b304a7e338d6ab5bad34ba9f1a43fb7e7d685d38ca`.

EMR: `1: 1`, `13: 2`, `14: 1`, `18: 2`, `19: 14`, `20: 1`, `21: 1`, `22: 1`, `24: 1`, `25: 1`, `27: 5`, `28: 1`, `33: 5`, `34: 5`, `35: 6`, `36: 6`, `37: 29`, `39: 7`, `40: 17`, `48: 1`, `58: 7`, `59: 4`, `60: 4`, `61: 4`, `62: 2`, `64: 2`, `70: 12`, `76: 1`, `82: 1`, `88: 4`, `90: 6`, `91: 11`, `95: 9`, `98: 3`, `115: 2`.

EMF+: `0x4001: 1`, `0x4002: 1`, `0x4004: 3`, `0x4008: 2`, `0x401b: 1`, `0x401e: 1`, `0x4021: 3`, `0x4022: 1`, `0x4024: 1`, `0x402a: 4`, `0x402c: 1`, `0x4030: 1`.

## Bounded admission and result contract

The files under `scripts/emf-audit/` are isolated experiment code, not a viewer/core API. Any production parser/admission contract belongs in `ooxml-core`; viewer code owns the worker lifecycle and DOM rendering. Do not copy the experiment into a second production parsing implementation.

### Initial subset demonstrated

- EMR_HEADER(1), EOF(14), MM_TEXT-only SETMAPMODE(17), SETPOLYFILLMODE(19), MOVETOEX(27), balanced SAVEDC(33)/RESTOREDC(-1)(34), SELECTOBJECT(37), solid/null CREATEPEN(38), solid/null CREATEBRUSHINDIRECT(39), DELETEOBJECT(40), ELLIPSE(42), RECTANGLE(43), LINETO(54).
- Strict fixed payload sizes, object/stock-handle checks, literal colors, bounded coordinates and pen widths. Header descriptions must remain within the header record.
- No unknown records, comments, images, EMF+, WMF, text/font records, palette-relative colors, clipping, paths, mapping extents, world transforms or nested metafiles. A valid file outside the subset is “unsupported”, not evidence that the source document is corrupt.
- The preflight's `structurallyValid` describes its bounded framing and fixed-payload checks, not full EMF conformance. Diagnostics about objects following an unsupported object-creation record are subset limitations, not independent claims that the original handle is invalid.
- The prototype demonstrates admission and neutral output for a synthetic rectangle. It is not a full visual-conformance suite for every admitted drawing record. Do not activate it merely because this audit exists.

### Proposed production budgets

| Resource                  | Initial limit / enforcement                                                                                                                                                                                                                                                                                    |
| ------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Encoded input             | 4 MiB per image, checked before copies/worker transfer; enforce an additional document-wide media budget and ZIP entry/expanded-size budget in core.                                                                                                                                                           |
| Records                   | 20,000 total EMR + EMF+ records; no reset per comment or nested stream.                                                                                                                                                                                                                                        |
| Record/object payload     | Validate fixed lengths and all variable arithmetic before slicing/allocating; at most 4,096 handles.                                                                                                                                                                                                           |
| State/region/output depth | 64; initial subset has no region/path trees or nested metafiles.                                                                                                                                                                                                                                               |
| Source coordinates        | Absolute value at most 1,000,000; cap transformed coordinates and reject non-finite values. Render within a bounded, clipped viewport.                                                                                                                                                                         |
| Images and fonts          | Zero in the initial subset. Before expansion, enforce image-header and decoder-specific decompressed-byte limits plus a shared allocation ledger, not merely source-byte or canvas-size caps.                                                                                                                  |
| Output surface            | At most 2048 per dimension and 4,194,304 pixels. The audit runner uses a smaller 512 x 512 maximum. Disable exactRasterOps/raster mirrors for the vector-only subset and reject records that require them.                                                                                                     |
| Output tree               | At most 25,000 nodes, depth 64, 2 MiB of strings and 100,000 numeric path operands in the sanitizer. Reject excessive output before serializer/postMessage.                                                                                                                                                    |
| Time                      | One dedicated worker; terminate after 3 seconds per image. Cancel immediately on document replacement/close and ignore stale-generation responses.                                                                                                                                                             |
| Concurrency/cache         | One conversion at a time, bounded queue, content-hash deduplication, explicit cache byte accounting/eviction; suggested aggregate retained-media budget 32 MiB.                                                                                                                                                |
| Memory guarantee          | Structural limits and image/font exclusion are enforceable. Browser worker total/native memory is **not** a hard quota. The Node audit worker's 96 MiB old-generation / 16 MiB young-generation settings also exclude ArrayBuffer/native allocations. Do not call those settings a 112 MiB total-memory limit. |

A timeout measured outside the worker is necessary: a Promise.race inside the same synchronous conversion cannot interrupt it. The test suite includes a deliberate synchronous infinite loop confined to the audit worker and verifies termination. No OOM stress test was run on the shared host.

### Fail-closed diagnostics and renderer-neutral output

Proposed result variants are `ok(vector)`, `unsupported(diagnostics)`, `invalid(diagnostics)`, `budget-exceeded(diagnostics)`, `timeout`, `cancelled` and `conversion-failed`. Include media-part and shape references, byte offset, EMR/EMF+ type, relevant limit and sanitized reason. Cap repeated diagnostics and report omitted counts. Do not log document text or binary content.

The prototype uses `unsupported` for any failed admission, with structural status and detailed codes available in the scan; a production adapter should split invalid and limit failures into the variants above. Examples include `emf.unsupported-record`, `emf.known-fidelity-risk`, `emf.comment-range`, `emf.plus-record-size`, `emf.record-limit`, `emf.coordinate-limit` and `emf.unsafe-output`.

The sanitization experiment converts a converter-owned SVG tree into numeric vector/group/path commands and explicit literal paint fields. It does not forward tags, arbitrary attributes, style strings, font names, URLs, data URLs, text or raw markup. Validate again after worker messaging and construct only trusted renderer primitives. Keep the original media bytes in the document model for preservation; display a bounded unsupported-media placeholder if the conversion is not wholly admitted. No generic partial-output escape hatch.

Before broader adoption, require converter hooks that enforce operation/allocation/output budgets while replaying, and an explicit completion/unsupported-record result. An after-the-fact sanitizer cannot prevent the converter from allocating an oversized tree first. For the tiny subset, preflight fixed sizes and record counts constrain reachable operations; each expansion must re-establish that argument.

## Reproduction and verification

From the viewer repository:

```sh
node --test scripts/emf-audit/audit.test.mjs
node scripts/emf-audit/run.mjs "$VISIO_CORPUS" "$EMF_AUDIT_OUTPUT"
```

Set `VISIO_CORPUS` to the existing external corpus root (containing the apache-poi and libvisio directories), and `EMF_AUDIT_OUTPUT` to an external scratch output directory. Both arguments are required. The runner writes `evidence.json`, including bundle hashes, environment, every synthetic result, each actual media record inventory, and thumbnail inventories. It does not save document images or media bytes. It uses the already installed sibling dependency and never installs or updates packages. The test suite verifies the exact 3.5.1 CJS bundle hash before executing its known-defect assertions.

The Node tests cover framing, record-local boundaries, aggregate/input/depth limits, deterministic truncations and 500 header mutations, fail-closed output validation, safe synthetic conversion and watchdog termination. Tests titled “3.5.1 characterization” intentionally assert the observed defects; passing those tests does not mean the defects are fixed. If the converter changes, compare evidence rather than updating assertions mechanically.

Verification at this review: **15/15 focused tests passed**, including rejection of oversized/sparse child arrays before output allocation, all five named embedded media parts were converted for audit, and 17 preview streams were inventoried. No browser/Windows test was run. The parent task remains responsible for full repository checks after parallel changes; this review does not claim a full viewer check pass or Visio parity.

## Adoption gates

1. Keep production rendering disabled until the exact record subset and failure UI meet the reviewed safety/fidelity checks. This is a technical acceptance gate within the authorized viewer work, not a request for new user permission. Resolve the three proven semantics defects before admitting their records; upgrading alone is not evidence. Remote converter changes and publication remain outside this audit's authorization.
2. Pin and re-audit the exact converter distribution, retain attribution, and exclude native Node code from browser bundles. Add a dependency hash/version-sensitive regression job.
3. Move the independently bounded admission model into the appropriate core metafile area, with shape/part diagnostics. Reuse the existing single viewer renderer and framework bindings.
4. Implement worker cancellation, timeout, stale-result isolation, no-network CSP, output revalidation and bounded cache/queue tests. Browser memory and font behaviors still require supported browser QA.
5. Expand classic GDI toward image2/image3 only after record-specific length/count/palette/clip/mapping tests and an independent render comparison. Keep the three mixed EMF+ media files rejected until their object/image paths meet the same budget rules.
6. Obtain trustworthy Windows/Visio paired references for representative media, odd/even driver strings, glyph indices, transforms, clipping and MM_TEXT/fixed/anisotropic mapping. Track substitution, unsupported-record and budget diagnostics alongside pixels. Self-generated or converter-against-itself comparisons are not parity evidence.
