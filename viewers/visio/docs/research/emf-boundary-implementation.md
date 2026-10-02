# Enhanced-metafile boundary checkpoint

2026-10-02. These changes improve inspection and isolate a future rendering boundary. EMF conversion in the live viewer remains disabled, and this is not Microsoft Visio parity.

## Implemented

- Core `inspectVisioEmfAdmission`: bounded classic record inspection, fixed/variable payload checks, graphics state and object lifetimes, clipping/mapping limits, and explicit compatibility exclusions. Every result says `renderingEnabled: false`.
- The VSDX importer inspects internal `EnhMetaFile` parts and reports specific reasons for omission. It never invokes the converter. Limits cover 32 unique assets, 32 MiB materialized media and 100,000 inspected records. Pending work is reserved and queued; repeated instances share a promise.
- Validated ZIP sizes reject oversized raster/metafile resources before inflation. ZIP integrity and actual expansion limits remain independent.
- Core `sanitizeVisioForeignVectorTree` and `validateVisioForeignVector`: separate unknown-tree and post-transport boundaries, producing independent frozen data with numeric commands, literal paint, bounded affine transforms and closed clip indices. No document IDs, URLs, scripts, styles, fonts or images survive the boundary.
- A prospective viewer renderer constructs fixed SVG tags and generated local resource identities. It is tested separately and is not attached to the live document model.

## Review findings resolved

Independent probes exposed repeated zero-operand path command amplification through clip reuse. Expanded command, operand and node budgets now apply to all use sites. Additional tests cover sparse arrays before allocation, getters/proxies, cyclic or dangling clip graphs, cumulative transforms, corrected arc radii and transport mutations.

Admission now excludes characterized converter failures involving mapping-mode switches, mapping changes under saved device contexts, incomplete anisotropic setup and full-circle/same-ray arcs. Intrinsic typed-array access avoids caller-defined properties and rejects shared, resizable or detached inputs. These exclusions do not repair the converter.

A record-count preflight can reject a large later asset while preserving the remaining budget for a smaller asset. Actual scan work, encoded bytes and unique part counts stay separately bounded. Exact exhaustion prevents later materialization.

## Real corpus and pixel evidence

All five named EMFs in the hash-pinned Apache POI 60973 drawing remain unsupported. The two classic assets have valid equivalent-WMF preservation records, but still require unresolved palette-relative colors and INSIDEFRAME pen behavior. The mixed EMF+ assets remain outside the subset. No fixture bytes are included in the repository.

Generated path-only SVG tests meet expected clip, transform and viewport pixels using native canvas 1.0.10's Skia SVG backend. The installed librsvg 2.60.0/Cairo 1.18.4 backend fails nested clipPath intersections. The divergence is reported explicitly; `scripts/test-foreign-vector-raster.mjs --require-librsvg` retains a conformance failure rather than accepting its pixels. The [SVG 1.1 clipping specification](https://www.w3.org/TR/SVG11/masking.html#EstablishingANewClippingPath) defines those intersections. Browser and native Visio verification remain absent.

Independent final review passed 227 focused core tests, four renderer boundary tests, 12 external adversarial regressions and strict TypeScript. The full aggregate counts and remaining limits are in [verification](../verification.md). Tests establish bounds and the supported semantics; they are not a guarantee of converter heap limits or complete visual fidelity.

## Reproduce

Run the normal setup and `npm run check`. To include the two optional original media checks, set `VISIO_EMF_CORPUS_FILE` to an explicitly supplied local copy of the hash-pinned `60973.vsdx` before the check. No automatic fixture download occurs.

After building, run `node scripts/test-foreign-vector-raster.mjs` for the secondary vector pixel checks. Run `npm run test:emf-audit` for pinned converter characterization. Tests bearing the word “characterization” deliberately detect existing converter defects; passing them does not mean those defects are fixed.
