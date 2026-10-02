# Unreleased core development bridge

`ooxml-visio.patch` contains the new canonical `ooxml-core/visio` area and its integration changes, against the exact revision in `core-revision.txt`. It is retained privately so this viewer can be built before the core changes are separately approved and released.

Do not maintain a second implementation here. Make format changes only in the sibling `ooxml` checkout, rerun its strict checks and tests, then regenerate this patch. The setup script never resets or overwrites an existing checkout.

Once the core area is released, replace the root `file:../ooxml` dependency with the reviewed published version, remove this bridge and verify a clean install and all entry points.

## Isolated converter package integration

`emf-converter-visio.patch` records local, unreleased corrections against upstream
v4.8.7 at the exact revision in `emf-converter-revision.txt`. The separate
`emf-converter-patched-revision.txt` identifies the reviewed local commit carrying
those corrections. `npm run setup:converter` creates or verifies the sibling
`emf-converter-current` checkout, installs the pinned npm lock and builds it.
An existing checkout must be at either recorded revision and its effective source
changes must match the patch exactly. Unrelated changes are never overwritten.
The older `emf-converter-visio` checkout and original core dependency stay available
for the 3.5.1 characterization audit; neither is replaced by this setup.

Run `npm run check:converter` for package tests, types, builds and the browser-package
check. After building the core, run `npm run test:converter-integration` to exercise
the actual browser distribution in a disposable worker through the core's bounded
adapter. The adapter accepts only a small generated-case classic EMF subset, makes
a private input copy, reinspects it, uses fixed non-raster converter options and
returns only independently validated neutral vectors. It contains no EMF renderer.

This bridge does not enable live document conversion or establish native Visio
fidelity. It remains isolated pending broader record-specific evidence, resource
limits and browser verification. A post-conversion output limit is not a guarantee
of peak converter heap use. The parent worker deadline can terminate conversion,
but cannot impose a hard memory quota.

The converter is Apache-2.0. Its unmodified license and shipped third-party notices
are retained as `emf-converter-LICENSE.txt` and
`emf-converter-THIRD_PARTY_NOTICES.txt`. The local patch changes inclusive header
bounds, mapping/state behavior and narrow primitive geometry. Package regressions
requiring unavailable native Windows fonts may skip; report skips explicitly.
The patch is not a new upstream release and setup does not publish anything.
