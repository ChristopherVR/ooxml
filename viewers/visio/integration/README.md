# Unreleased core development bridge

`ooxml-visio.patch` contains the new canonical `ooxml-core/visio` area and its integration changes, against the exact revision in `core-revision.txt`. It is retained privately so this viewer can be built before the core changes are separately approved and released.

Do not maintain a second implementation here. Make format changes only in the sibling `ooxml` checkout, rerun its strict checks and tests, then regenerate this patch. The setup script never resets or overwrites an existing checkout.

Once the core area is released, replace the root `file:../ooxml` dependency with the reviewed published version, remove this bridge and verify a clean install and all entry points.
