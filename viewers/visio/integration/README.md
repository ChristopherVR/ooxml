# Unreleased core development bridge

`ooxml-visio.patch` contains the new canonical `ooxml-core/visio` area and its integration changes, against the exact revision in `core-revision.txt`. It is retained privately so this viewer can be built before the core changes are separately approved and released.

Do not maintain a second implementation here. Make format changes only in the sibling `ooxml` checkout, rerun its strict checks and tests, then regenerate this patch. The setup script never resets or overwrites an existing checkout.

Once the core area is released, replace the root `file:../ooxml` dependency with the reviewed published version, remove this bridge and verify a clean install and all entry points.

## Released converter package integration

Normal development uses the exact published `emf-converter` 4.8.8 dev dependency
and the root npm lock's registry integrity. `npm ci` installs it; no converter
source checkout, patch application or source build is required. The core adapter
accepts a trusted package function, so this viewer-only dependency does not change
the core's existing converter dependency or its publication commit.

`npm run setup:converter` verifies the installed package version, lock integrity and
browser bundle hash recorded in `emf-converter-release.json`. `npm run check` now
includes `npm run check:converter`, which runs seven deadline-isolated integration
cases against the package's actual browser export. The released browser bundle is
byte-identical to the previously verified corrected source build. Node's browser
condition is explicit in that test worker and its resolved export is asserted;
Node codecs are never selected for these tests.

The core adapter accepts only a small generated-case classic EMF subset, makes a
private input copy, reinspects it, uses fixed non-raster converter options and
returns independently validated neutral vectors. It contains no EMF renderer.
Live document conversion remains disabled. These generated regressions do not
establish native Visio fidelity or a hard peak-memory quota.

The converter is Apache-2.0. Its unmodified license and shipped third-party notices
remain in `emf-converter-LICENSE.txt` and `emf-converter-THIRD_PARTY_NOTICES.txt`.

## Historical source correction and audit

`emf-converter-visio.patch` and its revision/lock files preserve the correction
provenance against v4.8.7. Those changes were published in source commit
`5a5df709b256803e66199293033fb47d27c3a787` and package 4.8.8. They are no longer an
unreleased prerequisite for this viewer. To reproduce that historical source
checkpoint explicitly, use `npm run setup:converter-source` and
`npm run check:converter-source`; these keep their isolated
`emf-converter-current` checkout and refuse unrelated edits. The root check runs
only synthetic setup safety tests, without creating this checkout.

The older `emf-converter-visio` checkout and core's original converter dependency
remain separate for the 3.5.1 characterization audit. Do not rewrite its known-defect
assertions as release evidence. Neither normal setup nor checks publish anything.
