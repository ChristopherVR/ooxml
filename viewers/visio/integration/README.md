# Core and converter integration

Normal installs use released registry packages: `ooxml-core` 0.14.1 and
`ooxml-ui` 0.15.0, with exact root pins and npm lock integrity. The seven
distribution packages require compatible released ranges. No sibling checkout,
development link or source patch is required for normal builds and checks:

```sh
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
npm run check
npm run test:browser
```

Format conformance remains in the core repository. This viewer tests the
released API, inert legacy VSD preview, VSDX editing, packed consumers and workers.
Legacy VSD is read-only here and does not establish native Visio visual fidelity.

## Historical optional core source helper

`core-revision.txt` retains the historical source baseline
`2c6e66afa7520e882d89498bebd117eb966476ea` from `ChristopherVR/ooxml`.
That commit already contains the canonical `ooxml-core/visio` area. No historical
Visio bootstrap patch is needed to reproduce it.

For intentional source development, `npm run setup:core` clones the selected
checkout when absent, checks out this historical pin, verifies the checkout,
installs the reviewed `core-package-lock.json` with `npm ci --ignore-scripts`, and
builds the Visio export. The integration lock was retained after verifying all
published manifest dependency/workspace fields match and successfully running a
clean npm CI install and Visio build. The upstream repository uses Bun; this lock
exists solely to make the viewer's npm CI setup deterministic.

Existing checkouts are never reset or patched by setup. Wrong revisions, staged
changes, unexpected tracked/untracked source changes and different npm locks are
refused before any install or build. Equivalent npm lock formatting is accepted.
`npm run test:core-setup` covers these refusal paths, the clean baseline, optional
patches, isolated cloning and local dependency linking without network access.

## Trying local core changes

All parsing and document-model changes still belong in the core repository.
To leave the canonical sibling untouched while testing a separate checkout, set
`VISIO_CORE_DIR` to its absolute path or a path relative to the viewer root:

```sh
VISIO_CORE_DIR=../my-ooxml-checkout npm run setup:core
npm ci --ignore-scripts
npm ci --prefix packages/bindings --ignore-scripts
VISIO_CORE_DIR=../my-ooxml-checkout npm run link:core
VISIO_CORE_DIR=../my-ooxml-checkout npm run check
VISIO_CORE_DIR=../my-ooxml-checkout npm run test:browser
```

For an intentionally modified checkout, build it manually instead of running
setup until its changes have been reviewed. `link:core` changes only the installed
`node_modules/ooxml-core` symlink. It refuses to overwrite a real directory and
leaves both package manifests and locks unchanged. Run it again after `npm ci`,
which restores the released registry dependency. Source overrides are local only;
restore normal registry installation before committing or validating release
consumers. Historical 3.5.1 converter audits retain their
original sibling-core dependency for that separate characterization evidence.

When unreleased changes need a reproducible private bridge, run
`VISIO_CORE_DIR=../my-ooxml-checkout node scripts/update-core-patch.mjs` after core
checks. It requires the pinned HEAD and an unstaged index, and records tracked and
untracked changes as optional `ooxml-visio.patch`. Empty patches are removed. Setup
applies this patch only to a newly cloned checkout and requires an exact match for
an existing checkout. Validate a fresh isolated setup before sharing a patch.
After a separately approved core publication, deliberately advance the pin and
remove incorporated changes from the optional patch. Do not maintain a second
format implementation in this repository.

## Released converter package integration

Normal development uses the exact published `emf-converter` 4.8.9 production
dependency and the root npm lock's registry integrity. `npm ci` installs it; no
converter source checkout, patch application or source build is required. The
core adapter accepts a trusted package function, so this viewer dependency does
not change the core's existing converter dependency or its publication baseline.

`npm run setup:converter` verifies the installed package version, production
dependency declaration, lock integrity and browser bundle SHA-256 recorded in
`emf-converter-release.json`. Registry `gitHead` and the peeled v4.8.9 tag both
identify source commit `f0ca94d898cc4bf6f51e8c4e81085a61a5bf5bf0`.
`npm run check:converter` runs deadline-isolated integration cases against the
package's actual browser export. The browser condition is explicit in that test
worker and its resolved export is asserted. Browser/worker bundle checks reject
Node-native canvas and filesystem loading.

The adapter accepts only a bounded generated-case classic EMF subset, copies and
reinspects the input, uses fixed non-raster converter options and returns
independently validated neutral vectors. It contains no EMF renderer. Generated
regressions do not establish native Visio fidelity or a hard peak-memory quota.
Unsupported records still need honest diagnostics and fallback presentation.

The converter is Apache-2.0. Its unmodified license and shipped third-party notices
remain in `emf-converter-LICENSE.txt` and `emf-converter-THIRD_PARTY_NOTICES.txt`.
Both were checked against the 4.8.9 package and remain unchanged.

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
