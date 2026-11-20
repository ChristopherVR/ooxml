# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

`visio-core` is the DOM-free Visio document API, and every read and edit through it now runs inside one memory budget sized from the host's heap, so no file can take down the Node process or worker that opens it. Pages of up to about 10,000 shapes as Visio saves them open with the new default limits, `maxXmlHeapBytes` lets a server that holds other large data lower the budget, and `parseVsdx` says up front when a drawing is too large to edit. The one breaking change is the runtime: Node.js 22.3 or newer. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) gives the before and after.

### Breaking changes

- The package requires Node.js 22.3 or newer and declares `"engines": { "node": ">=22.3" }`. The memory budget is sized from the heap limit Node reports, which `ooxml-core` reaches through `process.getBuiltinModule`, added in Node 22.3; on an older Node the budget assumes a 1 GiB heap, more than a host started with a smaller `--max-old-space-size` has ([guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#run-nodejs-223-or-newer)).

### Security

- No Visio file can exhaust the heap of the host that reads or edits it. A read or an edit has one memory budget for every package it opens and the model it builds, estimated before each XML part is parsed. A drawing that does not fit is refused with the code `LIMIT_XML_MEMORY` and a message saying how much memory it would need. Past 64 MiB a read may take at most 4,096 bytes of memory for each byte of its file, so a small file of dense XML is refused in well under a second. Measured in Node 24 under 512, 768 and 1024 MiB heaps over 29 files, every file opened or was refused, with peaks at or below 424, 669 and 952 MB; before, a 411 KB file could end a 512 or 768 MiB process.
- Every package is opened through the guarded zip reader, including signature listing, collaboration snapshots and embedded chart workbooks, so a small file can no longer inflate without limit. Signature listing ignores a signature part larger than 16 MiB.
- A collaborator can no longer replace an arbitrary package part through shared media, or add markup to the content types part.
- A drawing is reported as signed only when its signature verifies, never because it merely contains a signature part.
- The package is published by a release pipeline in which third-party code cannot push, tag or publish, and the signed tarball on each GitHub release is checked against the integrity the npm registry records.

### Added

- Package limits gain `maxXmlHeapBytes` and `maxXmlHeapPerInputByte` (4,096 by default), which bound the memory one read or edit may take. By default `maxXmlHeapBytes` follows the host's heap: the limit Node (`v8.getHeapStatistics().heap_size_limit`) or a Chromium page (`performance.memory.jsHeapSizeLimit`) reports, less 256 MiB, divided by 1.6. That is 280 MiB under `--max-old-space-size=512`, 600 MiB under 1024 and about 2.5 GiB under Node's usual 4 GiB default; where the runtime reports no limit (Firefox, Safari, web workers, Bun) it is 480 MiB. Set it lower when the same heap holds other large data or runs several reads at once.
- Drawings open with pages of up to about 10,000 shapes as Visio saves them. The default XML limits are now 1,500,000 nodes per part and 2,000,000 per package; the ten-second load deadline is unchanged. Such a page needs a heap limit of about 1,000 MiB in Node (`--max-old-space-size=820` or more).
- `parseVsdx` reports `edit-size-limit` first among its diagnostics when the drawing is too large to edit: its pages hold more than 100,000 ShapeSheet cells, or an edit of it would not fit `maxXmlHeapBytes`.

### Changed

- The package is built and tested against current releases of its dependencies, including `emf-converter` 5. The supported browsers are the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): the `visio` area this package re-exports, where the memory budget and limits live.
- [`visio-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/react/CHANGELOG.md): the viewer for React, which re-exports this API.
- [`visio-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/vue/CHANGELOG.md): the viewer for Vue, which re-exports this API.
- [`visio-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/angular/CHANGELOG.md): the viewer for Angular, which re-exports this API.
- [`visio-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/svelte/CHANGELOG.md): the viewer for Svelte, which re-exports this API.
- [`visio-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/solid/CHANGELOG.md): the viewer for Solid, which re-exports this API.
- [`visio-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/vanilla/CHANGELOG.md): the viewer for pages without a framework, which re-exports this API.
- [`visio-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/mcp/CHANGELOG.md): the same reads and edits as MCP tools.

## [1.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@1.1.2) - 2026-11-13

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@1.1.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@1.1.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

### Other

- Bring origin/main into the embedding work ([a3711a1](https://github.com/ChristopherVR/ooxml/commit/a3711a14049a2cd11b994e76eafec039db2dc2de))

## [1.0.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@1.0.1) - 2026-10-22

### Documentation

- **visio:** Sync the viewer docs with the monorepo and current code ([2436dcc](https://github.com/ChristopherVR/ooxml/commit/2436dcce519728d28a65a124947b6dbbb5854cd6))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@1.0.0) - 2026-10-22

### Features

- **visio:** Complete the building blocks in every framework ([0a69177](https://github.com/ChristopherVR/ooxml/commit/0a6917719b5040587c0a37dfe3b37a875b7d2a05))

### Bug Fixes

- **visio:** Ship the toolbar controls with the svelte package ([709fafa](https://github.com/ChristopherVR/ooxml/commit/709fafa524945526e613472143abf41054131332))

## [0.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.5.0) - 2026-10-21

### Features

- **visio:** Add native framework building block integrations ([689e934](https://github.com/ChristopherVR/ooxml/commit/689e9340824a41e32a5024938e16b483e34c4452))

## [0.4.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.4.1) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.4.0) - 2026-10-08

### Features

- **visio:** Enable new drawings and pointer selection gestures ([75ac409](https://github.com/ChristopherVR/ooxml/commit/75ac409e5e67c23203c765d3aa472b6817853090))

## [0.3.7](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.3.7) - 2026-10-08

## [0.3.6](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.3.6) - 2026-10-07

## [0.3.5](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.3.5) - 2026-10-07

### Refactor

- **visio:** Use extensionless relative source imports ([2495aa9](https://github.com/ChristopherVR/ooxml/commit/2495aa92d52a0135106695a2a39a72c9f98cf177))

## [0.3.4](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.3.4) - 2026-10-07

## [0.3.3](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.3.3) - 2026-10-06

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.3.2](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@0.3.2) - 2026-10-05

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

## 0.3.1

### Changes

- build(deps): adopt ooxml-core 0.15.0 and ooxml-ui 0.21.0 (818f4a4)

## 0.3.0

### Changes

- feat(viewer): add visio account, options dialog and live sharing (fe7fed1)

## 0.2.1

### Changes

- fix(packages): pin released core and shared ui dependencies (4d035f6)

## 0.2.0

### Changes

- feat(bindings): expose native reactive viewer state in every binding (ccd9ba8)
- feat(viewer): add visio ribbon toolbar on shared office-ui controls (afc7047)
- fix(deps): use released core for proven master shape moves (28989be)

## 0.1.1

### Changes

- docs(packages): align npm readmes with powerpoint structure (5c6f120)

## 0.1.0

### Changes

- feat(release): publish functional visio framework packages (bc55722)

## 0.0.2

### Changes

- Initial release.
