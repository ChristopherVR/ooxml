# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-core@2.0.0) - 2026-11-21

<!-- release-notes: 2.0 -->

`visio-core` is the DOM-free Visio document API, and every read and edit through it now runs inside one memory budget sized from the host's heap, so no file can take down the Node process or worker that opens it. Pages of up to about 10,000 shapes as Visio saves them open with the new default limits, `maxXmlHeapBytes` lets a server that holds other large data lower the budget, and `parseVsdx` says up front when a drawing is too large to edit. The one breaking change is the runtime: Node.js 22.3 or newer. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) gives the before and after.

**[Introducing OOXML 2.0](https://christophervr.github.io/ooxml/releases/2.0/)** tells the story of this release. Every change to Visio is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.0/#changes-visio) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.0.md#changes-visio)).

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
