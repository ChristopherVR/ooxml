# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.1](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@2.0.1) - 2026-11-27

### Documentation

- Rewrite the Excel viewer READMEs ([65fd477](https://github.com/ChristopherVR/ooxml/commit/65fd477dbbaf5b9f730a97a2b553c9da3a4f6e28))
- Close the privacy gaps for sharing, recovery and links ([72dfa8c](https://github.com/ChristopherVR/ooxml/commit/72dfa8cfe34f2073b6805a48b5c8f267b32f20fe))

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@2.0.0) - 2026-11-21

<!-- release-notes: 2.0 -->

`@christophervr/xlsx-core` re-exports `ooxml-core/xlsx` and `ooxml-core/xlsx/load`, so this release brings the 2.0 workbook model to headless code: servers, workers and scripts that read, calculate and write workbooks. Calculation is bounded per formula and can be bounded per recalculation, so an untrusted workbook can no longer hang a process, and lookups and aggregates over large ranges scale with the edit. Signatures are reported as a verified state, hyperlinks keep every safe scheme, and Node.js 22.3 is the minimum. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) has the before and after for each breaking change.

**[Introducing OOXML 2.0](https://christophervr.github.io/ooxml/releases/2.0/)** tells the story of this release. Every change to Excel is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.0/#changes-excel) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.0.md#changes-excel)).

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): the `xlsx` area this package re-exports; its changelog covers the rest of the core.
- [`@christophervr/xlsx-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/react/CHANGELOG.md): the React editor, which re-exports this API.
- [`xlsx-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/vue/CHANGELOG.md): the Vue editor, which re-exports this API.
- [`xlsx-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/angular/CHANGELOG.md): the Angular editor, which re-exports this API.
- [`xlsx-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/svelte/CHANGELOG.md): the Svelte editor, which re-exports this API.
- [`xlsx-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/solid/CHANGELOG.md): the Solid editor, which re-exports this API.
- [`xlsx-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/vanilla/CHANGELOG.md): the editor without a framework, which re-exports this API.
- [`xlsx-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/CHANGELOG.md): the Excel tools for AI agents, built on the same model.

## [1.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@1.1.2) - 2026-11-13

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@1.1.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@1.1.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@1.0.0) - 2026-10-29

### Features

- **xlsx:** Commit xlsx-core and xlsx-viewer-mcp to a stable 1.0 api ([e458ecf](https://github.com/ChristopherVR/ooxml/commit/e458ecf368421096ccffb366cc1e0ba17f438b77))

## [0.3.1](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.3.1) - 2026-10-23

### Other

- Word and Excel docs for wave 6 (docs6-4) ([da0f3e8](https://github.com/ChristopherVR/ooxml/commit/da0f3e83c59b8ff55104bd852cf46fa052c587b5))

### Documentation

- Describe wave 6 chart and smartart editing in word and excel ([d212908](https://github.com/ChristopherVR/ooxml/commit/d212908f59be7f5da51ce5347b23a665e4cedb08))

## [0.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.3.0) - 2026-10-21

### Features

- **xlsx:** Add native framework building blocks ([9689a65](https://github.com/ChristopherVR/ooxml/commit/9689a6520682e9db18f8cbee78d8434ad114b097))

## [0.2.10](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.2.10) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.2.9](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.2.9) - 2026-10-08

## [0.2.8](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.2.8) - 2026-10-07

## [0.2.7](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.2.7) - 2026-10-07

## [0.2.6](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.2.6) - 2026-10-07

## [0.2.5](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.2.5) - 2026-10-06

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.2.4](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@0.2.4) - 2026-10-05

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

## [0.2.3](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/@christophervr/xlsx-core@0.2.3) - 2026-10-05

### Build & CI

- **deps:** Adopt ooxml-core 0.17.0 and ooxml-ui 0.25.0 (by @ChristopherVR) ([49705cd](https://github.com/ChristopherVR/xlsx-viewer/commit/49705cd85cae5c14275ef15a544c689a3867abe4))
- **deps:** Adopt ooxml-core 0.18.0 and ooxml-ui 0.27.1 (by @ChristopherVR) ([6e66a43](https://github.com/ChristopherVR/xlsx-viewer/commit/6e66a43e8966eedc4bd19934e27085be76d33a95))
- **deps:** Adopt ooxml-core 0.19.0 (by @ChristopherVR) ([d981768](https://github.com/ChristopherVR/xlsx-viewer/commit/d981768e1f6c5f612ea6c69aaf9507933a5a694f))

## [0.2.2](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/@christophervr/xlsx-core@0.2.2) - 2026-10-05

### Build & CI

- **deps:** Adopt ooxml-core 0.14.1 and ooxml-ui 0.20 (by @ChristopherVR) ([2da396d](https://github.com/ChristopherVR/xlsx-viewer/commit/2da396d66cad5cb74915d665258ac8216c3cd602))
- **deps:** Adopt ooxml-core 0.15.0 and ooxml-ui 0.21.0 (by @ChristopherVR) ([6cbd2b5](https://github.com/ChristopherVR/xlsx-viewer/commit/6cbd2b5ad65734a3d9ce7f2c8feff93d975eafe5))

## [0.2.1](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/@christophervr/xlsx-core@0.2.1) - 2026-10-03

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([afbbf9c](https://github.com/ChristopherVR/xlsx-viewer/commit/afbbf9c9a88f70e0afeec47da60e6cd9339c76ba))

## [0.2.0](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/@christophervr/xlsx-core@0.2.0) - 2026-10-03

### Features

- **web-component:** Adopt core encryption and workbook features (by @ChristopherVR) ([857348f](https://github.com/ChristopherVR/xlsx-viewer/commit/857348fb817ce99d773c8d4dad6258ff1dc93384))

## [0.1.2](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/@christophervr/xlsx-core@0.1.2) - 2026-10-03

### Bug Fixes

- **deps:** Depend on ooxml-core ^0.9.0 and ooxml-ui ^0.1.8 (by @ChristopherVR) ([e2a8ae6](https://github.com/ChristopherVR/xlsx-viewer/commit/e2a8ae63533262f4fef4fb8c93feadd3ca3b15b8))

## [0.1.1](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/@christophervr/xlsx-core@0.1.1) - 2026-10-03

### Features

- Add the Excel viewer and editor (by @ChristopherVR) ([408fe71](https://github.com/ChristopherVR/xlsx-viewer/commit/408fe71adc283e887cca4e873673daa6934957d5))

