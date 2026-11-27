# Changelog

All notable changes to this viewer's packages, newest first. Each section names the packages
released that day; the full history of each package is in its own `CHANGELOG.md`
(`packages/<name>/`, `mcp/`, `server/`).

## 2026-11-27

<!-- release-notes: 2.5 -->

_Releases: visio-viewer-mcp@2.0.1, visio-core@2.0.1, visio-react-viewer@2.0.1, visio-vue-viewer@2.0.1, visio-angular-viewer@2.0.1, visio-svelte-viewer@2.0.1, visio-solid-viewer@2.0.1, visio-vanilla-viewer@2.0.1_

Every Visio package moves to 2.0.1, released with `ooxml-core` 2.5.0 and `ooxml-ui` 2.5.0. The bindings keep their props and mount the 2.5 Visio editor: stencils you make yourself with New Stencil, Add Selected Shape, Edit Master and Save Stencil, SVG, BMP, TIFF and EMF pictures, embedded OLE objects drawn from their saved previews with Open Object, connectors that copy, drop, reroute and space as Visio routes them, and Get Add-ins. The Svelte load handle now passes the file name through. Upgrade every OOXML package together; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.5.md#visio-recordsets-are-numbered-from-1) covers the one change to check, recordsets numbered from 1.

**[Introducing OOXML 2.5](https://christophervr.github.io/ooxml/releases/2.5/)** tells the story of this release. Every change to Visio is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.5/#changes-visio) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.5.md#changes-visio)).

### Packages

- [`visio-viewer-mcp` 2.0.1](mcp/CHANGELOG.md)
- [`visio-core` 2.0.1](packages/core/CHANGELOG.md)
- [`visio-react-viewer` 2.0.1](packages/react/CHANGELOG.md)
- [`visio-vue-viewer` 2.0.1](packages/vue/CHANGELOG.md)
- [`visio-angular-viewer` 2.0.1](packages/angular/CHANGELOG.md)
- [`visio-svelte-viewer` 2.0.1](packages/svelte/CHANGELOG.md)
- [`visio-solid-viewer` 2.0.1](packages/solid/CHANGELOG.md)
- [`visio-vanilla-viewer` 2.0.1](packages/vanilla/CHANGELOG.md)
- Released with them: [`ooxml-core` 2.5.0](../../src/core/CHANGELOG.md) and [`ooxml-ui` 2.5.0](../../src/ui/CHANGELOG.md)

## 2026-11-21

<!-- release-notes: 2.0 -->

_Releases: visio-core@2.0.0, visio-react-viewer@2.0.0, visio-vue-viewer@2.0.0, visio-angular-viewer@2.0.0, visio-svelte-viewer@2.0.0, visio-solid-viewer@2.0.0, visio-vanilla-viewer@2.0.0, visio-viewer-mcp@2.0.0_

Every Visio package moves to 2.0.0 together, alongside `ooxml-core` 2.0.0 and `ooxml-ui` 2.0.0. Visio drawings now open and edit inside one memory budget sized from the host's heap, pages of up to about 10,000 shapes open as Visio saves them, and a drawing too large to edit opens read-only with the reason. The viewer gains a `locale` property, ribbon tabs a host can hide, F6 region navigation and smoother zooming. Visio stays in Beta. Upgrade every OOXML package at once; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) gives the before and after of each breaking change.

**[Introducing OOXML 2.0](https://christophervr.github.io/ooxml/releases/2.0/)** tells the story of this release. Every change to Visio is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.0/#changes-visio) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.0.md#changes-visio)).

### Packages

- [`visio-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/core/CHANGELOG.md)
- [`visio-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/react/CHANGELOG.md)
- [`visio-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/vue/CHANGELOG.md)
- [`visio-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/angular/CHANGELOG.md)
- [`visio-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/svelte/CHANGELOG.md)
- [`visio-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/solid/CHANGELOG.md)
- [`visio-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/vanilla/CHANGELOG.md)
- [`visio-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/mcp/CHANGELOG.md)

## 2026-11-16

_Releases: visio-react-viewer@1.4.1, visio-vue-viewer@1.4.1, visio-angular-viewer@1.4.1, visio-svelte-viewer@1.4.1, visio-solid-viewer@1.4.1, visio-vanilla-viewer@1.4.1_

### Dependencies

- **deps:** Take the open dependency updates in one change ([8eab074](https://github.com/ChristopherVR/ooxml/commit/8eab0749af462b944cda96a92e723132e9031f7a))

## 2026-11-13

_Releases: visio-core@1.1.2, visio-react-viewer@1.4.0, visio-vue-viewer@1.4.0, visio-angular-viewer@1.4.0, visio-svelte-viewer@1.4.0, visio-solid-viewer@1.4.0, visio-vanilla-viewer@1.4.0_

### Features

- **visio:** Expose an awaited inline text commit ([8e3a23b](https://github.com/ChristopherVR/ooxml/commit/8e3a23b58cb054f8ab51f0e74e0f681bea47d659))

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## 2026-11-09

_Releases: visio-react-viewer@1.3.2, visio-vue-viewer@1.3.2, visio-angular-viewer@1.3.2, visio-svelte-viewer@1.3.2, visio-solid-viewer@1.3.2, visio-vanilla-viewer@1.3.2_

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 25 updates ([4046ca3](https://github.com/ChristopherVR/ooxml/commit/4046ca399665d485da35696889c0212cc6526f63))

## 2026-11-06

_Releases: visio-core@1.1.1, visio-react-viewer@1.3.1, visio-vue-viewer@1.3.1, visio-angular-viewer@1.3.1, visio-svelte-viewer@1.3.1, visio-solid-viewer@1.3.1, visio-vanilla-viewer@1.3.1_

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## 2026-11-05

_Releases: visio-react-viewer@1.3.0, visio-vue-viewer@1.3.0, visio-angular-viewer@1.3.0, visio-svelte-viewer@1.3.0, visio-solid-viewer@1.3.0, visio-vanilla-viewer@1.3.0_

### Features

- **visio:** Open drawings from cloud file providers ([f7563e7](https://github.com/ChristopherVR/ooxml/commit/f7563e70c8fbdb2475d7b211b94bf278b590de8c))
- **ui:** Expose openCloudFile on the word, excel and visio bindings ([b4837c7](https://github.com/ChristopherVR/ooxml/commit/b4837c76398c125755458861ce62679d7c179003))

### Bug Fixes

- **deps:** Update emf-converter to 4.25.3 ([e51a619](https://github.com/ChristopherVR/ooxml/commit/e51a6192b3a531022b3677753e79b6a015233b9d))

## 2026-11-04

_Releases: visio-viewer-mcp@1.1.0, visio-core@1.1.0, visio-react-viewer@1.2.0, visio-vue-viewer@1.2.0, visio-angular-viewer@1.2.0, visio-svelte-viewer@1.2.0, visio-solid-viewer@1.2.0, visio-vanilla-viewer@1.2.0_

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

### Other

- Bring origin/main into the embedding work ([a3711a1](https://github.com/ChristopherVR/ooxml/commit/a3711a14049a2cd11b994e76eafec039db2dc2de))

## 2026-11-03

_Releases: visio-viewer-mcp@1.0.1, visio-react-viewer@1.1.3, visio-vue-viewer@1.1.3, visio-angular-viewer@1.1.3, visio-svelte-viewer@1.1.3, visio-solid-viewer@1.1.3, visio-vanilla-viewer@1.1.3_

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([8e60b14](https://github.com/ChristopherVR/ooxml/commit/8e60b14e1d8ee3a1cb507670caa15a611632ef26))
- **deps:** Merge production dependency updates ([#236](https://github.com/ChristopherVR/ooxml/issues/236)) ([7eb137b](https://github.com/ChristopherVR/ooxml/commit/7eb137b03906ded043ce0e4adf83e01fcdaa44a9))
- **deps:** Bump the angular group with 8 updates ([b233fe0](https://github.com/ChristopherVR/ooxml/commit/b233fe0bd00783876e01b5f0c7488a8f9e139fe6))
- **deps:** Merge angular dependency updates ([#234](https://github.com/ChristopherVR/ooxml/issues/234)) ([5c642cf](https://github.com/ChristopherVR/ooxml/commit/5c642cf7e6fc39265778fb062e4b13772af47a8e))

## 2026-10-26

_Releases: visio-react-viewer@1.1.2, visio-vue-viewer@1.1.2, visio-angular-viewer@1.1.2, visio-svelte-viewer@1.1.2, visio-solid-viewer@1.1.2, visio-vanilla-viewer@1.1.2_

### Dependencies

- **deps:** Bump production minor and patch dependencies ([#213](https://github.com/ChristopherVR/ooxml/issues/213)) ([3cfdd01](https://github.com/ChristopherVR/ooxml/commit/3cfdd016aa61cacb0540491c5ab13974e32f8e1e))

## 2026-10-23

_Releases: visio-react-viewer@1.1.1, visio-vue-viewer@1.1.1, visio-angular-viewer@1.1.1, visio-svelte-viewer@1.1.1, visio-solid-viewer@1.1.1, visio-vanilla-viewer@1.1.1_

### Refactor

- **visio:** Build the Visio bindings on ooxml-ui/bindings ([70e9c6c](https://github.com/ChristopherVR/ooxml/commit/70e9c6c77505ef97f9100942aa149cfe26f1cc63))

### Documentation

- **visio:** Describe the shared binding kit, new props and chart editing ([ef66cf6](https://github.com/ChristopherVR/ooxml/commit/ef66cf68620585b6ed3ea070d5690cae2a258121))

### Other

- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))
- Visio, Teams and package check docs (docs6-5) ([e6b2151](https://github.com/ChristopherVR/ooxml/commit/e6b2151824f447a53875592ee9f0a5f89f89cc69))

## 2026-10-22

_Releases: visio-viewer-mcp@1.0.0, visio-core@1.0.1, visio-core@1.0.0, visio-react-viewer@1.1.0, visio-react-viewer@1.0.1, visio-react-viewer@1.0.0, visio-vue-viewer@1.1.0, visio-vue-viewer@1.0.1, visio-vue-viewer@1.0.0, visio-angular-viewer@1.1.0, visio-angular-viewer@1.0.1, visio-angular-viewer@1.0.0, visio-svelte-viewer@1.1.0, visio-svelte-viewer@1.0.1, visio-svelte-viewer@1.0.0, visio-solid-viewer@1.1.0, visio-solid-viewer@1.0.1, visio-solid-viewer@1.0.0, visio-vanilla-viewer@1.1.0, visio-vanilla-viewer@1.0.1, visio-vanilla-viewer@1.0.0_

### Features

- **visio:** Complete the building blocks in every framework ([0a69177](https://github.com/ChristopherVR/ooxml/commit/0a6917719b5040587c0a37dfe3b37a875b7d2a05))
- **visio:** Add file-command and host-owned storage to the viewer ([0c8e284](https://github.com/ChristopherVR/ooxml/commit/0c8e284015fe1cc1b895c840900482ba12f8cda4))
- **visio:** Take the shared customization contract in <visio-viewer> ([3a6bb30](https://github.com/ChristopherVR/ooxml/commit/3a6bb305b8ac04f0942b40e8cdfbd887093c6112))
- **visio:** Accept readOnly as the inverse of canEdit ([fbc325b](https://github.com/ChristopherVR/ooxml/commit/fbc325be8c9c12a7c2e4a3affd7af2428ba4ed49))
- **embedding:** Land the host customization, read-only and storage contracts ([822a6a0](https://github.com/ChristopherVR/ooxml/commit/822a6a073840b9e9d1671476542322b28332dc42))
- **visio:** Take the shared collaboration contract in <visio-viewer> ([63ad160](https://github.com/ChristopherVR/ooxml/commit/63ad160c8ff5dc043695756f1bb9572fb19c2a4a))

### Bug Fixes

- **visio:** Ship the toolbar controls with the svelte package ([709fafa](https://github.com/ChristopherVR/ooxml/commit/709fafa524945526e613472143abf41054131332))
- **visio:** Reset readOnly when the host stops passing it ([c3db6eb](https://github.com/ChristopherVR/ooxml/commit/c3db6eb272b248329afbca32f70a72965f3f9d04))

### Refactor

- **mcp:** Register every format's tools through core automation ([9a28bfa](https://github.com/ChristopherVR/ooxml/commit/9a28bfa684167472fff27f04a371c37d494a5fd5))

### Documentation

- **visio:** Sync the viewer docs with the monorepo and current code ([2436dcc](https://github.com/ChristopherVR/ooxml/commit/2436dcce519728d28a65a124947b6dbbb5854cd6))

### Other

- Integrate the customization contract into host storage ([629554b](https://github.com/ChristopherVR/ooxml/commit/629554b3030728bfe16e11ac8ae5b2530bb28a17))
- Bring origin/main into the embedding work ([a3711a1](https://github.com/ChristopherVR/ooxml/commit/a3711a14049a2cd11b994e76eafec039db2dc2de))

## 2026-10-21

_Releases: visio-viewer-mcp@0.4.0, visio-core@0.5.0, visio-react-viewer@0.8.0, visio-vue-viewer@0.8.0, visio-angular-viewer@0.8.0, visio-svelte-viewer@0.8.0, visio-solid-viewer@0.8.0, visio-vanilla-viewer@0.8.0_

### Features

- **visio:** Add native framework building block integrations ([689e934](https://github.com/ChristopherVR/ooxml/commit/689e9340824a41e32a5024938e16b483e34c4452))
- **visio:** Expose canvas building blocks and host permissions ([e464a1c](https://github.com/ChristopherVR/ooxml/commit/e464a1c0399434430925e0e59b430b5788f5d8d9))
- **visio:** Theme native building blocks with shared office tokens ([f86b882](https://github.com/ChristopherVR/ooxml/commit/f86b882e2081540997f14af754805c573b372dc9))

### Bug Fixes

- **ui:** Preserve live formula drafts and runtime compatibility ([3a7d4f1](https://github.com/ChristopherVR/ooxml/commit/3a7d4f184658c4d69b04bbd36bab7afa3896dd9e))
- **visio:** Release replaced native canvas owners ([11b216f](https://github.com/ChristopherVR/ooxml/commit/11b216f6a2bfd36c91cd9747afea209a24617d29))

### Dependencies

- **deps:** Declare native adapter development dependencies ([b5f9957](https://github.com/ChristopherVR/ooxml/commit/b5f9957a65d934ebb4d40ed02bb872dcd24575fb))

### Chores

- Integrate latest main before full verification ([3bec814](https://github.com/ChristopherVR/ooxml/commit/3bec814857ae675f31eba14fcec59a3db4e1cb77))

## 2026-10-20

_Releases: visio-viewer-mcp@0.3.5, visio-react-viewer@0.7.4, visio-react-viewer@0.7.3, visio-vue-viewer@0.7.4, visio-vue-viewer@0.7.3, visio-angular-viewer@0.7.4, visio-angular-viewer@0.7.3, visio-svelte-viewer@0.7.4, visio-svelte-viewer@0.7.3, visio-solid-viewer@0.7.4, visio-solid-viewer@0.7.3, visio-vanilla-viewer@0.7.4, visio-vanilla-viewer@0.7.3_

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([02d7713](https://github.com/ChristopherVR/ooxml/commit/02d7713ceb301921fd6c7706138d64f4c10a9d38))
- **deps:** Merge dependency update ([#94](https://github.com/ChristopherVR/ooxml/issues/94)) ([e381aba](https://github.com/ChristopherVR/ooxml/commit/e381abaa11794233cdea85a9dc6d098d4cb338c4))
- **deps:** Update emf-converter to 4.21.0 ([#106](https://github.com/ChristopherVR/ooxml/issues/106)) ([641ae31](https://github.com/ChristopherVR/ooxml/commit/641ae316b7de07ed6f3c5158e8b154c991e1b719))

## 2026-10-15

_Releases: visio-viewer-mcp@0.3.4_

### Bug Fixes

- **automation:** Add --root and untrusted-content notes to MCP servers ([0da2c7d](https://github.com/ChristopherVR/ooxml/commit/0da2c7d1ec6655f649a608ded1f36dea41f34962))

## 2026-10-13

_Releases: visio-react-viewer@0.7.2, visio-vue-viewer@0.7.2, visio-angular-viewer@0.7.2, visio-svelte-viewer@0.7.2, visio-solid-viewer@0.7.2, visio-vanilla-viewer@0.7.2_

### Dependencies

- **deps:** Update production dependencies ([bbd7b53](https://github.com/ChristopherVR/ooxml/commit/bbd7b530c7a15ae48c0c0a2ff4ea7dcdfe2aa9d3))

## 2026-10-11

_Releases: visio-react-viewer@0.7.1, visio-vue-viewer@0.7.1, visio-angular-viewer@0.7.1, visio-svelte-viewer@0.7.1, visio-solid-viewer@0.7.1, visio-vanilla-viewer@0.7.1_

### Styling

- Fix the oxfmt and oxlint warnings in the viewers ([db0314f](https://github.com/ChristopherVR/ooxml/commit/db0314f319e3d0508960872260792215477d9ca5))

## 2026-10-10

_Releases: visio-viewer-mcp@0.3.3, visio-react-viewer@0.7.0, visio-vue-viewer@0.7.0, visio-angular-viewer@0.7.0, visio-svelte-viewer@0.7.0, visio-solid-viewer@0.7.0, visio-vanilla-viewer@0.7.0_

### Features

- **visio:** Make ribbon add-in tabs a prop and event of every binding ([6e6a157](https://github.com/ChristopherVR/ooxml/commit/6e6a157bbced977145d428e2ca01168e0a413deb))

### Styling

- Apply oxfmt to files drifted from the formatter ([daeabf2](https://github.com/ChristopherVR/ooxml/commit/daeabf26125af866f90359ac8a5716234d5b8b36))

## 2026-10-09

_Releases: visio-viewer-mcp@0.3.2, visio-react-viewer@0.6.2, visio-vue-viewer@0.6.2, visio-angular-viewer@0.6.2, visio-svelte-viewer@0.6.2, visio-solid-viewer@0.6.2, visio-vanilla-viewer@0.6.2_

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

## 2026-10-08

_Releases: visio-viewer-mcp@0.3.1, visio-viewer-mcp@0.3.0, visio-viewer-mcp@0.2.4, visio-core@0.4.1, visio-core@0.4.0, visio-core@0.3.7, visio-react-viewer@0.6.1, visio-react-viewer@0.6.0, visio-react-viewer@0.5.9, visio-vue-viewer@0.6.1, visio-vue-viewer@0.6.0, visio-vue-viewer@0.5.9, visio-angular-viewer@0.6.1, visio-angular-viewer@0.6.0, visio-angular-viewer@0.5.9, visio-svelte-viewer@0.6.1, visio-svelte-viewer@0.6.0, visio-svelte-viewer@0.5.9, visio-solid-viewer@0.6.1, visio-solid-viewer@0.6.0, visio-solid-viewer@0.5.9, visio-vanilla-viewer@0.6.1, visio-vanilla-viewer@0.6.0, visio-vanilla-viewer@0.5.9_

### Features

- **visio:** Enable new drawings and pointer selection gestures ([75ac409](https://github.com/ChristopherVR/ooxml/commit/75ac409e5e67c23203c765d3aa472b6817853090))
- **visio:** Enable source-backed formatting and shape ordering ([c39f7be](https://github.com/ChristopherVR/ooxml/commit/c39f7bee397d27afa31155a96c01997122402773))
- **visio:** Enable multi-selection formatting and arrangement ([fb70098](https://github.com/ChristopherVR/ooxml/commit/fb70098a2b100d6ec37cd091acd618c91477cfb2))
- **visio:** Enable rich formatting and source duplication ([d50e388](https://github.com/ChristopherVR/ooxml/commit/d50e388f9190809d6a860894663cdbf774ca4dca))
- **visio:** Enable source-backed shape clipboard ([2c13c1f](https://github.com/ChristopherVR/ooxml/commit/2c13c1f39998cf37e83d3949a6685f84f4b16e4d))

### Bug Fixes

- **ci:** Recover unpublished releases and validate manifests ([ae190df](https://github.com/ChristopherVR/ooxml/commit/ae190df2352283303858719fe1db5e58b86903fb))

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## 2026-10-07

_Releases: visio-viewer-mcp@0.2.3, visio-viewer-mcp@0.2.2, visio-core@0.3.6, visio-core@0.3.5, visio-core@0.3.4, visio-react-viewer@0.5.8, visio-react-viewer@0.5.7, visio-react-viewer@0.5.6, visio-vue-viewer@0.5.8, visio-vue-viewer@0.5.7, visio-vue-viewer@0.5.6, visio-angular-viewer@0.5.8, visio-angular-viewer@0.5.7, visio-angular-viewer@0.5.6, visio-svelte-viewer@0.5.8, visio-svelte-viewer@0.5.7, visio-svelte-viewer@0.5.6, visio-solid-viewer@0.5.8, visio-solid-viewer@0.5.7, visio-solid-viewer@0.5.6, visio-vanilla-viewer@0.5.8, visio-vanilla-viewer@0.5.7, visio-vanilla-viewer@0.5.6_

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))
- **visio:** Use extensionless relative source imports ([2495aa9](https://github.com/ChristopherVR/ooxml/commit/2495aa92d52a0135106695a2a39a72c9f98cf177))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## 2026-10-06

_Releases: visio-core@0.3.3, visio-react-viewer@0.5.5, visio-vue-viewer@0.5.5, visio-angular-viewer@0.5.5, visio-svelte-viewer@0.5.5, visio-solid-viewer@0.5.5, visio-vanilla-viewer@0.5.5_

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Dependencies

- **deps:** Align emf-converter on 4.8.21 ([3c2859d](https://github.com/ChristopherVR/ooxml/commit/3c2859d0a43e66666015debaeaa985888727db95))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## 2026-10-05

_Releases: visio-viewer-mcp@0.2.1, visio-core@0.3.2, visio-react-viewer@0.5.4, visio-vue-viewer@0.5.4, visio-angular-viewer@0.5.4, visio-svelte-viewer@0.5.4, visio-solid-viewer@0.5.4, visio-vanilla-viewer@0.5.4_

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))
- feat(bindings): expose native reactive viewer state in every binding (ccd9ba8)
- feat(viewer): add visio ribbon toolbar on shared office-ui controls (afc7047)
- docs(packages): align npm readmes with powerpoint structure (5c6f120)
- Initial release.
- build(deps): adopt ooxml-core 0.15.0 and ooxml-ui 0.21.0 (818f4a4)
- feat(viewer): add visio account, options dialog and live sharing (fe7fed1)
- fix(packages): pin released core and shared ui dependencies (4d035f6)
- fix(deps): use released core for proven master shape moves (28989be)
- feat(release): publish functional visio framework packages (bc55722)
- build(viewer): adopt ooxml-ui 0.18.2 (44f44e2)
- build(viewer): adopt ooxml-ui 0.17 (2e19d89)
- build(viewer): adopt ooxml-ui 0.15 shared select and ribbon commands (e1fe065)
- build(viewer): adopt token based ooxml-ui 0.13 (590a5c8)
- refactor(viewer): use shared ooxml-ui ribbon, backstage, find and rulers (cd1088b)
- feat(viewer): add pan & zoom window and a shapes strip to reopen shapes (01c5a82)
- feat(viewer): add visio keytips for keyboard ribbon access (8a9d6d1)
- feat(viewer): add visio tell me, help tab and presentation mode (d048a45)
- feat(viewer): add visio file backstage, context menus and page bar (8c54263)
- feat(viewer): add visio's shapes window with a basic shapes stencil (98cb55e)
- feat(viewer): match the visio ribbon layout with honest disabled commands (bb74402)
- feat(visio): preview legacy VSD through the shared binary codec (6772ff3)
- fix(ui): match suite Pages and responsive editor composition (960d3bc)
- feat(viewer): preserve drafts across geometry edits and history (c5b585d)
- fix(viewer): align Office shell and Pages navigation (23d6912)
- feat(viewer): expose safe geometry commands and pin published core (8ddaf87)
- fix(viewer): follow the shared office theme inside frames (caad674)
- feat(ui): align Visio workspace and docs with PPTX design (011c728)
- feat(visio): restore concave arrows and rounded connectors (7662a75)
