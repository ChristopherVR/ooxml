# Changelog

All notable changes to this viewer's packages, newest first. Each section names the packages
released that day; the full history of each package is in its own `CHANGELOG.md`
(`packages/<name>/`, `mcp/`, `server/`).

## 2026-11-27

<!-- release-notes: 2.5 -->

_Releases: xlsx-viewer-mcp@2.0.1, @christophervr/xlsx-core@2.0.1, @christophervr/xlsx-react-viewer@2.1.0, xlsx-vue-viewer@2.1.0, xlsx-angular-viewer@2.1.0, xlsx-svelte-viewer@2.1.0, xlsx-solid-viewer@2.1.0, xlsx-vanilla-viewer@2.1.0_

Every Excel package is released with `ooxml-core` 2.5.0 and `ooxml-ui` 2.5.0: the six bindings move to 2.1.0, and `@christophervr/xlsx-core` and `xlsx-viewer-mcp` to 2.0.1. The bindings mount the 2.5 Excel editor, with Recommended PivotTables, PivotCharts, the PivotTable Analyze and Design tabs, the Chart Format tab, Picture Format and Shape Format, the Draw tab with ink saved as Excel ink, Place in Cell and linked data cards, and more of Insert and Page Layout; PivotTables, in-cell pictures and ink are shared in collaboration. They take a new optional `imageProvider` prop for `IMAGE()`, which fetches nothing until you set it, and forward `googleFonts` in every binding. Non-English interface languages now load on first use. Upgrade every OOXML package together; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.5.md#excel-command-string-tables-moved-behind-loadeditorstrings) covers the string tables, the Insert ribbon and the shared workbook changes.

**[Introducing OOXML 2.5](https://christophervr.github.io/ooxml/releases/2.5/)** tells the story of this release. Every change to Excel is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.5/#changes-excel) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.5.md#changes-excel)).

### Packages

- [`xlsx-viewer-mcp` 2.0.1](mcp/CHANGELOG.md)
- [`@christophervr/xlsx-core` 2.0.1](packages/core/CHANGELOG.md)
- [`@christophervr/xlsx-react-viewer` 2.1.0](packages/react/CHANGELOG.md)
- [`xlsx-vue-viewer` 2.1.0](packages/vue/CHANGELOG.md)
- [`xlsx-angular-viewer` 2.1.0](packages/angular/CHANGELOG.md)
- [`xlsx-svelte-viewer` 2.1.0](packages/svelte/CHANGELOG.md)
- [`xlsx-solid-viewer` 2.1.0](packages/solid/CHANGELOG.md)
- [`xlsx-vanilla-viewer` 2.1.0](packages/vanilla/CHANGELOG.md)
- Released with them: [`ooxml-core` 2.5.0](../../src/core/CHANGELOG.md) and [`ooxml-ui` 2.5.0](../../src/ui/CHANGELOG.md)

## 2026-11-21

<!-- release-notes: 2.0 -->

_Releases: @christophervr/xlsx-core@2.0.0, @christophervr/xlsx-react-viewer@2.0.0, xlsx-vue-viewer@2.0.0, xlsx-angular-viewer@2.0.0, xlsx-svelte-viewer@2.0.0, xlsx-solid-viewer@2.0.0, xlsx-vanilla-viewer@2.0.0, xlsx-viewer-mcp@2.0.0_

Every Excel package moves to 2.0.0 together, with `ooxml-core` and `ooxml-ui` 2.0.0. The editor is safer against crafted workbooks and formulas, faster on large sheets, frees itself when closed, and is checked for keyboard and screen reader use before every release. The framework bindings keep their props, events and handles. Upgrade every OOXML package you use in one change; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) has the before and after for each breaking change.

**[Introducing OOXML 2.0](https://christophervr.github.io/ooxml/releases/2.0/)** tells the story of this release. Every change to Excel is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.0/#changes-excel) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.0.md#changes-excel)).

### Packages

- [`@christophervr/xlsx-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/core/CHANGELOG.md)
- [`@christophervr/xlsx-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/react/CHANGELOG.md)
- [`xlsx-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/vue/CHANGELOG.md)
- [`xlsx-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/angular/CHANGELOG.md)
- [`xlsx-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/svelte/CHANGELOG.md)
- [`xlsx-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/solid/CHANGELOG.md)
- [`xlsx-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/vanilla/CHANGELOG.md)
- [`xlsx-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/CHANGELOG.md)

## 2026-11-13

_Releases: @christophervr/xlsx-core@1.1.2, @christophervr/xlsx-react-viewer@1.5.0, xlsx-vue-viewer@1.5.0, xlsx-angular-viewer@1.5.0, xlsx-svelte-viewer@1.5.0, xlsx-solid-viewer@1.5.0, xlsx-vanilla-viewer@1.5.0_

### Features

- **xlsx:** Host-owned context menus and workflow operations ([18854c4](https://github.com/ChristopherVR/ooxml/commit/18854c444095dfaea70f087f3c868211316436a7))
- **xlsx:** Filter and extend the native context menu ([554d354](https://github.com/ChristopherVR/ooxml/commit/554d354d5f2ec8ce9146ce43645301b379f71d11))
- **xlsx:** Make replace validation configurable ([093f478](https://github.com/ChristopherVR/ooxml/commit/093f4782640506ab8274134f70ad4abb8a2e4e50))

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## 2026-11-06

_Releases: @christophervr/xlsx-core@1.1.1, @christophervr/xlsx-react-viewer@1.4.1, xlsx-vue-viewer@1.4.1, xlsx-angular-viewer@1.4.1, xlsx-svelte-viewer@1.4.1, xlsx-solid-viewer@1.4.1, xlsx-vanilla-viewer@1.4.1_

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## 2026-11-05

_Releases: @christophervr/xlsx-react-viewer@1.4.0, xlsx-vue-viewer@1.4.0, xlsx-angular-viewer@1.4.0, xlsx-svelte-viewer@1.4.0, xlsx-solid-viewer@1.4.0, xlsx-vanilla-viewer@1.4.0_

### Features

- **xlsx:** Open workbooks from cloud places in File > Open ([f1b19e4](https://github.com/ChristopherVR/ooxml/commit/f1b19e495f8ac1fe43ed1a540f76ce5c6cbc2710))
- **ui:** Expose openCloudFile on the word, excel and visio bindings ([b4837c7](https://github.com/ChristopherVR/ooxml/commit/b4837c76398c125755458861ce62679d7c179003))

## 2026-11-04

_Releases: xlsx-viewer-mcp@1.1.0, @christophervr/xlsx-core@1.1.0, @christophervr/xlsx-react-viewer@1.3.0, xlsx-vue-viewer@1.3.0, xlsx-angular-viewer@1.3.0, xlsx-svelte-viewer@1.3.0, xlsx-solid-viewer@1.3.0, xlsx-vanilla-viewer@1.3.0_

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

### Dependencies

- **deps:** Bump the angular group with 8 updates ([b233fe0](https://github.com/ChristopherVR/ooxml/commit/b233fe0bd00783876e01b5f0c7488a8f9e139fe6))
- **deps:** Merge angular dependency updates ([#234](https://github.com/ChristopherVR/ooxml/issues/234)) ([5c642cf](https://github.com/ChristopherVR/ooxml/commit/5c642cf7e6fc39265778fb062e4b13772af47a8e))

## 2026-11-03

_Releases: xlsx-viewer-mcp@1.0.1_

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([8e60b14](https://github.com/ChristopherVR/ooxml/commit/8e60b14e1d8ee3a1cb507670caa15a611632ef26))
- **deps:** Merge production dependency updates ([#236](https://github.com/ChristopherVR/ooxml/issues/236)) ([7eb137b](https://github.com/ChristopherVR/ooxml/commit/7eb137b03906ded043ce0e4adf83e01fcdaa44a9))

## 2026-10-29

_Releases: xlsx-viewer-mcp@1.0.0, @christophervr/xlsx-core@1.0.0, @christophervr/xlsx-react-viewer@1.2.2, xlsx-vue-viewer@1.2.2, xlsx-angular-viewer@1.2.2, xlsx-svelte-viewer@1.2.2, xlsx-solid-viewer@1.2.2, xlsx-vanilla-viewer@1.2.2_

### Features

- **xlsx:** Commit xlsx-core and xlsx-viewer-mcp to a stable 1.0 api ([e458ecf](https://github.com/ChristopherVR/ooxml/commit/e458ecf368421096ccffb366cc1e0ba17f438b77))

## 2026-10-23

_Releases: @christophervr/xlsx-core@0.3.1, @christophervr/xlsx-react-viewer@1.2.1, xlsx-vue-viewer@1.2.1, xlsx-angular-viewer@1.2.1, xlsx-svelte-viewer@1.2.1, xlsx-solid-viewer@1.2.1, xlsx-vanilla-viewer@1.2.1_

### Refactor

- **docx:** Build the Word bindings on ooxml-ui/bindings ([90e8131](https://github.com/ChristopherVR/ooxml/commit/90e81310800081ad484e93a3bab25806a07bc3e4))
- **xlsx:** Build the Excel bindings on ooxml-ui/bindings ([31ca812](https://github.com/ChristopherVR/ooxml/commit/31ca812ba4513f0b0e0db643ab4b36643cbe62e5))

### Documentation

- Describe wave 6 chart and smartart editing in word and excel ([d212908](https://github.com/ChristopherVR/ooxml/commit/d212908f59be7f5da51ce5347b23a665e4cedb08))

### Other

- Word and Excel docs for wave 6 (docs6-4) ([da0f3e8](https://github.com/ChristopherVR/ooxml/commit/da0f3e83c59b8ff55104bd852cf46fa052c587b5))
- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))

## 2026-10-22

_Releases: xlsx-viewer-mcp@0.2.2, xlsx-viewer-mcp@0.2.1, @christophervr/xlsx-react-viewer@1.2.0, @christophervr/xlsx-react-viewer@1.1.0, @christophervr/xlsx-react-viewer@1.0.0, xlsx-vue-viewer@1.2.0, xlsx-vue-viewer@1.1.0, xlsx-vue-viewer@1.0.0, xlsx-angular-viewer@1.2.0, xlsx-angular-viewer@1.1.0, xlsx-angular-viewer@1.0.0, xlsx-svelte-viewer@1.2.0, xlsx-svelte-viewer@1.1.0, xlsx-svelte-viewer@1.0.0, xlsx-solid-viewer@1.2.0, xlsx-solid-viewer@1.1.0, xlsx-solid-viewer@1.0.0, xlsx-vanilla-viewer@1.2.0, xlsx-vanilla-viewer@1.1.0, xlsx-vanilla-viewer@1.0.0_

### Features

- **xlsx:** Route save and open through host-owned storage ([ef768a6](https://github.com/ChristopherVR/ooxml/commit/ef768a684f558275071be347d68406503aef8c23))
- **xlsx:** Forward file-command, workbook, ribbon and sheet events ([5aaf6db](https://github.com/ChristopherVR/ooxml/commit/5aaf6dbdc767a1ff8c30b1b8619c0dcb09cfdc38))
- **xlsx:** Take the shared customization contract in <xlsx-editor> ([f96a544](https://github.com/ChristopherVR/ooxml/commit/f96a5449a6a36fba59997197b484ddefb4f2d8e8))
- **embedding:** Land the host customization, read-only and storage contracts ([822a6a0](https://github.com/ChristopherVR/ooxml/commit/822a6a073840b9e9d1671476542322b28332dc42))
- **xlsx:** Let hosts own the data-validation alert in building blocks ([fcbe367](https://github.com/ChristopherVR/ooxml/commit/fcbe367f5cfe366cc425796e3fdff9c31aa99e25))
- **xlsx:** Complete the building-block toolbar and add vanilla ([af8114f](https://github.com/ChristopherVR/ooxml/commit/af8114ff92e023858cd5a6d2a9922d6dbb2ee111))
- **xlsx:** Make the building blocks read-only unless canEdit is true ([efe6158](https://github.com/ChristopherVR/ooxml/commit/efe6158de251e1e34694335292755cfec96d640b))

### Bug Fixes

- **xlsx:** Report read-only with the lock and host switch ([782adb9](https://github.com/ChristopherVR/ooxml/commit/782adb95a4c8ad4dded13d73e8a5fa5aee2c8666))
- **xlsx:** Name the share opener like the other building blocks ([1346849](https://github.com/ChristopherVR/ooxml/commit/1346849e085724ef8a4810cf540f8d802fd47cee))
- **xlsx:** Expect the readonly-change detail in the vanilla mount test ([d8cf911](https://github.com/ChristopherVR/ooxml/commit/d8cf9116eb28ff440538491a9af459b1481f4afc))

### Refactor

- **mcp:** Register every format's tools through core automation ([9a28bfa](https://github.com/ChristopherVR/ooxml/commit/9a28bfa684167472fff27f04a371c37d494a5fd5))

### Documentation

- **xlsx:** Document building blocks and list them in the sidebar ([9d48799](https://github.com/ChristopherVR/ooxml/commit/9d4879997eaea18e71ff61203c5a52c2b33fa4c0))

### Testing

- **xlsx:** Run the reviewed-edit tools end to end over mcp ([9f66379](https://github.com/ChristopherVR/ooxml/commit/9f66379b96ec2708d032294d1ef2f8557e24f284))

### Styling

- **xlsx:** Format the react binding test ([a2e8b70](https://github.com/ChristopherVR/ooxml/commit/a2e8b70fe6f25b51b02cbd612675f492f247b2fc))

### Other

- Integrate the customization contract into host storage ([629554b](https://github.com/ChristopherVR/ooxml/commit/629554b3030728bfe16e11ac8ae5b2530bb28a17))

## 2026-10-21

_Releases: xlsx-viewer-mcp@0.2.0, @christophervr/xlsx-core@0.3.0, @christophervr/xlsx-react-viewer@0.7.0, xlsx-vue-viewer@0.7.0, xlsx-angular-viewer@0.7.0, xlsx-svelte-viewer@0.7.0, xlsx-solid-viewer@0.7.0, xlsx-vanilla-viewer@0.7.0_

### Features

- **xlsx:** Add native framework building blocks ([9689a65](https://github.com/ChristopherVR/ooxml/commit/9689a6520682e9db18f8cbee78d8434ad114b097))
- **xlsx:** Expose composable spreadsheet building blocks ([9aef499](https://github.com/ChristopherVR/ooxml/commit/9aef499f76196be6dd008ae873ac8f0d4caeeb58))
- **xlsx:** Share office themes across editor building blocks ([a791ec2](https://github.com/ChristopherVR/ooxml/commit/a791ec2e68956921006c719397803e603286a8c0))

### Bug Fixes

- **deps:** Align private bindings with their react test host ([f4152ef](https://github.com/ChristopherVR/ooxml/commit/f4152ef04ccdc3e5c9d21ecf59f485b56de2ec52))
- **ui:** Harden composition lifecycle and consumer checks ([6aa52be](https://github.com/ChristopherVR/ooxml/commit/6aa52be88175ca4f46f6f3b1c223ef06a8f1f915))
- **ui:** Preserve live formula drafts and runtime compatibility ([3a7d4f1](https://github.com/ChristopherVR/ooxml/commit/3a7d4f184658c4d69b04bbd36bab7afa3896dd9e))
- **ui:** Guard framework demos and native input identity ([a094008](https://github.com/ChristopherVR/ooxml/commit/a094008b00578fff9ed861addbfe2aca2ce9ab09))

### Styling

- Fix the oxfmt and oxlint warnings in the viewers ([db0314f](https://github.com/ChristopherVR/ooxml/commit/db0314f319e3d0508960872260792215477d9ca5))

### Dependencies

- **deps:** Declare native adapter development dependencies ([b5f9957](https://github.com/ChristopherVR/ooxml/commit/b5f9957a65d934ebb4d40ed02bb872dcd24575fb))

## 2026-10-20

_Releases: xlsx-viewer-mcp@0.1.11_

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([02d7713](https://github.com/ChristopherVR/ooxml/commit/02d7713ceb301921fd6c7706138d64f4c10a9d38))
- **deps:** Merge dependency update ([#94](https://github.com/ChristopherVR/ooxml/issues/94)) ([e381aba](https://github.com/ChristopherVR/ooxml/commit/e381abaa11794233cdea85a9dc6d098d4cb338c4))

## 2026-10-15

_Releases: xlsx-viewer-mcp@0.1.10_

### Bug Fixes

- **automation:** Add --root and untrusted-content notes to MCP servers ([0da2c7d](https://github.com/ChristopherVR/ooxml/commit/0da2c7d1ec6655f649a608ded1f36dea41f34962))

## 2026-10-10

_Releases: xlsx-viewer-mcp@0.1.9, @christophervr/xlsx-react-viewer@0.6.0, xlsx-vue-viewer@0.6.0, xlsx-angular-viewer@0.6.0, xlsx-svelte-viewer@0.6.0, xlsx-solid-viewer@0.6.0, xlsx-vanilla-viewer@0.6.0_

### Features

- **xlsx:** Make ribbon add-in tabs a prop and event of every binding ([691f553](https://github.com/ChristopherVR/ooxml/commit/691f5531b9f5b4a5a5ac0645b43c8c0fc55be5d8))

### Styling

- Apply oxfmt to files drifted from the formatter ([daeabf2](https://github.com/ChristopherVR/ooxml/commit/daeabf26125af866f90359ac8a5716234d5b8b36))

## 2026-10-09

_Releases: xlsx-viewer-mcp@0.1.8_

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

## 2026-10-08

_Releases: xlsx-viewer-mcp@0.1.7, xlsx-viewer-mcp@0.1.6, @christophervr/xlsx-core@0.2.10, @christophervr/xlsx-core@0.2.9, @christophervr/xlsx-react-viewer@0.5.1, @christophervr/xlsx-react-viewer@0.5.0, @christophervr/xlsx-react-viewer@0.4.6, xlsx-vue-viewer@0.5.1, xlsx-vue-viewer@0.5.0, xlsx-vue-viewer@0.4.6, xlsx-angular-viewer@0.5.1, xlsx-angular-viewer@0.5.0, xlsx-angular-viewer@0.4.6, xlsx-svelte-viewer@0.5.1, xlsx-svelte-viewer@0.5.0, xlsx-svelte-viewer@0.4.6, xlsx-solid-viewer@0.5.1, xlsx-solid-viewer@0.5.0, xlsx-solid-viewer@0.4.6, xlsx-vanilla-viewer@0.5.1, xlsx-vanilla-viewer@0.5.0, xlsx-vanilla-viewer@0.4.6_

### Features

- **xlsx:** Expose collaboration in every binding ([2983ccc](https://github.com/ChristopherVR/ooxml/commit/2983ccc1e5d72749a31d2540e8f57b637f282b3e))

### Bug Fixes

- **ci:** Recover unpublished releases and validate manifests ([ae190df](https://github.com/ChristopherVR/ooxml/commit/ae190df2352283303858719fe1db5e58b86903fb))

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## 2026-10-07

_Releases: xlsx-viewer-mcp@0.1.5, @christophervr/xlsx-core@0.2.8, @christophervr/xlsx-core@0.2.7, @christophervr/xlsx-core@0.2.6, @christophervr/xlsx-react-viewer@0.4.5, @christophervr/xlsx-react-viewer@0.4.4, @christophervr/xlsx-react-viewer@0.4.3, xlsx-vue-viewer@0.4.5, xlsx-vue-viewer@0.4.4, xlsx-vue-viewer@0.4.3, xlsx-angular-viewer@0.4.5, xlsx-angular-viewer@0.4.4, xlsx-angular-viewer@0.4.3, xlsx-svelte-viewer@0.4.5, xlsx-svelte-viewer@0.4.4, xlsx-svelte-viewer@0.4.3, xlsx-solid-viewer@0.4.5, xlsx-solid-viewer@0.4.4, xlsx-solid-viewer@0.4.3, xlsx-vanilla-viewer@0.4.5, xlsx-vanilla-viewer@0.4.4, xlsx-vanilla-viewer@0.4.3_

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## 2026-10-06

_Releases: @christophervr/xlsx-core@0.2.5, @christophervr/xlsx-react-viewer@0.4.2, xlsx-vue-viewer@0.4.2, xlsx-angular-viewer@0.4.2, xlsx-svelte-viewer@0.4.2, xlsx-solid-viewer@0.4.2, xlsx-vanilla-viewer@0.4.2_

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))
- **demos:** Move the viewer demos to demos/<product> ([4031a96](https://github.com/ChristopherVR/ooxml/commit/4031a96e6d4a20192ddd263858bdd2c20edccb54))
- **e2e:** Move the viewer browser tests to e2e/<product> ([2e4106d](https://github.com/ChristopherVR/ooxml/commit/2e4106d2f240f3440782faf93e711e9eaef06f50))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## 2026-10-05

_Releases: @christophervr/xlsx-core@0.2.3, @christophervr/xlsx-react-viewer@0.4.0, xlsx-vue-viewer@0.4.0, xlsx-angular-viewer@0.4.0, xlsx-svelte-viewer@0.4.0, xlsx-solid-viewer@0.4.0, xlsx-vanilla-viewer@0.4.0_

### Features

- **web-component:** Add iterative calculation to File options (by @ChristopherVR) ([d5f1e4c](https://github.com/ChristopherVR/xlsx-viewer/commit/d5f1e4cc7d0feb0fc7847599bfaba27a46ee1f47))

### Bug Fixes

- **web-component:** Ask for the password when only the SHA-512 hash is set (by @ChristopherVR) ([55bd264](https://github.com/ChristopherVR/xlsx-viewer/commit/55bd264ebeb92b470df4a48e639327e68dcb85d8))

### Refactor

- **web-component:** Draw the title bar, status bar and File view with shared elements (by @ChristopherVR) ([6dc3085](https://github.com/ChristopherVR/xlsx-viewer/commit/6dc3085a6b642a5036eedd09098f04c192da84e2))
- **web-component:** Draw the tab row, File button and collapse with office-ui-ribbon (by @ChristopherVR) ([c15e3db](https://github.com/ChristopherVR/xlsx-viewer/commit/c15e3db64eb32f4d0b38ae024dad850952e73df1))
- **web-component:** Fold ribbon groups with the shared createRibbonOverflow (by @ChristopherVR) ([258045b](https://github.com/ChristopherVR/xlsx-viewer/commit/258045b9a332786eac72ace14b84727e944ffbee))
- **web-component:** Use the shared KeyTip assignment and badges (by @ChristopherVR) ([e022e06](https://github.com/ChristopherVR/xlsx-viewer/commit/e022e06eec847d0842104ca01365ac32078e2da0))
- **web-component:** Import the formula text helpers from ooxml-core (by @ChristopherVR) ([be896de](https://github.com/ChristopherVR/xlsx-viewer/commit/be896de75c8d13d5879f7debfdd320bdd4336091))

### Documentation

- **site:** Add live demos page and publish install commands (by @ChristopherVR) ([cb35517](https://github.com/ChristopherVR/xlsx-viewer/commit/cb35517541e726f72e4686412c666e4699bbb633))

### Build & CI

- **deps:** Adopt ooxml-core 0.17.0 and ooxml-ui 0.25.0 (by @ChristopherVR) ([49705cd](https://github.com/ChristopherVR/xlsx-viewer/commit/49705cd85cae5c14275ef15a544c689a3867abe4))
- **deps:** Adopt ooxml-ui 0.26.0 (by @ChristopherVR) ([94984f9](https://github.com/ChristopherVR/xlsx-viewer/commit/94984f959c1c75265294e5c49e3c22071e1f54f4))
- **deps:** Adopt ooxml-ui 0.27.0 (by @ChristopherVR) ([f24314b](https://github.com/ChristopherVR/xlsx-viewer/commit/f24314b56e662a05c6095338f32779536cc0d465))
- **deps:** Adopt ooxml-core 0.18.0 and ooxml-ui 0.27.1 (by @ChristopherVR) ([6e66a43](https://github.com/ChristopherVR/xlsx-viewer/commit/6e66a43e8966eedc4bd19934e27085be76d33a95))
- **deps:** Adopt ooxml-core 0.19.0 (by @ChristopherVR) ([d981768](https://github.com/ChristopherVR/xlsx-viewer/commit/d981768e1f6c5f612ea6c69aaf9507933a5a694f))

## 2026-10-05

_Releases: @christophervr/xlsx-core@0.2.2, @christophervr/xlsx-react-viewer@0.3.4, xlsx-vue-viewer@0.3.4, xlsx-angular-viewer@0.3.4, xlsx-svelte-viewer@0.3.4, xlsx-solid-viewer@0.3.4, xlsx-vanilla-viewer@0.3.4_

### Refactor

- **web-component:** Use the shared office-ui-context-menu (by @ChristopherVR) ([6aa323c](https://github.com/ChristopherVR/xlsx-viewer/commit/6aa323ce664715552c7e35da4c8ef6611b1d1bc4))

### Build & CI

- **deps:** Adopt ooxml-core 0.14.1 and ooxml-ui 0.20 (by @ChristopherVR) ([2da396d](https://github.com/ChristopherVR/xlsx-viewer/commit/2da396d66cad5cb74915d665258ac8216c3cd602))
- **deps:** Adopt ooxml-core 0.15.0 and ooxml-ui 0.21.0 (by @ChristopherVR) ([6cbd2b5](https://github.com/ChristopherVR/xlsx-viewer/commit/6cbd2b5ad65734a3d9ce7f2c8feff93d975eafe5))
- Follow ooxml-core and ooxml-ui releases automatically (by @ChristopherVR) ([41a587d](https://github.com/ChristopherVR/xlsx-viewer/commit/41a587d0d4b607734b67c9cc87351524d40957e1))

## 2026-10-03

_Releases: xlsx-viewer-mcp@0.1.3, @christophervr/xlsx-core@0.2.1, @christophervr/xlsx-react-viewer@0.3.3, xlsx-vue-viewer@0.3.3, xlsx-angular-viewer@0.3.3, xlsx-svelte-viewer@0.3.3, xlsx-solid-viewer@0.3.3, xlsx-vanilla-viewer@0.3.3_

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([afbbf9c](https://github.com/ChristopherVR/xlsx-viewer/commit/afbbf9c9a88f70e0afeec47da60e6cd9339c76ba))

### Chores

- **repo:** Remove Codex co-author trailers (by @ChristopherVR) ([f594bc2](https://github.com/ChristopherVR/xlsx-viewer/commit/f594bc238ccfb180bf1599abffde1093d3b08711))

## 2026-10-03

_Releases: xlsx-viewer-mcp@0.1.2, @christophervr/xlsx-react-viewer@0.3.2, xlsx-vue-viewer@0.3.2, xlsx-angular-viewer@0.3.2, xlsx-svelte-viewer@0.3.2, xlsx-solid-viewer@0.3.2, xlsx-vanilla-viewer@0.3.2_

### Bug Fixes

- **release:** Stage planned package versions and changelogs (by @ChristopherVR) ([66e7c29](https://github.com/ChristopherVR/xlsx-viewer/commit/66e7c29b4c8d5109c358ac1df77a4afad91a86f1))

## 2026-10-03

_Releases: @christophervr/xlsx-core@0.2.0, @christophervr/xlsx-react-viewer@0.3.0, xlsx-vue-viewer@0.3.0, xlsx-angular-viewer@0.3.0, xlsx-svelte-viewer@0.3.0, xlsx-solid-viewer@0.3.0, xlsx-vanilla-viewer@0.3.0_

### Features

- **web-component:** Adopt core encryption and workbook features (by @ChristopherVR) ([857348f](https://github.com/ChristopherVR/xlsx-viewer/commit/857348fb817ce99d773c8d4dad6258ff1dc93384))

## 2026-10-03

_Releases: @christophervr/xlsx-core@0.1.2, @christophervr/xlsx-react-viewer@0.2.0, xlsx-vue-viewer@0.2.0, xlsx-angular-viewer@0.2.0, xlsx-svelte-viewer@0.2.0, xlsx-solid-viewer@0.2.0, xlsx-vanilla-viewer@0.2.0_

### Features

- **bindings:** Keep in-editor changes and add readOnly/ribbon callbacks (by @ChristopherVR) ([6196ce1](https://github.com/ChristopherVR/xlsx-viewer/commit/6196ce10a9145bb1a409bd694086b772740045fa))

### Bug Fixes

- **web-component:** Restore sheet selections and fire selection once (by @ChristopherVR) ([fdfe205](https://github.com/ChristopherVR/xlsx-viewer/commit/fdfe205cd0a3bdb3ad575447a5bd842591eae942))
- **deps:** Depend on ooxml-core ^0.9.0 and ooxml-ui ^0.1.8 (by @ChristopherVR) ([e2a8ae6](https://github.com/ChristopherVR/xlsx-viewer/commit/e2a8ae63533262f4fef4fb8c93feadd3ca3b15b8))

## 2026-10-03

_Releases: @christophervr/xlsx-core@0.1.1, @christophervr/xlsx-react-viewer@0.1.1, xlsx-vue-viewer@0.1.1, xlsx-angular-viewer@0.1.1, xlsx-svelte-viewer@0.1.1, xlsx-solid-viewer@0.1.1, xlsx-vanilla-viewer@0.1.1_

### Features

- Add the Excel viewer and editor (by @ChristopherVR) ([408fe71](https://github.com/ChristopherVR/xlsx-viewer/commit/408fe71adc283e887cca4e873673daa6934957d5))

### Testing

- Assert the currency format, not font-dependent text (by @ChristopherVR) ([8fe9421](https://github.com/ChristopherVR/xlsx-viewer/commit/8fe94212c09b906ea786e370e55fd9bd8c6b78ef))
