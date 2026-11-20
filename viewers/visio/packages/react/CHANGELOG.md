# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

For React apps, `<VisioViewer>` now opens drawings with pages of up to about 10,000 shapes inside a memory budget sized from the browser, refuses crafted files that would exhaust it, and opens a drawing too large to edit read-only with the reason in the reader's language. A `locale` prop sets the interface language, and `customization.ribbon.hiddenTabs` removes ribbon tabs, File included. Two changes can break an existing app: the viewer takes dark mode from `--office-color-scheme` or `data-office-theme` instead of the page's CSS `color-scheme`, and Node.js 22.3 is the minimum for builds and server rendering. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) gives the before and after of each breaking change.

### Breaking changes

- The viewer no longer inherits the page's CSS `color-scheme`. It reads the `--office-color-scheme` token, which is `light` unless `data-office-theme="dark"` applies, or the system prefers dark and no `data-office-theme="light"` overrides it. A container that sets only `color-scheme: dark` now gets a light viewer: add `--office-color-scheme: dark` to it, or use `data-office-theme="dark"`, which sets both ([guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#the-editors-take-dark-mode-from-the-theme-not-from-color-scheme)).
- The package requires Node.js 22.3 or newer and declares `"engines": { "node": ">=22.3" }`. `ooxml-core` sizes the memory budget of a Visio read or edit from the heap limit Node reports, which it reaches through an API added in 22.3. Raise the Node version of your build machines, servers and CI ([guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#run-nodejs-223-or-newer)).
- A collaboration update that holds more than one struct per 4 bytes, or more than one deleted range per 2 bytes, is refused before it is decoded, and a refusal stops the sync. The budget grows with the update, so the first sync of a real drawing stays well inside it. Code that calls `ooxml-core/collab` directly replaces `DEFAULT_UPDATE_LIMITS` with `defaultUpdateLimits(bytes)` ([guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#collaboration-updates-are-bounded-and-a-refusal-stops-the-sync)).
- The status bar's zoom button is named by the percentage it shows, for example "100%, Zoom to fit". Browser tests that find it by the name "Zoom to fit" use `[data-id="zoomFit"]` instead ([guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#tests-find-the-zoom-button-by-its-new-name)).
- `emf-converter` 5 loads its image decoders and colour data on demand. Package managers and bundlers need no change; if you copy compiled files by hand, deploy `dist/browser-chunks/` and `dist/node-chunks/` next to the entry files ([guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#deploy-the-new-decoder-folders-if-you-copy-files-by-hand)).

### Security

- No Visio file can exhaust the memory of the page or worker that reads or edits it. Each read or edit has one memory budget, sized from the host's heap and estimated before each XML part is parsed. A drawing that does not fit is refused with the code `LIMIT_XML_MEMORY` and a message in the viewer's language saying how much memory it would need; a small file of dense XML is refused in well under a second.
- Every package is opened through the guarded zip reader, including collaboration snapshots, signature listing and embedded chart workbooks, so a small file can no longer inflate without limit.
- Chart point counts and point indices no longer decide how much memory is allocated, so a chart part of a few hundred bytes can no longer exhaust memory.
- SVG pictures are shown from `data:` URLs instead of object URLs of your origin, so opening one in its own tab can no longer run document-controlled script there.
- A collaborator can no longer replace an arbitrary package part through shared media or add markup to the content types part, and the values peers send are checked before they are applied.
- A drawing is shown as signed only when its signature verifies, never because it merely contains a signature part.
- The package is published by a release pipeline in which third-party code cannot push, tag or publish, and the signed tarball on each GitHub release is checked against the integrity the npm registry records.

### Added

- A `locale` prop sets the interface language.
- `customization.ribbon.hiddenTabs` hides native ribbon tabs by id, such as `'insert'` or `'file'`. Hiding `'file'` removes the File tab, its backstage and every route into it at every width, while Save, its shortcuts, host `storage` and the `file-command` event keep working. Tell me leaves out commands that live only on a hidden tab.
- Drawings open with pages of up to about 10,000 shapes as Visio saves them: the default XML limits are 1,500,000 nodes per part and 2,000,000 per package, and the ten-second load deadline is unchanged. A page that is still too large says how many shapes it has.
- A drawing too large to edit here (more than 100,000 ShapeSheet cells, or an edit that would not fit the memory budget) opens read-only and says why. `parseVsdx` reports `edit-size-limit` first among its diagnostics, and a refused edit is shown in the viewer's language.
- A signed drawing shows its verified state in the status bar, in a notice under the ribbon and on File > Info: Signed, Signed (signer not verified), Invalid signature or Signature not verified. The notice says that saving changes removes the signatures.
- When sharing stops because an update was refused, `events['document-error']` reports why in the display language.
- F6 and Shift+F6 move between the ribbon, the Shapes window, the drawing, the task pane, the page tabs and the status bar, as in Word, Excel and PowerPoint.

### Changed

- `zoom-change` fires once per zoom gesture instead of once per wheel step: 300 ms after the last input, when the gesture ends, or before a pointer press or scroll. Until then the `zoom` property still reads the zoom the gesture started from.
- When the viewer is removed from the page and not attached again before the current task ends, it adds one constructed stylesheet of its own to `document.adoptedStyleSheets` (on first use) and changes its rule, which matches nothing. A host that replaces `document.adoptedStyleSheets` may drop it; the next removal adds it again.
- Charts keep the point counts and points their file declares through viewing, editing and saving. Very large charts show a bounded view, the load reports `CHART_POINTS_LIMITED`, and they stay fully editable.
- The package is built and tested against current releases of its dependencies, including `emf-converter` 5. The supported browsers are the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).

### Fixed

- `capabilities.export: false` now removes Export CSV from Review > Shape Reports, and `capabilities.copy: false` removes its Copy.
- Dialogs keep Tab and Shift+Tab inside, close on Escape wherever focus has ended up, and give focus back to the control that opened them.
- A shape keeps keyboard focus when the drawing redraws, so arrow nudges can be repeated and Delete still works. Focus moves to the canvas when the shape is gone.
- Pressing a selection or rotation handle while a zoom is still previewing starts the drag, and the rotation handle moves with the previewed drawing.
- `--vv-accent` colours the accent text again, ahead of the shared ink. The new `--vv-accent-text` sets that text alone.
- Chart Edit Data writes only the workbook cells an edit changed, so a text cell in a value range is no longer overwritten with 0. Removing a bubble chart row removes its bubble sizes, adding a scatter or bubble row leaves the new x value blank, and a value edit on a series shorter than the longest one is no longer refused.
- In Chrome and Edge, a closed viewer is freed again when its drawing used a clip, mask, filter or fill reference, such as a glow. The browser's style cache had kept it, with its whole drawing, alive.
- The Share dialog shows the same caption as Word and Excel for an invitation without its session key.

### Performance

- Zooming with Ctrl+wheel, a pinch, the zoom slider or the + and - keys scales the drawing smoothly and renders the page once when you stop. The end zoom and scroll position are unchanged.
- The page chrome and canvas sizes are written only when they change.

### Accessibility

- Every change is checked in a real browser before release: Tab goes round the page without getting stuck, F6 reaches every region, dialogs keep and return focus, and axe-core runs in light mode, dark mode and on a phone.
- The current page tab reads at 4.5:1 in every accent, and keeps a visible label in forced-colours mode. Pressed buttons and selected tabs keep a system highlight there too.
- Group shapes show a focus ring, and connectors describe the shapes they join.
- Gallery buttons such as Design > Backgrounds and Insert > Callout show their caption in the display language, and the More button of an inline gallery is named in that language.
- A chart is named by its alt text, falling back to its title, and screen readers can read its series and values.
- Field and shapes search boundaries reach 3:1, and the Tell me box, the search fields and the File tab show a focus ring of at least 3:1 in light and dark themes.

### Released together

- [`visio-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/core/CHANGELOG.md): the document API this package re-exports.
- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): parsing, editing, the memory budget and collaboration behind the viewer.
- [`ooxml-ui` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/ui/CHANGELOG.md): the `<visio-viewer>` element, its ribbon and dialogs, and the theme tokens.
- [`visio-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/vue/CHANGELOG.md): the same viewer for Vue.
- [`visio-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/angular/CHANGELOG.md): the same viewer for Angular.
- [`visio-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/svelte/CHANGELOG.md): the same viewer for Svelte.
- [`visio-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/solid/CHANGELOG.md): the same viewer for Solid.
- [`visio-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/packages/vanilla/CHANGELOG.md): the same viewer for pages without a framework.
- [`visio-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/mcp/CHANGELOG.md): the same document operations as MCP tools.

## [1.4.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.4.1) - 2026-11-16

### Dependencies

- **deps:** Take the open dependency updates in one change ([8eab074](https://github.com/ChristopherVR/ooxml/commit/8eab0749af462b944cda96a92e723132e9031f7a))

## [1.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.4.0) - 2026-11-13

### Features

- **visio:** Expose an awaited inline text commit ([8e3a23b](https://github.com/ChristopherVR/ooxml/commit/8e3a23b58cb054f8ab51f0e74e0f681bea47d659))

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.3.2](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.3.2) - 2026-11-09

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 25 updates ([4046ca3](https://github.com/ChristopherVR/ooxml/commit/4046ca399665d485da35696889c0212cc6526f63))

## [1.3.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.3.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.3.0) - 2026-11-05

### Features

- **visio:** Open drawings from cloud file providers ([f7563e7](https://github.com/ChristopherVR/ooxml/commit/f7563e70c8fbdb2475d7b211b94bf278b590de8c))
- **ui:** Expose openCloudFile on the word, excel and visio bindings ([b4837c7](https://github.com/ChristopherVR/ooxml/commit/b4837c76398c125755458861ce62679d7c179003))

### Bug Fixes

- **deps:** Update emf-converter to 4.25.3 ([e51a619](https://github.com/ChristopherVR/ooxml/commit/e51a6192b3a531022b3677753e79b6a015233b9d))

## [1.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.2.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.1.3](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.1.3) - 2026-11-03

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([8e60b14](https://github.com/ChristopherVR/ooxml/commit/8e60b14e1d8ee3a1cb507670caa15a611632ef26))
- **deps:** Merge production dependency updates ([#236](https://github.com/ChristopherVR/ooxml/issues/236)) ([7eb137b](https://github.com/ChristopherVR/ooxml/commit/7eb137b03906ded043ce0e4adf83e01fcdaa44a9))

## [1.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.1.2) - 2026-10-26

### Dependencies

- **deps:** Bump production minor and patch dependencies ([#213](https://github.com/ChristopherVR/ooxml/issues/213)) ([3cfdd01](https://github.com/ChristopherVR/ooxml/commit/3cfdd016aa61cacb0540491c5ab13974e32f8e1e))

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.1.1) - 2026-10-23

### Other

- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))
- Visio, Teams and package check docs (docs6-5) ([e6b2151](https://github.com/ChristopherVR/ooxml/commit/e6b2151824f447a53875592ee9f0a5f89f89cc69))

### Refactor

- **visio:** Build the Visio bindings on ooxml-ui/bindings ([70e9c6c](https://github.com/ChristopherVR/ooxml/commit/70e9c6c77505ef97f9100942aa149cfe26f1cc63))

### Documentation

- **visio:** Describe the shared binding kit, new props and chart editing ([ef66cf6](https://github.com/ChristopherVR/ooxml/commit/ef66cf68620585b6ed3ea070d5690cae2a258121))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.1.0) - 2026-10-22

### Features

- **visio:** Add file-command and host-owned storage to the viewer ([0c8e284](https://github.com/ChristopherVR/ooxml/commit/0c8e284015fe1cc1b895c840900482ba12f8cda4))
- **visio:** Take the shared customization contract in <visio-viewer> ([3a6bb30](https://github.com/ChristopherVR/ooxml/commit/3a6bb305b8ac04f0942b40e8cdfbd887093c6112))
- **visio:** Accept readOnly as the inverse of canEdit ([fbc325b](https://github.com/ChristopherVR/ooxml/commit/fbc325be8c9c12a7c2e4a3affd7af2428ba4ed49))
- **embedding:** Land the host customization, read-only and storage contracts ([822a6a0](https://github.com/ChristopherVR/ooxml/commit/822a6a073840b9e9d1671476542322b28332dc42))

### Bug Fixes

- **visio:** Reset readOnly when the host stops passing it ([c3db6eb](https://github.com/ChristopherVR/ooxml/commit/c3db6eb272b248329afbca32f70a72965f3f9d04))

### Other

- Integrate the customization contract into host storage ([629554b](https://github.com/ChristopherVR/ooxml/commit/629554b3030728bfe16e11ac8ae5b2530bb28a17))
- Bring origin/main into the embedding work ([a3711a1](https://github.com/ChristopherVR/ooxml/commit/a3711a14049a2cd11b994e76eafec039db2dc2de))

## [1.0.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.0.1) - 2026-10-22

### Documentation

- **visio:** Sync the viewer docs with the monorepo and current code ([2436dcc](https://github.com/ChristopherVR/ooxml/commit/2436dcce519728d28a65a124947b6dbbb5854cd6))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@1.0.0) - 2026-10-22

### Features

- **visio:** Take the shared collaboration contract in <visio-viewer> ([63ad160](https://github.com/ChristopherVR/ooxml/commit/63ad160c8ff5dc043695756f1bb9572fb19c2a4a))
- **visio:** Complete the building blocks in every framework ([0a69177](https://github.com/ChristopherVR/ooxml/commit/0a6917719b5040587c0a37dfe3b37a875b7d2a05))

### Bug Fixes

- **visio:** Ship the toolbar controls with the svelte package ([709fafa](https://github.com/ChristopherVR/ooxml/commit/709fafa524945526e613472143abf41054131332))

## [0.8.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.8.0) - 2026-10-21

### Features

- **visio:** Expose canvas building blocks and host permissions ([e464a1c](https://github.com/ChristopherVR/ooxml/commit/e464a1c0399434430925e0e59b430b5788f5d8d9))
- **visio:** Add native framework building block integrations ([689e934](https://github.com/ChristopherVR/ooxml/commit/689e9340824a41e32a5024938e16b483e34c4452))
- **visio:** Theme native building blocks with shared office tokens ([f86b882](https://github.com/ChristopherVR/ooxml/commit/f86b882e2081540997f14af754805c573b372dc9))

### Bug Fixes

- **ui:** Preserve live formula drafts and runtime compatibility ([3a7d4f1](https://github.com/ChristopherVR/ooxml/commit/3a7d4f184658c4d69b04bbd36bab7afa3896dd9e))

### Chores

- Integrate latest main before full verification ([3bec814](https://github.com/ChristopherVR/ooxml/commit/3bec814857ae675f31eba14fcec59a3db4e1cb77))

## [0.7.4](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.7.4) - 2026-10-20

### Dependencies

- **deps:** Update emf-converter to 4.21.0 ([#106](https://github.com/ChristopherVR/ooxml/issues/106)) ([641ae31](https://github.com/ChristopherVR/ooxml/commit/641ae316b7de07ed6f3c5158e8b154c991e1b719))

## [0.7.3](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.7.3) - 2026-10-20

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([02d7713](https://github.com/ChristopherVR/ooxml/commit/02d7713ceb301921fd6c7706138d64f4c10a9d38))
- **deps:** Merge dependency update ([#94](https://github.com/ChristopherVR/ooxml/issues/94)) ([e381aba](https://github.com/ChristopherVR/ooxml/commit/e381abaa11794233cdea85a9dc6d098d4cb338c4))

## [0.7.2](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.7.2) - 2026-10-13

### Dependencies

- **deps:** Update production dependencies ([bbd7b53](https://github.com/ChristopherVR/ooxml/commit/bbd7b530c7a15ae48c0c0a2ff4ea7dcdfe2aa9d3))

## [0.7.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.7.1) - 2026-10-11

### Styling

- Fix the oxfmt and oxlint warnings in the viewers ([db0314f](https://github.com/ChristopherVR/ooxml/commit/db0314f319e3d0508960872260792215477d9ca5))

## [0.7.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.7.0) - 2026-10-10

### Features

- **visio:** Make ribbon add-in tabs a prop and event of every binding ([6e6a157](https://github.com/ChristopherVR/ooxml/commit/6e6a157bbced977145d428e2ca01168e0a413deb))

## [0.6.2](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.6.2) - 2026-10-09

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

## [0.6.1](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.6.1) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.6.0](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.6.0) - 2026-10-08

### Features

- **visio:** Enable source-backed formatting and shape ordering ([c39f7be](https://github.com/ChristopherVR/ooxml/commit/c39f7bee397d27afa31155a96c01997122402773))
- **visio:** Enable multi-selection formatting and arrangement ([fb70098](https://github.com/ChristopherVR/ooxml/commit/fb70098a2b100d6ec37cd091acd618c91477cfb2))
- **visio:** Enable rich formatting and source duplication ([d50e388](https://github.com/ChristopherVR/ooxml/commit/d50e388f9190809d6a860894663cdbf774ca4dca))
- **visio:** Enable source-backed shape clipboard ([2c13c1f](https://github.com/ChristopherVR/ooxml/commit/2c13c1f39998cf37e83d3949a6685f84f4b16e4d))
- **visio:** Enable new drawings and pointer selection gestures ([75ac409](https://github.com/ChristopherVR/ooxml/commit/75ac409e5e67c23203c765d3aa472b6817853090))

## [0.5.9](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.5.9) - 2026-10-08

## [0.5.8](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.5.8) - 2026-10-07

## [0.5.7](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.5.7) - 2026-10-07

### Refactor

- **visio:** Use extensionless relative source imports ([2495aa9](https://github.com/ChristopherVR/ooxml/commit/2495aa92d52a0135106695a2a39a72c9f98cf177))

## [0.5.6](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.5.6) - 2026-10-07

## [0.5.5](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.5.5) - 2026-10-06

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Dependencies

- **deps:** Align emf-converter on 4.8.21 ([3c2859d](https://github.com/ChristopherVR/ooxml/commit/3c2859d0a43e66666015debaeaa985888727db95))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.5.4](https://github.com/ChristopherVR/ooxml/releases/tag/visio-react-viewer@0.5.4) - 2026-10-05

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

## 0.5.3

### Changes

- build(deps): adopt ooxml-core 0.15.0 and ooxml-ui 0.21.0 (818f4a4)

## 0.5.2

### Changes

- build(viewer): adopt ooxml-ui 0.18.2 (44f44e2)
- build(viewer): adopt ooxml-ui 0.17 (2e19d89)

## 0.5.1

### Changes

- build(viewer): adopt ooxml-ui 0.15 shared select and ribbon commands (e1fe065)

## 0.5.0

### Changes

- build(viewer): adopt token based ooxml-ui 0.13 (590a5c8)
- refactor(viewer): use shared ooxml-ui ribbon, backstage, find and rulers (cd1088b)
- feat(viewer): add visio account, options dialog and live sharing (fe7fed1)

## 0.4.0

### Changes

- feat(viewer): add pan & zoom window and a shapes strip to reopen shapes (01c5a82)
- feat(viewer): add visio keytips for keyboard ribbon access (8a9d6d1)
- feat(viewer): add visio tell me, help tab and presentation mode (d048a45)
- feat(viewer): add visio file backstage, context menus and page bar (8c54263)
- feat(viewer): add visio's shapes window with a basic shapes stencil (98cb55e)
- feat(viewer): match the visio ribbon layout with honest disabled commands (bb74402)

## 0.3.0

### Changes

- fix(packages): pin released core and shared ui dependencies (4d035f6)
- feat(visio): preview legacy VSD through the shared binary codec (6772ff3)

## 0.2.0

### Changes

- feat(bindings): expose native reactive viewer state in every binding (ccd9ba8)
- feat(viewer): add visio ribbon toolbar on shared office-ui controls (afc7047)
- fix(deps): use released core for proven master shape moves (28989be)

## 0.1.1

### Changes

- docs(packages): align npm readmes with powerpoint structure (5c6f120)
- fix(ui): match suite Pages and responsive editor composition (960d3bc)

## 0.1.0

### Changes

- feat(release): publish functional visio framework packages (bc55722)
- feat(viewer): preserve drafts across geometry edits and history (c5b585d)
- fix(viewer): align Office shell and Pages navigation (23d6912)
- feat(viewer): expose safe geometry commands and pin published core (8ddaf87)
- fix(viewer): follow the shared office theme inside frames (caad674)
- feat(ui): align Visio workspace and docs with PPTX design (011c728)
- feat(visio): restore concave arrows and rounded connectors (7662a75)

## 0.0.2

### Changes

- Initial release.
