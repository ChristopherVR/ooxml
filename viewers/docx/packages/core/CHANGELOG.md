# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

For code that reads, edits and writes Word documents without an editor, `docx-core` is now safer to point at untrusted files and faster on large ones. Every package is opened through the guarded zip reader under limits you can set with the new `limits` option, a signed document reports whether its signature verifies in `model.signatures`, and saving an edit to a long document costs time in proportion to the edit. Three changes can need code: Node.js 22.3, Word style catalogs as values, and the signature message leaving `warnings`. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) has the before and after for each change.

### Breaking changes

- Node.js 22.3 or newer is required; the package declares `"engines": { "node": ">=22.3" }`. See [Run Node.js 22.3 or newer](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#run-nodejs-223-or-newer).
- Word style catalogs (`paragraphStyles`, `characterStyles`, `tableStyles`) are values. Inheritance is resolved once per `styles` record, so change a style by replacing the record and the style, as the editor's numbering edits do, not by changing a style object in place. See [Word style catalogs are values](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#word-style-catalogs-are-values).
- Loading a signed document no longer adds an English "signature not verified" message to `warnings`. Read `model.signatures.state` and word it with `signatureStateText` from `ooxml-core/i18n`. See [Signatures are reported as a verified summary](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#signatures-are-reported-as-a-verified-summary).

### Security

- Every package is opened through the guarded zip reader, so a small file can no longer inflate without limit. Signature listing ignores a package whose signature part is larger than 16 MiB.
- Scans of raw XML that took time quadratic in the number of empty or unclosed elements, such as the declared-font scan that runs before a document is parsed, are now linear. A few kilobytes of such elements could block a load for a minute.
- A document counts as signed only when its signature verifies, never because it merely contains a signature part. Checking is bounded: an RSA key under 1024 or over 16384 bits is reported as `unsupported`, each signature tries at most four certificates and reads at most 10,000 references and 100 MB of part data, and a signature part over 16 MB is skipped.
- The package is published by a release pipeline in which third-party code cannot push, tag or publish, with npm provenance, and the signed tarballs on each GitHub release are checked against the integrity the npm registry records.

### Added

- `loadDocx` takes a `limits` option (`DocxLoadLimits`). It sets the compressed package size, the number of parts, the inflated size of one part and of all parts, and the size of `word/document.xml`, which was fixed at 25 MiB. The defaults are unchanged and exported as `DEFAULT_DOCX_LOAD_LIMITS`, a value that is not a positive finite number throws a `RangeError`, and saving reopens the package under the limits it was loaded with.
- A signed document carries its verified state as `model.signatures`, a `SignatureSummary` (`{ state, signatures }`) whose state is `verified`, `verified-untrusted`, `invalid` or `unsupported`. A browser has no trust store, so an intact signature from any signer is `verified-untrusted` there.
- `InlineImage.textBoxContent` holds a text box's `w:txbxContent` as real paragraphs and tables, read by the same parser as the document body, with the anchor and insets of its `wps:bodyPr`, for every shape that has a `wps:txbx` whatever its preset. `drawingShapeSvg` draws a drawing's shape without its own text when those paragraphs are drawn over it.
- `ShapeSpec.fontColor` is the colour of a shape's `wps:style` font reference, and `simpleTextBoxDrawing` gives the frame and paragraphs an edited simple text box is saved with.

### Changed

- A text box written with `txBox="1"` (Word's Insert > Text Box) is read with its frame like any shape, and an unchanged box is saved byte for byte.
- The text inside a drawing shape counts as editable only when it is the plain centred text the editor writes, so changing the shape's size, fill or outline no longer rewrites richer text as plain centred lines.

### Fixed

- An empty element no longer swallows the element after it in the raw XML scans that run before a document is parsed. Because of it, a font listed before an embedded font in the font table was treated as embedded.
- Saving a shape whose text takes its colour from the shape's style writes no colour on its runs, so Word shows the same colour after a resize or refill, and text whose runs share one colour keeps it when it is rewritten.
- Saving a chart edit writes only the embedded workbook cells that edit changed, so a text cell in a value range is no longer overwritten with 0. Removing a row of a bubble chart removes its bubble sizes with it, and adding a row to a scatter or bubble chart leaves the new x value blank instead of writing a category label into the x values.

### Performance

- Saving no longer moves blocks that did not change. An edit to a 20,000-block document saves in about 2.5 s instead of about 17 s, inserting, deleting or moving a paragraph costs the same as editing one, and the saved document is the same as before, byte for byte.
- Loading walks `document.xml` once for every unsupported feature it reports instead of once per feature, and copies the loaded model once instead of twice; saves no longer copy it again.
- Style inheritance is resolved once per catalog, so a long `basedOn` chain no longer makes loading quadratic or slows each paragraph and run as it is laid out.

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): the `docx` area this package re-exports, and the shared zip, XML and signature fixes.
- [`docx-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/react/CHANGELOG.md): the Word editor for React, built on this model.
- [`docx-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/vue/CHANGELOG.md): the Word editor for Vue.
- [`docx-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/angular/CHANGELOG.md): the Word editor for Angular.
- [`docx-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/svelte/CHANGELOG.md): the Word editor for Svelte.
- [`docx-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/solid/CHANGELOG.md): the Word editor for Solid.
- [`docx-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/vanilla/CHANGELOG.md): the Word editor without a framework.
- [`docx-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/mcp/CHANGELOG.md): Word tools for MCP clients over the same model.

## [1.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@1.1.2) - 2026-11-13

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@1.1.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@1.1.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@1.0.0) - 2026-10-29

### Features

- **docx:** Commit docx-core and docx-viewer-mcp to a stable 1.0 api ([7aef089](https://github.com/ChristopherVR/ooxml/commit/7aef089212f2c496360a0ec1cca4b1831d8d0771))

## [0.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.3.0) - 2026-10-22

### Features

- **docx:** Add navigation, status and full toolbar blocks to bindings ([d3ff60f](https://github.com/ChristopherVR/ooxml/commit/d3ff60f9c90291d11d7ff59b8ffcea96dc9bc464))

## [0.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.2.0) - 2026-10-21

### Features

- **docx:** Add native framework building blocks ([77db7c6](https://github.com/ChristopherVR/ooxml/commit/77db7c6a9965a6851fd0ee42488228a7931a632e))

## [0.1.13](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.1.13) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.1.12](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.1.12) - 2026-10-08

## [0.1.11](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.1.11) - 2026-10-07

## [0.1.10](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.1.10) - 2026-10-07

## [0.1.9](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.1.9) - 2026-10-07

## [0.1.8](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.1.8) - 2026-10-06

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.1.7](https://github.com/ChristopherVR/ooxml/releases/tag/docx-core@0.1.7) - 2026-10-05

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

## [0.1.6](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.6) - 2026-10-05

### Refactor

- **web-component:** Take the editor from ooxml-ui/docx and update the toolchain (by @ChristopherVR) ([3cd4d87](https://github.com/ChristopherVR/docx-viewer/commit/3cd4d876838421743fca9d7fa0666d775d597207))

### Build & CI

- **deps:** Adopt ooxml-ui 0.26.0 (by @ChristopherVR) ([5a3bde4](https://github.com/ChristopherVR/docx-viewer/commit/5a3bde4e165a9dd71c95416328df79675e220122))
- **deps:** Adopt ooxml-core 0.18.0 and ooxml-ui 0.27.1 (by @ChristopherVR) ([a020eb9](https://github.com/ChristopherVR/docx-viewer/commit/a020eb9bd5a7b8fdeebb816d3ec3c78c9c56a3da))
- **deps:** Adopt ooxml-core 0.19.0 and ooxml-ui 0.27.1 (by @ChristopherVR) ([3049491](https://github.com/ChristopherVR/docx-viewer/commit/3049491c83cfd55036cd794c31f4a1ba57219515))

## [0.1.5](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.5) - 2026-10-05

### Build & CI

- **deps:** Adopt ooxml-core 0.15.0 and ooxml-ui 0.21.0 (by @ChristopherVR) ([91eac57](https://github.com/ChristopherVR/docx-viewer/commit/91eac573a278cc9a2f23e4a85871ddc6eb33d9c6))
- **deps:** Update all dependencies to latest (by @ChristopherVR) ([b0fd02d](https://github.com/ChristopherVR/docx-viewer/commit/b0fd02decdbcdca871d4ac8b21967c6a24d4cbbf))

## [0.1.4](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.4) - 2026-10-03

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([e6c664c](https://github.com/ChristopherVR/docx-viewer/commit/e6c664ccc06d5e207e4fd7c5a455433ab8868976))

## [0.1.3](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.3) - 2026-10-03

### Dependencies

- **deps:** Consume shared alignment and layout source ranges (by @ChristopherVR) ([81b8c51](https://github.com/ChristopherVR/docx-viewer/commit/81b8c5148f139f9e8eadde14491eb167b042072b))

## [0.1.2](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.2) - 2026-10-03

### Dependencies

- **deps:** Consume shared equation and cell margin support (by @ChristopherVR) ([9e943fa](https://github.com/ChristopherVR/docx-viewer/commit/9e943fa61bb2f380c2a2cd1153bb1f0f6e276819))

## [0.1.1](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-core@0.1.1) - 2026-10-03

### Features

- Add shared Word editor and multi-framework demos (by @ChristopherVR) ([77ea606](https://github.com/ChristopherVR/docx-viewer/commit/77ea606625752f976ed426967614b3bf31e13403))
- **editor:** Add Word run formatting and safe table editing (by @ChristopherVR) ([c06bbaf](https://github.com/ChristopherVR/docx-viewer/commit/c06bbaf25d6b5baa3e57dbcd2ee1bc88e0be1b16))
- **editor:** Add line spacing and hard line breaks (by @ChristopherVR) ([b414ec4](https://github.com/ChristopherVR/docx-viewer/commit/b414ec46857e1309ff31aa7603bf033874301c84))
- **editor:** Add collaboration, multilingual editing and search (by @ChristopherVR) ([da68728](https://github.com/ChristopherVR/docx-viewer/commit/da687283507df6e7aa6886fe866405c454aaa16c))
- Add list numbering, sections, breaks, headers/footers and notes (by @ChristopherVR) ([6a5462f](https://github.com/ChristopherVR/docx-viewer/commit/6a5462f973d64dfa3268baba4fb0c73f3fa76d5a))
- **layout:** Add pagination engine and Print Layout view (by @ChristopherVR) ([c30c6e1](https://github.com/ChristopherVR/docx-viewer/commit/c30c6e1d456e2357239468e1946d4577e544fb9c))
- **review:** Add tracked changes and comments (by @ChristopherVR) ([4d7a1e5](https://github.com/ChristopherVR/docx-viewer/commit/4d7a1e52e0417ccff8319355dd244da30af2b40e))
- **core:** Add theme, character styles and table fidelity (by @ChristopherVR) ([8bf918c](https://github.com/ChristopherVR/docx-viewer/commit/8bf918cbd8f2115f668d0446acbe146253249eeb))
- Add inline pictures, hyperlinks and bookmarks (by @ChristopherVR) ([7e0ff2b](https://github.com/ChristopherVR/docx-viewer/commit/7e0ff2bf6c6428057d4ef8eaa0bf2ccb6c4e68f5))
- **layout:** Show headers, footers and page numbers in Print Layout (by @ChristopherVR) ([825fbf5](https://github.com/ChristopherVR/docx-viewer/commit/825fbf5de71598bbdc4a2c60888805d9f8256ba7))
- Edit headers and footers in place (by @ChristopherVR) ([9699a8b](https://github.com/ChristopherVR/docx-viewer/commit/9699a8b50c4d962dc30af079d087c323f55c0862))
- Edit footnotes and endnotes in place (by @ChristopherVR) ([4976310](https://github.com/ChristopherVR/docx-viewer/commit/497631029e768d1f58761f26a6fdfc8b0e139884))
- Insert footnotes and endnotes (by @ChristopherVR) ([7ae74f4](https://github.com/ChristopherVR/docx-viewer/commit/7ae74f45712b6ab8ff962cd9bb088a38acac0ad5))
- Edit sections, page setup and section breaks (by @ChristopherVR) ([235d2f1](https://github.com/ChristopherVR/docx-viewer/commit/235d2f1740e19b283bcd807173c84cac881f0c60))
- **sections:** Edit different first page and page numbering (by @ChristopherVR) ([a6d12da](https://github.com/ChristopherVR/docx-viewer/commit/a6d12da51c6880eb8d5ee6dfb8df9da3493bf390))
- Pictures and links in headers, footers and notes (by @ChristopherVR) ([e4522f5](https://github.com/ChristopherVR/docx-viewer/commit/e4522f5e1de53afa78a71e418bec90edfc93aa94))
- **pictures:** Approximate floating picture wrapping (by @ChristopherVR) ([f904915](https://github.com/ChristopherVR/docx-viewer/commit/f9049151102532f15621e1435d429627fe9d6d64))
- **sections:** Vertical alignment, odd/even breaks and odd/even headers (by @ChristopherVR) ([309b113](https://github.com/ChristopherVR/docx-viewer/commit/309b11359f14e3aa497b0a1ed8b600db35357a0f))
- **fields:** Update DATE and TIME fields in Print Layout (by @ChristopherVR) ([6890466](https://github.com/ChristopherVR/docx-viewer/commit/6890466525f91ae404b3ae998d69b82ff8228eb6))
- **pictures:** SVG pictures with PNG fallback (by @ChristopherVR) ([e583d8a](https://github.com/ChristopherVR/docx-viewer/commit/e583d8a4835571fb5d6d2ed8fb2a6525f61ecfda))
- **fields:** Editable complex and simple fields (by @ChristopherVR) ([43a7f7d](https://github.com/ChristopherVR/docx-viewer/commit/43a7f7d4da59210c7ecd80551402eb5f8a9b506a))
- **toc:** Insert and update tables of contents; paragraph tab stops (by @ChristopherVR) ([ce93a8d](https://github.com/ChristopherVR/docx-viewer/commit/ce93a8d529794878394d9b85aef10a16d5a75769))
- **print-layout:** Paginate inline pictures and position floating pictures (by @ChristopherVR) ([d68e089](https://github.com/ChristopherVR/docx-viewer/commit/d68e08945823b0c65badb45c855c2c7eccec3c25))
- **comments:** Comment ranges spanning paragraphs; decimal comment ids (by @ChristopherVR) ([166e482](https://github.com/ChristopherVR/docx-viewer/commit/166e4825710cee33048ad581492c8632c0c8f661))
- **revisions:** Link tracked moves and keep their range markers (by @ChristopherVR) ([2ab59b2](https://github.com/ChristopherVR/docx-viewer/commit/2ab59b2f42f6eb877c34d3625b5448b4ae67548a))
- **print-layout:** Indents, tab stops with leaders, and list labels (by @ChristopherVR) ([a4f55a6](https://github.com/ChristopherVR/docx-viewer/commit/a4f55a6c29354ff452debae4af05fc37c8ad4140))
- **toc:** Hyperlinked entries with _Toc bookmarks and PAGEREF fields (by @ChristopherVR) ([5254ad9](https://github.com/ChristopherVR/docx-viewer/commit/5254ad96f008b9f7e78972052c680e0f45c94e3d))
- **track-changes:** Record moves; keep pasted formatting valid (by @ChristopherVR) ([5196ee7](https://github.com/ChristopherVR/docx-viewer/commit/5196ee7b9107d741251b11a1893115cba5939139))
- Word defaults for new documents; Print Layout tables (by @ChristopherVR) ([614d0fd](https://github.com/ChristopherVR/docx-viewer/commit/614d0fde2b600e601317c9aa1c226bda1e86b5f1))
- **print-layout:** Footnotes, endnotes and superscripts (by @ChristopherVR) ([2c48fc5](https://github.com/ChristopherVR/docx-viewer/commit/2c48fc5c2e8411f63c18cd06287426ab02bab363))
- **tables:** Row heights, header rows, keep-together and default cell margins (by @ChristopherVR) ([788b74a](https://github.com/ChristopherVR/docx-viewer/commit/788b74a6358f2081844f056fa91813569078c06f))
- **pagination:** Keep-with-next, keep-lines, widow control, contextual spacing (by @ChristopherVR) ([55ece2b](https://github.com/ChristopherVR/docx-viewer/commit/55ece2b0b6823983ea00fc44767e249c680e648f))
- Paragraph borders and shading (by @ChristopherVR) ([d9af53f](https://github.com/ChristopherVR/docx-viewer/commit/d9af53f85c6e9592f5df4cf285a6e1f65e845942))
- Added schemas (by @ChristopherVR) ([13ee39c](https://github.com/ChristopherVR/docx-viewer/commit/13ee39cc509504dbcc53f2b37b1e710432e4f526))
- **core:** Schema-typed parsing with generated ECMA-376 simple types (by @claude) ([5aeb7b5](https://github.com/ChristopherVR/docx-viewer/commit/5aeb7b552eab29c6f3cdd8f88756d3a6a752065e))
- **core:** Validate the model before save; write table cell props, exact jc and unique docPr ids (by @claude) ([fa24aac](https://github.com/ChristopherVR/docx-viewer/commit/fa24aacd9341604fa357c78a0a0d7ef34e4c134c))
- **core:** Create header/footer parts on save and write paragraph shading and borders (by @ChristopherVR) ([ec49241](https://github.com/ChristopherVR/docx-viewer/commit/ec492419dc52e83bbfda5866dac8f964111a153b))
- **core:** Model page background and hyphenation settings (by @ChristopherVR) ([65e12fc](https://github.com/ChristopherVR/docx-viewer/commit/65e12fcba549494bff6c8d21782e96b9ccf26fb2))
- Insert > Drop Cap with framePr round trip and layout folding (by @ChristopherVR) ([cabd832](https://github.com/ChristopherVR/docx-viewer/commit/cabd8327d111f8d86a0a0294fb476b5e8510352d))
- Read text boxes and show their text read-only (by @ChristopherVR) ([f9e3c56](https://github.com/ChristopherVR/docx-viewer/commit/f9e3c565134048bb5b9082a3b8ef286534cb8d74))
- Home > Multilevel List with legal and outline styles (by @ChristopherVR) ([c113f64](https://github.com/ChristopherVR/docx-viewer/commit/c113f64cfdad3df8ae1e6bd1e7259680d3bd63f7))
- Layout > Line Numbers with Print Layout rendering (by @ChristopherVR) ([881da88](https://github.com/ChristopherVR/docx-viewer/commit/881da88edbfb0d5592ef4d31819164c54a2cff22))
- Marker font controls for Define New Multilevel List (by @ChristopherVR) ([ef462cf](https://github.com/ChristopherVR/docx-viewer/commit/ef462cfe43fc24f62a1f65a0313967127a8ac23f))
- Heading-linked styles in Define New Multilevel List (by @ChristopherVR) ([c89e0e0](https://github.com/ChristopherVR/docx-viewer/commit/c89e0e001222c53b2134b8be7f8b3d2e07ca5c36))
- **web-component:** Heading-linked outlines in the Multilevel List gallery (by @ChristopherVR) ([40258a3](https://github.com/ChristopherVR/docx-viewer/commit/40258a3ace72df7c0e040fd7fd55b34222a8a2be))
- Editable document properties (title, subject, author, tags, comments) (by @ChristopherVR) ([3f53c23](https://github.com/ChristopherVR/docx-viewer/commit/3f53c233dd20c7c961a4137779ee3098e203aa84))
- Explicit tab stop per level in Define New Multilevel List (by @ChristopherVR) ([e7667ad](https://github.com/ChristopherVR/docx-viewer/commit/e7667ad28995dafc80e4a18f855f4058b663d81b))
- References > Table of Figures with update, and fix escaped field tests (by @ChristopherVR) ([4b239e2](https://github.com/ChristopherVR/docx-viewer/commit/4b239e2e2bf0759ba84cb27956969f7a80d53d9f))
- Table and cell border editor (Home > Borders and Borders and Shading Apply to) (by @ChristopherVR) ([2597e65](https://github.com/ChristopherVR/docx-viewer/commit/2597e65dc1bb083222922d89c68492c345264955))
- Layout > Page Borders (model, pgBorders write, Print Layout drawing, dialog) (by @ChristopherVR) ([99c4159](https://github.com/ChristopherVR/docx-viewer/commit/99c41590fac6aa108a518f7682d56428b8dd6824))
- Insert > Text Box (inline, editable), plus text-box line display and image run merge fixes (by @ChristopherVR) ([4d76849](https://github.com/ChristopherVR/docx-viewer/commit/4d768492d4ec50661356d63ff1431d0397cf77dc))
- Layout > Watermark (text watermark VML in headers, Print Layout layer, dialog) (by @ChristopherVR) ([21ec767](https://github.com/ChristopherVR/docx-viewer/commit/21ec7671afec67ba76f8da9b56d5059594591269))

### Bug Fixes

- **tables:** Keep cell formatting on edit, render table styles, write valid new tables (by @ChristopherVR) ([bf63864](https://github.com/ChristopherVR/docx-viewer/commit/bf63864b2f04487832e2ab366a9fd9385cba81dd))
- **revisions:** Save editor revision ids as decimals (by @ChristopherVR) ([bebb185](https://github.com/ChristopherVR/docx-viewer/commit/bebb1859f05b11a7528e9065cd061645a7dc058e))
- **formatting:** Word toggle-property semantics and explicit offs (by @ChristopherVR) ([84ad25d](https://github.com/ChristopherVR/docx-viewer/commit/84ad25d34f2e86460451bbaf05c0e9c2d9b1dcae))
- **core:** Write schema-ordered numbering, sectPr and drawing extents; make schema test real (by @claude) ([2b95fb7](https://github.com/ChristopherVR/docx-viewer/commit/2b95fb7ec32a2eae925d29f06d1302dfd059e785))
- **packaging:** Viewer Svelte component imports bindings helpers; typed svelte.d.ts (by @claude) ([ac57529](https://github.com/ChristopherVR/docx-viewer/commit/ac575296c0d13efb5ac4625c34ffd136b1e6c603))

### Other

- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([b4d3e44](https://github.com/ChristopherVR/docx-viewer/commit/b4d3e44463195e949316d3ac6df7bf9137485b4a))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([3cddda1](https://github.com/ChristopherVR/docx-viewer/commit/3cddda15ca30ac85612049ae80d63d947a230470))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([98b3036](https://github.com/ChristopherVR/docx-viewer/commit/98b3036664d3dedccfeb36dcbd9c99e7826795ec))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([52332bf](https://github.com/ChristopherVR/docx-viewer/commit/52332bf134d1c30a6ba63f7c61abbd79dad4b053))

### Refactor

- **core:** Remove writer helpers superseded by write-inline (by @ChristopherVR) ([f400024](https://github.com/ChristopherVR/docx-viewer/commit/f40002444f3c1dc14f1272e967c03c06865dabaa))
- Split ribbon, schema, model, adapter and collaboration modules (by @ChristopherVR) ([a3a1aee](https://github.com/ChristopherVR/docx-viewer/commit/a3a1aee3bf584c448f5de260b24a584dce346a60))
- Split range markers and body editor plugins into modules (by @ChristopherVR) ([7490a17](https://github.com/ChristopherVR/docx-viewer/commit/7490a17239c905b80c290b69402836fb704f50e6))
- **core:** Split oversized model and write modules (by @claude) ([b821cb6](https://github.com/ChristopherVR/docx-viewer/commit/b821cb64068d2b1887576770ea63de1fdcfbbc74))
- **core,legacy:** Make source safe under noUncheckedIndexedAccess (by @claude) ([78b75e7](https://github.com/ChristopherVR/docx-viewer/commit/78b75e79a22e6d5e3aa0f3d9ebe068350143d74b))
- **core:** Brand model measurements as Twips, SignedTwips and EighthPoints (by @claude) ([b4ee400](https://github.com/ChristopherVR/docx-viewer/commit/b4ee40048cb4d8a5f31ee26c5c23f7ff6cdbee89))
- **core:** Consume @christophervr/ooxml-units from the shared ooxml-core repo (by @ChristopherVR) ([3be6e4c](https://github.com/ChristopherVR/docx-viewer/commit/3be6e4c3a3e249f917fbf452c1ef2b2618f57216))
- **core:** Build the Word XML helpers on @christophervr/ooxml-xml (by @ChristopherVR) ([8795a34](https://github.com/ChristopherVR/docx-viewer/commit/8795a344f0656bcd810403927bf4a2c0518fac52))
- **core:** Consume relationship and content-type handling from ooxml-opc (by @ChristopherVR) ([5fbdea1](https://github.com/ChristopherVR/docx-viewer/commit/5fbdea1af1766ffba4b3580f50b735956591dc1f))
- Consume the single @christophervr/ooxml-core package (by @ChristopherVR) ([9fdcc2f](https://github.com/ChristopherVR/docx-viewer/commit/9fdcc2fc2a50dc59bf585e3b2138cc83e6047a26))
- **core:** Reduce docx-core to a thin re-export of @christophervr/ooxml-core/docx (by @ChristopherVR) ([76fb536](https://github.com/ChristopherVR/docx-viewer/commit/76fb536f2f9f6f9931226a10ac8b9aeef46c316f))
- **packages:** Publish one self-contained package per framework (by @ChristopherVR) ([f009c62](https://github.com/ChristopherVR/docx-viewer/commit/f009c6266315805753bcc55ce0688272396a8d9b))
- Adopt ooxml-core layout, load and collab areas (by @ChristopherVR) ([aa42730](https://github.com/ChristopherVR/docx-viewer/commit/aa427301f8af62707e50db6c915afa0b91b579a6))

### Documentation

- **schemas:** Add Ecma and W3C licence texts for the vendored XSDs (by @claude) ([210a1a5](https://github.com/ChristopherVR/docx-viewer/commit/210a1a576906172045c844c279c6e6e1ceda23c3))

### Testing

- Make browser runner resilient to missing pinned Chromium; document schema licence provenance (by @claude) ([6299813](https://github.com/ChristopherVR/docx-viewer/commit/6299813b28fa36f2de2078f1ab1c260194434187))
- **core:** Make tests safe under noUncheckedIndexedAccess (by @claude) ([ac9f888](https://github.com/ChristopherVR/docx-viewer/commit/ac9f8887cb56aad68b36d09fb05868d611cdf0c3))

### Build & CI

- Consume the published @christophervr/ooxml-core ^0.1.0 (by @ChristopherVR) ([38df96a](https://github.com/ChristopherVR/docx-viewer/commit/38df96a11de41072ece01648b8da118fe5729b7a))
- Publish under unscoped docx-* package names (by @ChristopherVR) ([44b369f](https://github.com/ChristopherVR/docx-viewer/commit/44b369ff2df6e34f210f7752ebd510df9464b405))

### Styling

- Apply oxfmt across the workspace (by @ChristopherVR) ([e9b3725](https://github.com/ChristopherVR/docx-viewer/commit/e9b3725582062b8140f76f7019912961b81b7ce8))

### Dependencies

- **deps:** Consume the published office core table fixes (by @ChristopherVR) ([a7c1616](https://github.com/ChristopherVR/docx-viewer/commit/a7c16165fb91b88b4688a69fca13c4eecef3659e))

### Chores

- **types:** Enable exactOptionalPropertyTypes (by @claude) ([6f46221](https://github.com/ChristopherVR/docx-viewer/commit/6f4622160f37ae94a3bb5fa50710456b798dcaab))
- **types:** Enable noUncheckedIndexedAccess; keep test-support out of release builds (by @claude) ([203b289](https://github.com/ChristopherVR/docx-viewer/commit/203b289f0e8ca5d877cc6ff990ebc35755a853a3))
- Renamed ooxml packages [ci-skip] (by @ChristopherVR) ([3e75728](https://github.com/ChristopherVR/docx-viewer/commit/3e757282589b0327b90657d0d4c4e9350ce5025a))

