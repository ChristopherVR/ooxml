# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

For React apps, `WordEditor` keeps its props, callbacks and `ref` handle: what changes is the editor it mounts. Documents load behind safety limits you can now tune, a signed document shows whether its signature actually verifies, saving an edit to a long document takes time in proportion to the edit, and keyboard, screen reader and forced-colours use are checked in a real browser before every release. Read the breaking changes for dark mode, which now follows the theme rather than CSS `color-scheme`, and for collaboration, where an editor removed from the page leaves its room. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) has the before and after for each change.

### Breaking changes

- Node.js 22.3 or newer is required for builds, server rendering and tests; the package declares `"engines": { "node": ">=22.3" }`. See [Run Node.js 22.3 or newer](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#run-nodejs-223-or-newer).
- The editor no longer inherits the page's CSS `color-scheme`. It reads the new `--office-color-scheme` token, which is `light` unless the theme's dark mode applies (`data-office-theme="dark"`, or the system's dark preference where no `data-office-theme="light"` overrides it). A container that sets only `color-scheme: dark` now gets a light editor: also set `--office-color-scheme: dark` on it, or use `data-office-theme="dark"`, which sets both. See [the migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#the-editors-take-dark-mode-from-the-theme-not-from-color-scheme).
- An editor that shares a room (`collaboration`) leaves it when it is removed from the page and not attached again before the current task ends, and joins the same room again when it is attached later. Moving it between parents in one task keeps it in the room. To hide an editor and keep it syncing, hide it with CSS instead of detaching it. See [the migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#an-editor-removed-from-the-page-leaves-its-room).
- A collaboration update that holds more than one struct per 4 bytes, or more than one deleted range per 2 bytes, is refused before it is decoded, and the refusal stops sharing. The budget grows with the update, so the first sync of a real document stays under half of it. Code that passed `DEFAULT_UPDATE_LIMITS` from `ooxml-core/collab` uses `defaultUpdateLimits(bytes)`, `MIN_BYTES_PER_STRUCT` and `MIN_BYTES_PER_DELETE_RANGE`. See [the migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#collaboration-updates-are-bounded-and-a-refusal-stops-the-sync).
- The `--dve-input` colour now defaults to `--office-control-border` (`#8a8886`) instead of `--office-border`, and it colours the shared form controls inside the editor. A host that wants lighter field boundaries sets `--office-control-border` or `--dve-input`, and keeps it at 3:1 against the background. See [the migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#form-controls-use---office-control-border).
- Loading a signed document no longer adds an English signature message to `model.warnings`. Read `model.signatures.state` and word it with `signatureStateText` from `ooxml-core/i18n`. See [the migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#signatures-are-reported-as-a-verified-summary).
- Word style catalogs (`paragraphStyles`, `characterStyles`, `tableStyles`) are values: inheritance is resolved once per `styles` record, so change a style by replacing the record and the style, not by changing a style object in place. See [the migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#word-style-catalogs-are-values).
- If you copy compiled files by hand instead of installing with a package manager, deploy the `browser-chunks/` and `node-chunks/` folders next to the entry files and keep their relative paths: `emf-converter` 5 loads image decoders and colour data from them on demand. Package managers and bundlers need no change. See [the migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#deploy-the-new-decoder-folders-if-you-copy-files-by-hand).

### Security

- A document is shown as signed only when its signature verifies. The status bar, a notice under the ribbon and File > Info show Signed, Signed (signer not verified), Invalid signature or Signature not verified, in every editor language, and the notice says that saving changes removes the signatures. Before, any signature part was enough to show a document as signed. Checking a signature is bounded, so a crafted one cannot stall the page.
- Every document is opened through the guarded zip reader, embedded chart workbooks included, so a small file can no longer inflate without limit.
- Scans of raw XML that took time quadratic in the number of empty or unclosed elements, such as the declared-font scan that runs before every load, are now linear. A few kilobytes of such elements could block a load for a minute.
- Chart point counts and point indices no longer decide how much memory is allocated, so a chart part of a few hundred bytes can no longer exhaust memory.
- SVG pictures are shown from `data:` URLs instead of object URLs of your application's origin, so opening one in its own tab can no longer run document-controlled script in that origin. Raster pictures keep object URLs, which now always carry an image type, so a type a collaboration peer claims cannot make a picture open as markup. Placeholder sizes can no longer add style text.
- Packages are published by a release pipeline in which third-party code cannot push, tag or publish, with npm provenance, and the signed tarballs on each GitHub release are checked against the integrity the npm registry records.

### Added

- `customization.ribbon.hiddenTabs` hides native ribbon tabs by id (`WordRibbonTabId`, such as `'insert'`, `'references'` or `'mailings'`). `'file'` removes the File tab and its backstage at every width, with its key tip and every other route into the backstage, while Save, its shortcuts, `storage` and the `file-command` event keep working. Tell me leaves out commands that live only on a hidden tab, and leaving the option out keeps the full ribbon.
- `<docx-editor>` has `getTextSelection()` and `focusText(range)`, so a host that reloads a document can put the user's caret back. Reach it through the `ref` handle's `element`.
- `themeColors` takes `primaryInk`, the primary colour as text (`--dve-primary-ink`). It defaults to `primary`, and the dark presets lift it.
- `loadDocument` takes a `limits` option (`DocxLoadLimits`) for the compressed package size, the number of parts, the inflated size of one part and of all parts, and the size of `word/document.xml`, which was fixed at 25 MiB. The defaults are unchanged and exported as `DEFAULT_DOCX_LOAD_LIMITS`.
- When sharing stops because an update was refused, the editor raises `document-error` with the reason in the display language instead of a raw English message.

### Changed

- Text boxes written with `txBox="1"` (Word's Insert > Text Box) are read with their frame and drawn like any shape. Selecting one shows the Shape Format tab, Add Text and Insert > Text Box edit it in the Text Box dialog as before, and an unchanged box is saved byte for byte.
- The text inside a drawing shape is editable only when it is the plain centred text the editor writes. Text with bold, sizes, borders or other alignment is read-only, so resizing or refilling the shape no longer rewrites it as plain centred lines.
- Draft view's header, footer and note previews leave floating drawings out, as Word's Draft view does, so a tall letterhead text box no longer stretches the header preview. Inline drawings still show, and Print Layout is unchanged.
- Charts keep the point counts and points their file declares, however large, through viewing, editing and saving, and stay fully editable. The views the editor draws and edits share a budget of blank positions (at most 1,048,576 per view and 4,194,304 per document); a view past it shows fewer trailing blanks, then only its first points, and the load reports `CHART_POINTS_LIMITED`.
- The status bar's zoom button is named by the percentage it shows first, for example "100%, Zoom to fit". Tests that find it by name need the new name.
- A removed editor adds one constructed stylesheet of its own to `document.adoptedStyleSheets` (on first use) and changes its rule, which matches nothing, so Chrome and Edge release the editor. A host that replaces `document.adoptedStyleSheets` may drop it; the next removal adds it again.
- The supported browsers are the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).
### Fixed

- A text box drawn as a shape (a `rect` holding a `wps:txbx`, as letterheads often have) shows its paragraphs with their runs, bold and italic, sizes, colours, alignment and paragraph borders, inset and clipped to the shape over its fill and outline. Such shapes used to show as an empty dashed box or the word "Shape". `txBox="1"` text boxes are drawn in Print Layout, floating and inline, in the body and in headers and footers, instead of being left out.
- Lines and connectors such as `straightConnector1` in headers and footers are drawn in Print Layout with their outline colour, width and dash, instead of an empty placeholder box.
- Shape text whose runs set no colour takes the colour of the shape's `wps:style` font reference, resolved through the theme, as Word draws it, and saving keeps it so Word shows the same colour after a resize or refill. A shape with neither a style nor its own fill or outline is drawn without them instead of in the default blue.
- Pictures in Draft view's previews and inside text boxes are drawn from the same media cache as the body, so re-rendering a preview creates no new picture URLs, all of them are released with the document, and pictures that arrive late in a collaboration session repaint there too.
- Saving a chart edit writes only the embedded workbook cells that edit changed. Before, a text cell in a value range was overwritten with 0 by any chart save, even a rename. Removing a row of a bubble chart removes its bubble sizes with it, adding a row to a scatter or bubble chart leaves the new x value blank, and Edit Data no longer refuses a value edit on a series shorter than the longest one in a chart with a whole-column range.
- Dialogs keep focus: Tab and Shift+Tab can no longer leave them, Escape closes them wherever focus has ended up, and focus returns to the control that opened them, including inside shadow roots.
- A removed editor no longer stays in memory. Print, font and ribbon listeners are released on removal and restored when the same element is attached again, open menus and popups close, and an editor that was typed in or shared a room is freed too. In Chrome and Edge, an editor whose document drew a clipped, masked or filtered picture or shape (a tiled or cropped picture, for example) is no longer kept alive by the browser's style cache. Removing one editor no longer closes a popover, overflow menu or word count panel open in another editor on the same page.
- Hyperlinks on the page use the document link colour (`--office-canvas-link`, `#0563c1`) instead of the interface accent, which was hard to read in dark mode.
- A font listed before an embedded font in the font table is no longer treated as embedded, so its web font is requested when font downloads are on.
- The Share dialog shows the same caption as Excel and Visio for an invitation without its session key.
### Performance

- Saving no longer moves blocks that did not change. An edit to a 20,000-block document saves in about 2.5 s instead of about 17 s, inserting, deleting or moving a paragraph costs the same as editing one, and the saved document is the same as before, byte for byte.
- Loading walks `document.xml` once for every unsupported feature it reports instead of once per feature, and copies the loaded model once instead of twice; saves no longer copy it again. Style inheritance is resolved once per catalog, so a long `basedOn` chain no longer makes loading quadratic or slows layout.
- Opening and closing documents in a long session no longer grows memory by the size of the editor each time. In the Office suite, a Word open-and-close cycle went from about 3 MB and 12,000 DOM nodes to about 0.2 MB and none.
- The SmartArt Design colour and style galleries load the built-in SmartArt definitions (about 390 KB compressed) when a SmartArt is selected, not when the editor opens.

### Accessibility

- Every change is checked in a real browser before it can be released: Tab goes round the page without getting stuck, F6 moves between regions, dialogs keep focus and give it back, and saving leaves focus where it was. axe-core runs on the editor in light mode, dark mode and on a phone and fails on any new critical or serious problem, and the checks also cover what a screen reader is told about a document, accent contrast, reflow at 320 by 256 pixels (400% zoom), and that a saved file keeps its alt text, headings, table header rows and language.
- Form control boundaries reach 3:1, including the ribbon's combo boxes and drop-downs, the mode selector and the title bar's search and Tell me boxes. Tell me, the search fields and the File tab show a focus ring of at least 3:1 in light and dark themes.
- The Font and Size boxes expose their expanded state and the list they control.
- Ribbon commands whose caption is shorter than their command name have their own accessible name in every locale: Insert > Pictures is "Insert pictures" ("Bilder einfügen", "Insérer des images", "Insertar imágenes", "插入图片"). Dialogs, menus and tooltips keep their wording.
- In forced-colours mode, the active content control, contextual ribbon tabs, tall ribbon buttons, building-block toggles and every selected, current, checked or pressed state keep a system highlight instead of disappearing.
- A chart is named by its alt text, falling back to its title, and screen readers can read its series and values.

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): the Word model, loading, layout and saving behind the editor, and the shared security fixes.
- [`ooxml-ui` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/ui/CHANGELOG.md): the `<docx-editor>` element, its ribbon and dialogs, and the shared Office controls it mounts.
- [`docx-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/core/CHANGELOG.md): the Word model API this package re-exports.
- [`docx-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/vue/CHANGELOG.md): the same editor for Vue.
- [`docx-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/angular/CHANGELOG.md): the same editor for Angular.
- [`docx-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/svelte/CHANGELOG.md): the same editor for Svelte.
- [`docx-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/solid/CHANGELOG.md): the same editor for Solid.
- [`docx-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/packages/vanilla/CHANGELOG.md): the same editor without a framework.
- [`docx-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/mcp/CHANGELOG.md): Word tools for MCP clients over the same core.

## [1.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@1.4.0) - 2026-11-13

### Features

- **docx:** Expose read aloud through the building blocks ([01dc2ed](https://github.com/ChristopherVR/ooxml/commit/01dc2edb2618a74273841147670dd2fd4976b128))
- **docx:** Export the read aloud result types ([d33509f](https://github.com/ChristopherVR/ooxml/commit/d33509f5fbcd68d94c6e600c6b94cd0ad9b1fcba))
- **docx:** Make read aloud stop on edit configurable ([cf73e1b](https://github.com/ChristopherVR/ooxml/commit/cf73e1bf4165f7244b336939f09c711a0f2c8cf4))

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.3.1](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@1.3.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@1.3.0) - 2026-11-05

### Features

- **docx:** Forward fileProviders through every binding ([cb585af](https://github.com/ChristopherVR/ooxml/commit/cb585af52665b172455ba592e74dae71a67c8c91))
- **ui:** Expose openCloudFile on the word, excel and visio bindings ([b4837c7](https://github.com/ChristopherVR/ooxml/commit/b4837c76398c125755458861ce62679d7c179003))

## [1.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@1.2.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@1.1.2) - 2026-10-29

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@1.1.1) - 2026-10-23

### Other

- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))
- Word SmartArt editing (wave6-13) ([751ba15](https://github.com/ChristopherVR/ooxml/commit/751ba157ca578e8ff6b57371a243180101ed562e))
- Word and Excel docs for wave 6 (docs6-4) ([da0f3e8](https://github.com/ChristopherVR/ooxml/commit/da0f3e83c59b8ff55104bd852cf46fa052c587b5))

### Refactor

- **docx:** Build the Word bindings on ooxml-ui/bindings ([90e8131](https://github.com/ChristopherVR/ooxml/commit/90e81310800081ad484e93a3bab25806a07bc3e4))

### Documentation

- **docs:** Describe word smartart editing and its limits ([7a1ef43](https://github.com/ChristopherVR/ooxml/commit/7a1ef43c5bc0ba910b0f2468516e3d3578158605))
- Describe wave 6 chart and smartart editing in word and excel ([d212908](https://github.com/ChristopherVR/ooxml/commit/d212908f59be7f5da51ce5347b23a665e4cedb08))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@1.1.0) - 2026-10-22

### Features

- **docx:** Route save and open through host-owned storage ([ec27cd2](https://github.com/ChristopherVR/ooxml/commit/ec27cd28c4f594446e08b60ab05db2b64b72c683))
- **docx:** Forward file, review and ribbon events through every binding ([feac0f1](https://github.com/ChristopherVR/ooxml/commit/feac0f1de2287f9549209c1b35c4b76bbe944257))
- **docx:** Take the shared customization contract in <docx-editor> ([65f05d3](https://github.com/ChristopherVR/ooxml/commit/65f05d3d70cc28e6633f300fb4758aca7ba34184))
- **embedding:** Land the host customization, read-only and storage contracts ([822a6a0](https://github.com/ChristopherVR/ooxml/commit/822a6a073840b9e9d1671476542322b28332dc42))

### Bug Fixes

- **docx:** Forward readOnly only when the prop changes ([fb2b324](https://github.com/ChristopherVR/ooxml/commit/fb2b3245789ffb8b05f93cd046d0beb397916b44))
- **docx:** Report read-only with the lock and host switch ([735fb0d](https://github.com/ChristopherVR/ooxml/commit/735fb0d78ff18dc71eeeb83db2a76f26e423ba30))
- **docx:** Honour find: false in Ctrl+F and the shortcut help ([f38d2fa](https://github.com/ChristopherVR/ooxml/commit/f38d2fae33cad0444e201ab99635d56a985c9b26))

### Other

- Integrate the customization contract into host storage ([629554b](https://github.com/ChristopherVR/ooxml/commit/629554b3030728bfe16e11ac8ae5b2530bb28a17))

### Testing

- **docx:** Assert the binding customization state inline ([e83b995](https://github.com/ChristopherVR/ooxml/commit/e83b995ee36701cd6c6842e04a2dba3589897c43))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@1.0.0) - 2026-10-22

### Features

- **docx:** Take the shared collaboration contract in <docx-editor> ([58e5caa](https://github.com/ChristopherVR/ooxml/commit/58e5caa1103f2129b32976cf1def24ebd031cad9))
- **docx:** Add navigation, status and full toolbar blocks to bindings ([d3ff60f](https://github.com/ChristopherVR/ooxml/commit/d3ff60f9c90291d11d7ff59b8ffcea96dc9bc464))
- **docx:** Make the building blocks read-only unless canEdit is true ([f4f4227](https://github.com/ChristopherVR/ooxml/commit/f4f4227e43b8b6f9463f0b81fc23a925cb62dc01))

### Bug Fixes

- **docx:** Name the toolbar openers apart from the host options ([d5bc691](https://github.com/ChristopherVR/ooxml/commit/d5bc691e2aaa946147e55edf590a61f64c7a671a))

### Documentation

- **docx:** Document the building blocks and list them in the sidebar ([4b98e67](https://github.com/ChristopherVR/ooxml/commit/4b98e671c708d3e5ea286bd4797e29d88bfe1303))
- Align the building-block pages with the code ([37ab0f6](https://github.com/ChristopherVR/ooxml/commit/37ab0f6b89ff2a9a9ec14ace28bf320af01ff03a))

## [0.6.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.6.0) - 2026-10-21

### Features

- **docx:** Expose composable editor building blocks ([79bf480](https://github.com/ChristopherVR/ooxml/commit/79bf4801d8efb256a3e3e6fa81cca2613e6fd0fd))
- **docx:** Add native framework building blocks ([77db7c6](https://github.com/ChristopherVR/ooxml/commit/77db7c6a9965a6851fd0ee42488228a7931a632e))
- **docx:** Inherit shared office themes across editor blocks ([3adb658](https://github.com/ChristopherVR/ooxml/commit/3adb658f212e93a49beec631a618c8802b0c183a))

### Bug Fixes

- **deps:** Align private bindings with their react test host ([f4152ef](https://github.com/ChristopherVR/ooxml/commit/f4152ef04ccdc3e5c9d21ecf59f485b56de2ec52))
- **ui:** Harden composition lifecycle and consumer checks ([6aa52be](https://github.com/ChristopherVR/ooxml/commit/6aa52be88175ca4f46f6f3b1c223ef06a8f1f915))
- **ui:** Preserve live formula drafts and runtime compatibility ([3a7d4f1](https://github.com/ChristopherVR/ooxml/commit/3a7d4f184658c4d69b04bbd36bab7afa3896dd9e))

### Styling

- Fix the oxfmt and oxlint warnings in the viewers ([db0314f](https://github.com/ChristopherVR/ooxml/commit/db0314f319e3d0508960872260792215477d9ca5))

### Dependencies

- **deps:** Declare native adapter development dependencies ([b5f9957](https://github.com/ChristopherVR/ooxml/commit/b5f9957a65d934ebb4d40ed02bb872dcd24575fb))

## [0.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.5.0) - 2026-10-10

### Features

- **docx:** Make ribbon add-in tabs a prop and event of every binding ([819c1a3](https://github.com/ChristopherVR/ooxml/commit/819c1a38c2beb9605676d29994b71a7a6f7e4a15))

## [0.4.10](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.4.10) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.4.9](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.4.9) - 2026-10-08

## [0.4.8](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.4.8) - 2026-10-07

## [0.4.7](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.4.7) - 2026-10-07

## [0.4.6](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.4.6) - 2026-10-07

## [0.4.5](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.4.5) - 2026-10-06

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.4.4](https://github.com/ChristopherVR/ooxml/releases/tag/docx-react-viewer@0.4.4) - 2026-10-05

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

## [0.4.3](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-react-viewer@0.4.3) - 2026-10-05

### Refactor

- **web-component:** Draw the ribbon tab row, collapse and overflow with shared elements (by @ChristopherVR) ([cc57cd9](https://github.com/ChristopherVR/docx-viewer/commit/cc57cd97e602d6f7614d0b1308db895f390ce4c7))
- **web-component:** Use the shared KeyTip assignment and badges (by @ChristopherVR) ([c540a20](https://github.com/ChristopherVR/docx-viewer/commit/c540a20fba0a7a68996eaacd8ee9013bfe1e0456))
- **web-component:** Import the document statistics from ooxml-core (by @ChristopherVR) ([172e87d](https://github.com/ChristopherVR/docx-viewer/commit/172e87df2ac0bb633e5c1c77f674c9da82f3d455))
- **web-component:** Import column, line spacing and section helpers from ooxml-core (by @ChristopherVR) ([77f820f](https://github.com/ChristopherVR/docx-viewer/commit/77f820f3e802ed4f3711c5b9a956265969eb4f5d))
- **web-component:** Take the editor from ooxml-ui/docx and update the toolchain (by @ChristopherVR) ([3cd4d87](https://github.com/ChristopherVR/docx-viewer/commit/3cd4d876838421743fca9d7fa0666d775d597207))

### Build & CI

- **deps:** Adopt ooxml-ui 0.26.0 (by @ChristopherVR) ([5a3bde4](https://github.com/ChristopherVR/docx-viewer/commit/5a3bde4e165a9dd71c95416328df79675e220122))
- **deps:** Adopt ooxml-ui 0.27.0 (by @ChristopherVR) ([7db5d6e](https://github.com/ChristopherVR/docx-viewer/commit/7db5d6ea9e836d05a4a089055855f92c7a48c5f9))
- **deps:** Adopt ooxml-core 0.18.0 and ooxml-ui 0.27.1 (by @ChristopherVR) ([a020eb9](https://github.com/ChristopherVR/docx-viewer/commit/a020eb9bd5a7b8fdeebb816d3ec3c78c9c56a3da))
- **deps:** Adopt ooxml-core 0.19.0 and ooxml-ui 0.27.1 (by @ChristopherVR) ([3049491](https://github.com/ChristopherVR/docx-viewer/commit/3049491c83cfd55036cd794c31f4a1ba57219515))
- **deps:** Adopt ooxml-ui 0.27.2 (by @github-actions[bot]) ([14b579f](https://github.com/ChristopherVR/docx-viewer/commit/14b579fd9737e7d0f43a9e61beb43150de7d2186))

## [0.4.2](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-react-viewer@0.4.2) - 2026-10-05

### Refactor

- **web-component:** Use the shared office-ui-context-menu (by @ChristopherVR) ([4a5a4a3](https://github.com/ChristopherVR/docx-viewer/commit/4a5a4a37e50f51a5cd3b9e595323d78421a274b8))
- **web-component:** Draw the status bar with shared ooxml-ui elements (by @ChristopherVR) ([70871a8](https://github.com/ChristopherVR/docx-viewer/commit/70871a84308d14df45b0787cc88a6a35ab2158a4))
- **web-component:** Build the format dialogs on the shared office-ui-dialog (by @ChristopherVR) ([a273ad8](https://github.com/ChristopherVR/docx-viewer/commit/a273ad8d97a800b73594bd204b8545dc5b069a1b))
- **web-component:** Build the File view on the shared office-ui-backstage (by @ChristopherVR) ([93ee934](https://github.com/ChristopherVR/docx-viewer/commit/93ee9341f451d3d593b241b39acc38c982a325f7))
- **web-component:** Draw the title bar and shortcut help with shared elements (by @ChristopherVR) ([a45dd66](https://github.com/ChristopherVR/docx-viewer/commit/a45dd66883f3fd7cbe1d89fad1eb8326f601d786))

### Build & CI

- **deps:** Adopt ooxml-core 0.15.0 and ooxml-ui 0.21.0 (by @ChristopherVR) ([91eac57](https://github.com/ChristopherVR/docx-viewer/commit/91eac573a278cc9a2f23e4a85871ddc6eb33d9c6))
- **deps:** Update all dependencies to latest (by @ChristopherVR) ([b0fd02d](https://github.com/ChristopherVR/docx-viewer/commit/b0fd02decdbcdca871d4ac8b21967c6a24d4cbbf))

## [0.4.1](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-react-viewer@0.4.1) - 2026-10-03

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([e6c664c](https://github.com/ChristopherVR/docx-viewer/commit/e6c664ccc06d5e207e4fd7c5a455433ab8868976))

## [0.4.0](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-react-viewer@0.4.0) - 2026-10-03

### Features

- **web-component:** Edit selected cell vertical alignment (by @ChristopherVR) ([f744de1](https://github.com/ChristopherVR/docx-viewer/commit/f744de14876ad306fd8791f8ff2e38b987741b0b))

### Bug Fixes

- **web-component:** Preserve formatting in clipboard commands (by @ChristopherVR) ([453a451](https://github.com/ChristopherVR/docx-viewer/commit/453a451c1ce24044a2896feb58704f7eb318fb5f))
- **web-component:** Map print clicks to visible glyph boundaries (by @ChristopherVR) ([fc4e047](https://github.com/ChristopherVR/docx-viewer/commit/fc4e04778bdb4f38d6b8ec54e693c5f4f0b08220))

### Dependencies

- **deps:** Consume shared alignment and layout source ranges (by @ChristopherVR) ([81b8c51](https://github.com/ChristopherVR/docx-viewer/commit/81b8c5148f139f9e8eadde14491eb167b042072b))

## [0.3.0](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-react-viewer@0.3.0) - 2026-10-03

### Features

- **web-component:** Edit margins for selected table cells (by @ChristopherVR) ([52605c4](https://github.com/ChristopherVR/docx-viewer/commit/52605c40112d7c0f3a6c6c585c97dbabdafc7aed))
- **web-component:** Display and preserve imported equations (by @ChristopherVR) ([3286dc2](https://github.com/ChristopherVR/docx-viewer/commit/3286dc25f8258e17988a7a718b79c2c9466507f3))

### Bug Fixes

- **web-component:** Wrap scaled words across formatting runs (by @ChristopherVR) ([7821875](https://github.com/ChristopherVR/docx-viewer/commit/7821875536909eda84493337f23c8e0d861cc2ba))

### Dependencies

- **deps:** Consume shared equation and cell margin support (by @ChristopherVR) ([9e943fa](https://github.com/ChristopherVR/docx-viewer/commit/9e943fa61bb2f380c2a2cd1153bb1f0f6e276819))

## [0.2.0](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-react-viewer@0.2.0) - 2026-10-03

### Features

- **web-component:** Improve file home and live editor options (by @ChristopherVR) ([7db1f08](https://github.com/ChristopherVR/docx-viewer/commit/7db1f08933994c46e98dab9b2464d2249391ac7b))
- **web-component:** Edit table shading with undo and preservation (by @ChristopherVR) ([1525735](https://github.com/ChristopherVR/docx-viewer/commit/15257358ed6653bb87047d501eec6558fb9b6b8e))

### Dependencies

- **deps:** Consume the published office core table fixes (by @ChristopherVR) ([a7c1616](https://github.com/ChristopherVR/docx-viewer/commit/a7c16165fb91b88b4688a69fca13c4eecef3659e))
- **deps:** Align shared controls with the current office core (by @ChristopherVR) ([5f379cd](https://github.com/ChristopherVR/docx-viewer/commit/5f379cd71b4ad47fa8bb217ce6067af2c82da89a))

## [0.1.0](https://github.com/ChristopherVR/docx-viewer/releases/tag/docx-react-viewer@0.1.0) - 2026-10-02

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
- **editor:** Add Word-style window chrome and full-screen demo (by @ChristopherVR) ([7a5744c](https://github.com/ChristopherVR/docx-viewer/commit/7a5744c156bec56803dcd85adb53825641c9bbaf))
- **editor:** Add picture, hyperlink, character style and hidden-text commands (by @ChristopherVR) ([535758a](https://github.com/ChristopherVR/docx-viewer/commit/535758a75d88e2c886117f17d73a55e01fd6fed6))
- **editor:** Resize pictures, edit alt text, and honor toggle cancellation (by @ChristopherVR) ([5a3aa2d](https://github.com/ChristopherVR/docx-viewer/commit/5a3aa2d8eaad0569ced99a89aaad656913556675))
- **layout:** Show headers, footers and page numbers in Print Layout (by @ChristopherVR) ([825fbf5](https://github.com/ChristopherVR/docx-viewer/commit/825fbf5de71598bbdc4a2c60888805d9f8256ba7))
- Edit headers and footers in place (by @ChristopherVR) ([9699a8b](https://github.com/ChristopherVR/docx-viewer/commit/9699a8b50c4d962dc30af079d087c323f55c0862))
- Edit footnotes and endnotes in place (by @ChristopherVR) ([4976310](https://github.com/ChristopherVR/docx-viewer/commit/497631029e768d1f58761f26a6fdfc8b0e139884))
- Insert footnotes and endnotes (by @ChristopherVR) ([7ae74f4](https://github.com/ChristopherVR/docx-viewer/commit/7ae74f45712b6ab8ff962cd9bb088a38acac0ad5))
- Edit sections, page setup and section breaks (by @ChristopherVR) ([235d2f1](https://github.com/ChristopherVR/docx-viewer/commit/235d2f1740e19b283bcd807173c84cac881f0c60))
- **tables:** Render table style text formatting and cell margins (by @ChristopherVR) ([3cdfe41](https://github.com/ChristopherVR/docx-viewer/commit/3cdfe414077d94554f104882fb0887b58d2a6c15))
- **sections:** Edit different first page and page numbering (by @ChristopherVR) ([a6d12da](https://github.com/ChristopherVR/docx-viewer/commit/a6d12da51c6880eb8d5ee6dfb8df9da3493bf390))
- Pictures and links in headers, footers and notes (by @ChristopherVR) ([e4522f5](https://github.com/ChristopherVR/docx-viewer/commit/e4522f5e1de53afa78a71e418bec90edfc93aa94))
- **pictures:** Approximate floating picture wrapping (by @ChristopherVR) ([f904915](https://github.com/ChristopherVR/docx-viewer/commit/f9049151102532f15621e1435d429627fe9d6d64))
- **editor:** Ribbon targets the header, footer or note being edited (by @ChristopherVR) ([cd9d309](https://github.com/ChristopherVR/docx-viewer/commit/cd9d30953f667e414e62f833a72f2d1927b51bad))
- **sections:** Vertical alignment, odd/even breaks and odd/even headers (by @ChristopherVR) ([309b113](https://github.com/ChristopherVR/docx-viewer/commit/309b11359f14e3aa497b0a1ed8b600db35357a0f))
- **fields:** Update DATE and TIME fields in Print Layout (by @ChristopherVR) ([6890466](https://github.com/ChristopherVR/docx-viewer/commit/6890466525f91ae404b3ae998d69b82ff8228eb6))
- **pictures:** SVG pictures with PNG fallback (by @ChristopherVR) ([e583d8a](https://github.com/ChristopherVR/docx-viewer/commit/e583d8a4835571fb5d6d2ed8fb2a6525f61ecfda))
- **fields:** Editable complex and simple fields (by @ChristopherVR) ([43a7f7d](https://github.com/ChristopherVR/docx-viewer/commit/43a7f7d4da59210c7ecd80551402eb5f8a9b506a))
- **toc:** Insert and update tables of contents; paragraph tab stops (by @ChristopherVR) ([ce93a8d](https://github.com/ChristopherVR/docx-viewer/commit/ce93a8d529794878394d9b85aef10a16d5a75769))
- **print-layout:** Paginate inline pictures and position floating pictures (by @ChristopherVR) ([d68e089](https://github.com/ChristopherVR/docx-viewer/commit/d68e08945823b0c65badb45c855c2c7eccec3c25))
- **print-layout:** Style-resolved run metrics and metric-compatible fonts (by @ChristopherVR) ([f3f37fb](https://github.com/ChristopherVR/docx-viewer/commit/f3f37fb314cd5885c446aa4fcbf0cc7626e89d18))
- **comments:** Comment ranges spanning paragraphs; decimal comment ids (by @ChristopherVR) ([166e482](https://github.com/ChristopherVR/docx-viewer/commit/166e4825710cee33048ad581492c8632c0c8f661))
- **revisions:** Link tracked moves and keep their range markers (by @ChristopherVR) ([2ab59b2](https://github.com/ChristopherVR/docx-viewer/commit/2ab59b2f42f6eb877c34d3625b5448b4ae67548a))
- **print-layout:** Indents, tab stops with leaders, and list labels (by @ChristopherVR) ([a4f55a6](https://github.com/ChristopherVR/docx-viewer/commit/a4f55a6c29354ff452debae4af05fc37c8ad4140))
- **print-layout:** Header and footer pictures (by @ChristopherVR) ([b993c0d](https://github.com/ChristopherVR/docx-viewer/commit/b993c0dd088a4d114f8649eba3e8ffc5cd1374b1))
- **toc:** Hyperlinked entries with _Toc bookmarks and PAGEREF fields (by @ChristopherVR) ([5254ad9](https://github.com/ChristopherVR/docx-viewer/commit/5254ad96f008b9f7e78972052c680e0f45c94e3d))
- **editor:** Tab stops and leaders on the editing surface (by @ChristopherVR) ([f8181a8](https://github.com/ChristopherVR/docx-viewer/commit/f8181a887aba2a8e78b67706f06966cb2d67bee2))
- **track-changes:** Record moves; keep pasted formatting valid (by @ChristopherVR) ([5196ee7](https://github.com/ChristopherVR/docx-viewer/commit/5196ee7b9107d741251b11a1893115cba5939139))
- Word defaults for new documents; Print Layout tables (by @ChristopherVR) ([614d0fd](https://github.com/ChristopherVR/docx-viewer/commit/614d0fde2b600e601317c9aa1c226bda1e86b5f1))
- **print-layout:** Footnotes, endnotes and superscripts (by @ChristopherVR) ([2c48fc5](https://github.com/ChristopherVR/docx-viewer/commit/2c48fc5c2e8411f63c18cd06287426ab02bab363))
- **print-layout:** Baseline-aligned text (by @ChristopherVR) ([4ad14db](https://github.com/ChristopherVR/docx-viewer/commit/4ad14db7534cbc6935460f35b90e3140a19a54f7))
- **tables:** Row heights, header rows, keep-together and default cell margins (by @ChristopherVR) ([788b74a](https://github.com/ChristopherVR/docx-viewer/commit/788b74a6358f2081844f056fa91813569078c06f))
- **pagination:** Keep-with-next, keep-lines, widow control, contextual spacing (by @ChristopherVR) ([55ece2b](https://github.com/ChristopherVR/docx-viewer/commit/55ece2b0b6823983ea00fc44767e249c680e648f))
- Paragraph borders and shading (by @ChristopherVR) ([d9af53f](https://github.com/ChristopherVR/docx-viewer/commit/d9af53f85c6e9592f5df4cf285a6e1f65e845942))
- Added schemas (by @ChristopherVR) ([13ee39c](https://github.com/ChristopherVR/docx-viewer/commit/13ee39cc509504dbcc53f2b37b1e710432e4f526))
- **web-component:** Design tokens, dark mode and focus ring (by @claude) ([65d4d97](https://github.com/ChristopherVR/docx-viewer/commit/65d4d97d9dac947cf9a1d84c80f53f5d7e993e47))
- **core:** Schema-typed parsing with generated ECMA-376 simple types (by @claude) ([5aeb7b5](https://github.com/ChristopherVR/docx-viewer/commit/5aeb7b552eab29c6f3cdd8f88756d3a6a752065e))
- **web-component:** Typed events, reflected attributes and shared binding keys (by @claude) ([8b3d7a2](https://github.com/ChristopherVR/docx-viewer/commit/8b3d7a2c72e23a458ff12ef42a3cb9ba8da7b05d))
- **core:** Validate the model before save; write table cell props, exact jc and unique docPr ids (by @claude) ([fa24aac](https://github.com/ChristopherVR/docx-viewer/commit/fa24aacd9341604fa357c78a0a0d7ef34e4c134c))
- **web-component:** Page thumbnail rail, save/dirty API and ribbon customisation (by @claude) ([0239e27](https://github.com/ChristopherVR/docx-viewer/commit/0239e275630bdf4129de81febf143fcaf037846a))
- **web-component:** Central keyboard shortcuts, shortcut help and context menu (by @claude) ([581369f](https://github.com/ChristopherVR/docx-viewer/commit/581369feb861e700f81eec62f9a43d44dfc7afa0))
- **web-component:** Stable kebab-case ribbon action ids for hiddenActions (by @claude) ([a6441cb](https://github.com/ChristopherVR/docx-viewer/commit/a6441cb8d97e250bb3dc7c38c20879881841ed11))
- **web-component:** Per-locale strings with de, es and zh-CN, no hard-coded UI English (by @claude) ([7c609c1](https://github.com/ChristopherVR/docx-viewer/commit/7c609c1b1e9dab547a0f19af761915ce33a0c7a1))
- **core:** Create header/footer parts on save and write paragraph shading and borders (by @ChristopherVR) ([ec49241](https://github.com/ChristopherVR/docx-viewer/commit/ec492419dc52e83bbfda5866dac8f964111a153b))
- **web-component:** Word-style ribbon across every tab, dialogs and real commands (by @ChristopherVR) ([94a776b](https://github.com/ChristopherVR/docx-viewer/commit/94a776bc2d2aa4d4c40b7fcb2688b5fdefd2dc27))
- **web-component:** Word-style File backstage with Info, Save As, Print, Export and Options (by @ChristopherVR) ([1f7c3c0](https://github.com/ChristopherVR/docx-viewer/commit/1f7c3c027c9efff0c4d6cb55164d936cd11332f1))
- **web-component:** Word-style Styles gallery and Editing group on Home (by @ChristopherVR) ([df9232e](https://github.com/ChristopherVR/docx-viewer/commit/df9232e012aea84701f63d12c8d9440c1dd6fc9c))
- **web-component:** Home > Sort for selected paragraphs (by @ChristopherVR) ([9498d92](https://github.com/ChristopherVR/docx-viewer/commit/9498d929bcaae303adf6e6b5eea97df2c584bff6))
- **web-component:** Insert tab gets a table size picker, Bookmark dialog, Cover Page and Horizontal Line (by @ChristopherVR) ([a5009d7](https://github.com/ChristopherVR/docx-viewer/commit/a5009d753d0a75c5978ec00dc96d5e25fa4e1387))
- **web-component:** Page Setup dialog on Layout and outstanding-work update (by @ChristopherVR) ([cf7b815](https://github.com/ChristopherVR/docx-viewer/commit/cf7b8153510d728b0c36f037f84d2d52009ab678))
- **web-component:** View > Ruler with margins, inch ticks and indent markers (by @ChristopherVR) ([ee075dc](https://github.com/ChristopherVR/docx-viewer/commit/ee075dc997a90e490e73ce97ff58d9414ba853ed))
- **web-component:** References tab gets TOC depth and removal, captions and cross-references (by @ChristopherVR) ([49dcb35](https://github.com/ChristopherVR/docx-viewer/commit/49dcb357f772beb0671118ed2dd892ee7cae1f52))
- **web-component:** Review comments group gains Delete, Previous and Next (by @ChristopherVR) ([9921a8e](https://github.com/ChristopherVR/docx-viewer/commit/9921a8e732cdf017ca9fd89774624990ad84bdd4))
- **core:** Model page background and hyphenation settings (by @ChristopherVR) ([65e12fc](https://github.com/ChristopherVR/docx-viewer/commit/65e12fcba549494bff6c8d21782e96b9ccf26fb2))
- **web-component:** Ribbon parity pass with galleries, colour pickers, KeyTips, overflow and Zoom dialog (by @ChristopherVR) ([6555c45](https://github.com/ChristopherVR/docx-viewer/commit/6555c450267346b295995eea353b611e7c57217f))
- **web-component:** Margins gallery gains Custom Margins (by @ChristopherVR) ([904fd50](https://github.com/ChristopherVR/docx-viewer/commit/904fd50793d90c3397c4f84963a8b86f0620193e))
- **web-component:** Drag Ruler indent markers (by @ChristopherVR) ([1d6faac](https://github.com/ChristopherVR/docx-viewer/commit/1d6faacd6225604a9b4ec224c04eeae2f650a292))
- Insert > Drop Cap with framePr round trip and layout folding (by @ChristopherVR) ([cabd832](https://github.com/ChristopherVR/docx-viewer/commit/cabd8327d111f8d86a0a0294fb476b5e8510352d))
- Read text boxes and show their text read-only (by @ChristopherVR) ([f9e3c56](https://github.com/ChristopherVR/docx-viewer/commit/f9e3c565134048bb5b9082a3b8ef286534cb8d74))
- Home > Multilevel List with legal and outline styles (by @ChristopherVR) ([c113f64](https://github.com/ChristopherVR/docx-viewer/commit/c113f64cfdad3df8ae1e6bd1e7259680d3bd63f7))
- **web-component:** Borders and Shading dialog (by @ChristopherVR) ([ea1d591](https://github.com/ChristopherVR/docx-viewer/commit/ea1d591038fa646df4987d2f81000859e88819a3))
- Layout > Line Numbers with Print Layout rendering (by @ChristopherVR) ([881da88](https://github.com/ChristopherVR/docx-viewer/commit/881da88edbfb0d5592ef4d31819164c54a2cff22))
- **web-component:** References > Update Fields for SEQ, REF and PAGEREF (by @ChristopherVR) ([ff8c082](https://github.com/ChristopherVR/docx-viewer/commit/ff8c082167f909749d28a371721ae89847fb221a))
- Marker font controls for Define New Multilevel List (by @ChristopherVR) ([ef462cf](https://github.com/ChristopherVR/docx-viewer/commit/ef462cfe43fc24f62a1f65a0313967127a8ac23f))
- Heading-linked styles in Define New Multilevel List (by @ChristopherVR) ([c89e0e0](https://github.com/ChristopherVR/docx-viewer/commit/c89e0e001222c53b2134b8be7f8b3d2e07ca5c36))
- **web-component:** Heading-linked outlines in the Multilevel List gallery (by @ChristopherVR) ([40258a3](https://github.com/ChristopherVR/docx-viewer/commit/40258a3ace72df7c0e040fd7fd55b34222a8a2be))
- Editable document properties (title, subject, author, tags, comments) (by @ChristopherVR) ([3f53c23](https://github.com/ChristopherVR/docx-viewer/commit/3f53c233dd20c7c961a4137779ee3098e203aa84))
- Explicit tab stop per level in Define New Multilevel List (by @ChristopherVR) ([e7667ad](https://github.com/ChristopherVR/docx-viewer/commit/e7667ad28995dafc80e4a18f855f4058b663d81b))
- References > Table of Figures with update, and fix escaped field tests (by @ChristopherVR) ([4b239e2](https://github.com/ChristopherVR/docx-viewer/commit/4b239e2e2bf0759ba84cb27956969f7a80d53d9f))
- **web-component:** Update Fields also rebuilds the TOC and tables of figures (by @ChristopherVR) ([c56b3ae](https://github.com/ChristopherVR/docx-viewer/commit/c56b3ae06e1d42f2440b715aa1d5b182eda8516e))
- Table and cell border editor (Home > Borders and Borders and Shading Apply to) (by @ChristopherVR) ([2597e65](https://github.com/ChristopherVR/docx-viewer/commit/2597e65dc1bb083222922d89c68492c345264955))
- Layout > Page Borders (model, pgBorders write, Print Layout drawing, dialog) (by @ChristopherVR) ([99c4159](https://github.com/ChristopherVR/docx-viewer/commit/99c41590fac6aa108a518f7682d56428b8dd6824))
- Insert > Text Box (inline, editable), plus text-box line display and image run merge fixes (by @ChristopherVR) ([4d76849](https://github.com/ChristopherVR/docx-viewer/commit/4d768492d4ec50661356d63ff1431d0397cf77dc))
- Layout > Watermark (text watermark VML in headers, Print Layout layer, dialog) (by @ChristopherVR) ([21ec767](https://github.com/ChristopherVR/docx-viewer/commit/21ec7671afec67ba76f8da9b56d5059594591269))
- **web-component:** Inserting a caption also refreshes REF results that quote captions (by @ChristopherVR) ([c86273a](https://github.com/ChristopherVR/docx-viewer/commit/c86273a6959a63e2a08ad9bde21146fea858b500))
- **bindings:** Re-export the document model from every editor package (by @ChristopherVR) ([abf4e79](https://github.com/ChristopherVR/docx-viewer/commit/abf4e796049ab6584047c0cb66ba6c0af5396259))
- **web-component:** Display SmartArt from its cached drawing (by @ChristopherVR) ([4d5424c](https://github.com/ChristopherVR/docx-viewer/commit/4d5424cda33dccf12f009d33d2115fbbc74729b6))

### Bug Fixes

- **tables:** Keep cell formatting on edit, render table styles, write valid new tables (by @ChristopherVR) ([bf63864](https://github.com/ChristopherVR/docx-viewer/commit/bf63864b2f04487832e2ab366a9fd9385cba81dd))
- **notes:** Show note references in the document's number format (by @ChristopherVR) ([857fde5](https://github.com/ChristopherVR/docx-viewer/commit/857fde5c5908c01fa886420b33b3dc2faa6212fa))
- **revisions:** Save editor revision ids as decimals (by @ChristopherVR) ([bebb185](https://github.com/ChristopherVR/docx-viewer/commit/bebb1859f05b11a7528e9065cd061645a7dc058e))
- **formatting:** Word toggle-property semantics and explicit offs (by @ChristopherVR) ([84ad25d](https://github.com/ChristopherVR/docx-viewer/commit/84ad25d34f2e86460451bbaf05c0e9c2d9b1dcae))
- **core:** Write schema-ordered numbering, sectPr and drawing extents; make schema test real (by @claude) ([2b95fb7](https://github.com/ChristopherVR/docx-viewer/commit/2b95fb7ec32a2eae925d29f06d1302dfd059e785))
- **web-component:** Fit ribbon labels per locale, split tab keys, relocalize presence (by @claude) ([0881cfc](https://github.com/ChristopherVR/docx-viewer/commit/0881cfcfaa6d1c7dcd995e74a33fa37f4187dfc0))
- **packaging:** Viewer Svelte component imports bindings helpers; typed svelte.d.ts (by @claude) ([ac57529](https://github.com/ChristopherVR/docx-viewer/commit/ac575296c0d13efb5ac4625c34ffd136b1e6c603))
- **web-component:** Keep Define New Multilevel List within 720px with marker font row (by @ChristopherVR) ([cfee3e7](https://github.com/ChristopherVR/docx-viewer/commit/cfee3e78ef6812f44c204a4f5b844cc7f0f066e4))

### Other

- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([b4d3e44](https://github.com/ChristopherVR/docx-viewer/commit/b4d3e44463195e949316d3ac6df7bf9137485b4a))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([3cddda1](https://github.com/ChristopherVR/docx-viewer/commit/3cddda15ca30ac85612049ae80d63d947a230470))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([98b3036](https://github.com/ChristopherVR/docx-viewer/commit/98b3036664d3dedccfeb36dcbd9c99e7826795ec))
- Interrupted parity workstream (see docs/outstanding-work.md) (by @ChristopherVR) ([52332bf](https://github.com/ChristopherVR/docx-viewer/commit/52332bf134d1c30a6ba63f7c61abbd79dad4b053))

### Refactor

- **core:** Remove writer helpers superseded by write-inline (by @ChristopherVR) ([f400024](https://github.com/ChristopherVR/docx-viewer/commit/f40002444f3c1dc14f1272e967c03c06865dabaa))
- **editor:** Move parts, page setup and list commands out of the component (by @ChristopherVR) ([bc8da60](https://github.com/ChristopherVR/docx-viewer/commit/bc8da6060474a78300616d83a55ec58c1d07b04a))
- Split ribbon, schema, model, adapter and collaboration modules (by @ChristopherVR) ([a3a1aee](https://github.com/ChristopherVR/docx-viewer/commit/a3a1aee3bf584c448f5de260b24a584dce346a60))
- Split range markers and body editor plugins into modules (by @ChristopherVR) ([7490a17](https://github.com/ChristopherVR/docx-viewer/commit/7490a17239c905b80c290b69402836fb704f50e6))
- **core:** Split oversized model and write modules (by @claude) ([b821cb6](https://github.com/ChristopherVR/docx-viewer/commit/b821cb64068d2b1887576770ea63de1fdcfbbc74))
- **web-component:** Split run mark construction out of run-adapter (by @claude) ([d4a2adf](https://github.com/ChristopherVR/docx-viewer/commit/d4a2adfe007d805912fa96bdf35e59d9dcc59243))
- **core,legacy:** Make source safe under noUncheckedIndexedAccess (by @claude) ([78b75e7](https://github.com/ChristopherVR/docx-viewer/commit/78b75e79a22e6d5e3aa0f3d9ebe068350143d74b))
- **web-component:** Make UI packages safe under noUncheckedIndexedAccess (by @claude) ([ec4a233](https://github.com/ChristopherVR/docx-viewer/commit/ec4a233783fe6066f1e8f814e671fb15638b8d82))
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
- Depend on the unscoped ooxml-core and ooxml-ui (by @ChristopherVR) ([c5be466](https://github.com/ChristopherVR/docx-viewer/commit/c5be466453eff0b35474b076558ba437f333b13e))

### Styling

- Apply oxfmt across the workspace (by @ChristopherVR) ([e9b3725](https://github.com/ChristopherVR/docx-viewer/commit/e9b3725582062b8140f76f7019912961b81b7ce8))
- **demo:** Widen style pickers and add a favicon (by @ChristopherVR) ([48c0044](https://github.com/ChristopherVR/docx-viewer/commit/48c0044ba1dadff52cd37460b440d1d9d6805ed1))

### Chores

- **types:** Enable exactOptionalPropertyTypes (by @claude) ([6f46221](https://github.com/ChristopherVR/docx-viewer/commit/6f4622160f37ae94a3bb5fa50710456b798dcaab))
- **types:** Enable noUncheckedIndexedAccess; keep test-support out of release builds (by @claude) ([203b289](https://github.com/ChristopherVR/docx-viewer/commit/203b289f0e8ca5d877cc6ff990ebc35755a853a3))
- Renamed ooxml packages [ci-skip] (by @ChristopherVR) ([3e75728](https://github.com/ChristopherVR/docx-viewer/commit/3e757282589b0327b90657d0d4c4e9350ce5025a))

