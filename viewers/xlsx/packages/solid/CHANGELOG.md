# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

`SpreadsheetEditor` now renders during server rendering and hydrates the server's element, instead of throwing `document is not defined`, and keeps the same props, callbacks and `editorRef` handle. Underneath, workbooks that calculate on open no longer freeze the page, crafted files and formulas can no longer hang or exhaust the tab, a closed editor is freed at once, and the grid reads as a real table to screen readers with Excel's F6 regions. A few defaults change: dark mode comes from the theme instead of CSS `color-scheme`, signatures are reported as a verified state, and Node.js 22.3 is the minimum. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) has the before and after for each breaking change.

### Breaking changes

- Node.js 22.3 or newer is required (`"engines": { "node": ">=22.3" }`), which matters for server rendering, tests and build scripts that import the package.
- The editor no longer inherits the page's CSS `color-scheme`. It reads the new `--office-color-scheme` token, which is `light` unless the theme's dark mode applies (`data-office-theme="dark"`, or the system's dark preference where no `data-office-theme="light"` overrides it). A container that sets only `color-scheme: dark` now gets a light editor: also set `--office-color-scheme: dark` on it, or use `data-office-theme="dark"`, which sets both.
- `--xve-input` now defaults to `--office-control-border` (`#8a8886`) instead of `--office-border`, and it colours the shared form controls inside the editor. To restyle field boundaries, set `--office-control-border` or `--xve-input`, and keep it at 3:1 against the background.
- Changing `surfaceOnly` on an `<xlsx-editor>` that is already attached to the page now throws, as it does on `docx-editor`. This only affects code that creates or reaches the element itself (for example through `onReady`); set it before the element is appended.
- An editor that shares a room now leaves it when it is removed from the page and not attached again before the current task ends, and joins the same room again when it is attached later. Moving it between parents in one task keeps it in the room. To hide an editor and keep it syncing, hide it with CSS instead of detaching it.
- Collaboration updates denser than one struct per 4 bytes, or one deleted range per 2 bytes, are refused before they are decoded. The budget grows with the update, so the first sync of a real workbook stays well under it. A host that passes its own limits replaces `DEFAULT_UPDATE_LIMITS` from `ooxml-core/collab` with `defaultUpdateLimits(bytes)`.
- Hyperlinks: every well-formed scheme that cannot run script, such as `tel:`, `sip:`, `msteams:` and `onenote:`, is now stored, and Insert Link accepts these addresses as Excel does. Only http, https and mailto still open unless the host opts in. The `unsupported-scheme` rejection reason is gone, so drop any code that checks for it.
- `Workbook.signatures` is now a `SignatureSummary` (`state` and `signatures`) instead of `{ count, parts }`, and loading a signed workbook no longer adds an English signature message to the warnings `onWorkbookWarning` reports. Read `signatures.state` and word it with `signatureStateText` from `ooxml-core/i18n`.
- Each formula now has its own work and text budget. A formula past it, such as `MMULT` of two 1,500 by 1,500 arrays, returns `#NUM!` while every other formula still computes; workbooks Excel calculates in ordinary time stay well inside it. Code that implements the `CalcEngine` interface itself must add the new `stopped` member.

### Security

- A workbook is never shown as signed just because it contains a signature part. The status bar, the notice under the ribbon and File > Info show the verified state.
- A small crafted workbook can no longer hang or crash the editor while it opens. Each formula is bounded, `REGEXTEST`, `REGEXEXTRACT` and `REGEXREPLACE` run on a matcher that cannot freeze a recalculation, the calculation on open runs off the page under a time limit, and what a recalculation stores is bounded.
- Every package part is read through the guarded zip reader, including signature listing and embedded chart workbooks, so a small file can no longer inflate without limit.
- A chart of a few hundred bytes can no longer exhaust memory: point counts and indices no longer decide how much is allocated, and a ChartEx histogram draws at most 1,000 bins.
- A damaged or crafted legacy `.xls` workbook can no longer keep the reader busy for seconds. A 9.7 KB workbook that took over 20 seconds now reads in about 5 ms.
- XML scans that took quadratic time on runs of empty or unclosed elements, such as header and footer pictures and the declared-font scan, are now linear.
- SVG pictures are shown from `data:` URLs instead of object URLs of your origin, so opening one in its own tab can no longer run document-controlled script there.
- Values a collaborator sends are checked. A shared hyperlink keeps every target the loader keeps and leaves out only a script or disguised one, and the left-out value is reported rather than dropped silently.
- Releases are published by a pipeline in which no third-party code can push, tag or publish, with npm provenance, and the signed tarballs on each GitHub release are checked against the integrity the registry records.

### Added

- `customization.ribbon.hiddenTabs` hides native ribbon tabs by id (`ExcelRibbonTabId` in `ooxml-core/xlsx/ui`). `'file'` removes the File tab and its backstage at every width, with its key tip and every other route in, while Save, its shortcuts, host `storage` and `file-command` keep working; Tell me leaves out commands that live only on a hidden tab. The binding forwards it with the rest of `customization`.
- A signed workbook shows its verified state in the status bar, in a notice under the ribbon and on File > Info: Signed, Signed (signer not verified), Invalid signature or Signature not verified, in every editor language. The notice says that saving changes removes the signatures. Nothing is shown for an unsigned workbook.
- When sharing stops because an update was refused, a toast says why in the display language, and the editor says once per session when content a collaborator sent was left out.
- Themes take an optional `primaryInk` (`--xve-primary-ink`), the primary colour used as text. It defaults to `primary`, and the dark presets lift it.
- The re-exported model adds `calculateOnOpen`, `needsCalculationOnOpen`, `CalcEngineOptions.timeLimitMs` and `workCap` for bounded recalculation, `formulaPolicyIssue` and `cellInputPolicyIssue` to tell why a formula reaches outside the workbook, and `describeCell` and `accessibleWindow` for screen reader output.

### Changed

- Opening a workbook that asks for a full calculation on open (`fullCalcOnLoad`) no longer freezes the page. The editor calculates on the page for at most as long as reading took; past that, a worker (`xlsx-open-worker.js`, shipped in `ooxml-ui`) runs the calculation and sends back only the cells it wrote. It stops after ten seconds, unfinished formulas show their saved values, and a notice says that F9 recalculates. Files that need no calculation open as fast as before. Bundlers that understand `new Worker(new URL(..., import.meta.url))` pick the worker up automatically; a Content Security Policy must allow `worker-src 'self'`, and without workers the calculation runs on the page under the same limit.
- Charts keep the point counts and points their file declares, however large, through viewing, editing and saving. A view past the point budget shows fewer trailing blanks or only its first points, the load reports `CHART_POINTS_LIMITED`, and the chart stays fully editable.
- When an editor is removed from the page, it adds one constructed stylesheet of its own to `document.adoptedStyleSheets` (on first use) and changes its rule, which matches nothing; this empties a browser style cache that kept closed editors alive. A host that replaces `document.adoptedStyleSheets` may drop it; the next removal adds it again.
- `solid-js` 1.9 is still required. Solid 2 is not supported yet, because the stable Solid compiler plugin for Vite does not support it.
- Built and tested against current releases of its dependencies, including `emf-converter` 5. `emf-converter` now loads image decoders on demand from `dist/browser-chunks/` and `dist/node-chunks/`; package managers and bundlers need no change, but if you copy compiled files by hand, deploy those folders next to the entry files.
- Supported browsers are the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least ([browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md)).

### Fixed

- `SpreadsheetEditor` renders during server rendering instead of throwing `document is not defined`, and the browser reuses the server's element when it hydrates.
- A closed editor no longer stays in memory. Print, font and ribbon listeners are released on removal and restored when the same element is attached again; menus, popups, the AutoFilter menu and form control lists still open are closed; and an editor removed while sharing a room, or right after a paint, is freed at once. In Chrome and Edge, an editor whose sheet drew a clipped, masked or filtered shape, picture or form control is freed too.
- Removing one Excel editor no longer closes a ribbon popover or overflow menu open in another editor on the same page.
- Ctrl+PageUp and Ctrl+PageDown switch sheets instead of moving the active cell. F6 and Shift+F6 move through every part of the window, including the status bar. Dialogs keep Tab inside, close on Escape wherever focus is, and return focus to the control that opened them.
- The sheet no longer changes with the interface theme: gridlines use `--office-canvas-gridline` (Excel's `#e1e1e1`), so dark mode no longer draws dark gridlines on white paper. The active sheet tab uses chrome colours and reads at 4.5:1 in dark mode, and `--xve-selection-border` colours the selected headers and the active sheet name again.
- Form control boundaries reach 3:1 against the background, including the formula bar, the Name Box, dialog fields and the ribbon's combo boxes and drop-downs.
- Workbooks with `LET`, `LAMBDA` or the functions that take a `LAMBDA` save their names with the `_xlpm.` prefix, so Excel opens them.
- Formulas compare text as Excel does (for example "ß" matches "ss" in comparisons and criteria), exact lookups compare numbers exactly (`MATCH(0.3, ...)` no longer finds `0.1+0.2`), and criteria, text length limits and `REGEXREPLACE` occurrences follow Excel. `REGEX` patterns with nested repeats, such as `(a+)+`, now work.
- In iterative calculation, `SUMIF`, `COUNTIF` and the other criteria functions read current values on every pass.
- A collaborator now receives `tel:`, `sip:`, `msteams:`, `onenote:` and other hyperlink targets, the first sync of a large workbook (about 500,000 mixed cells) is no longer refused, and a refused update is explained in the display language.
- A chart keeps its alt text when the drawing is saved again after the chart was moved.

### Performance

- Opening and closing workbooks in a long session no longer grows memory by the size of the editor each time: in the Office suite, each cycle went from about 13 MB and 4,000 DOM nodes to about 0.2 MB and none.
- Recalculation scales with the edit: exact `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `XMATCH`, `MATCH` and `INDEX`/`MATCH` searches over the same range share an index, identical `SUM`, `AVERAGE`, `MIN`, `MAX` and similar calls over the same large ranges are computed once, and an edit that changes a row's automatic height recalculates only what depends on the edited cells.
- `MMULT` runs about three and a half times faster, and 1,000 nearest-match `XLOOKUP` or `XMATCH` lookups in 100,000 rows take about a second instead of six.
- Undo and redo of Replace All on large sheets, finding volatile formulas and selecting whole columns no longer visit the whole sheet or workbook. On sheets with more than 20,000 stored rows the status bar fills in sum, average and count just after the selection is drawn.
- The editor no longer loads the whole `ooxml-ui` root entry (and with it the Teams app and the other products' editors) to define its dialogs and icons.

### Accessibility

- Screen readers can explore the grid as a table: visible cells are grid cells in rows with column and row headers, indexed as the sheet is (B2 is row 2, column 2). The active cell is announced with its column, formula and comment, and stays correct after an undo, a reload or a recalculation; pictures are announced once, by their alt text.
- A chart is named by its alt text, falling back to its title, and screen readers can read its series and values.
- The Tell me box, the Name Box, the formula bar, search fields and the File tab show a visible focus ring with at least 3:1 contrast in light and dark themes. The formula bar's Insert Function button keeps its name, and its fx glyph is an icon.
- In forced-colours mode, applied filter buttons, slicer selections, dialog tabs and options, function suggestions, contextual ribbon tabs and sheet tabs keep a system highlight, and gridlines use the system's muted text colour.
- Every change is checked in a real browser before release: keyboard regions and focus, axe in light mode, dark mode and on a phone, reflow at 320 by 256 pixels, and that a saved workbook keeps its alt text and language.

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): the workbook model, formula engine, reader and writer behind the editor.
- [`ooxml-ui` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/ui/CHANGELOG.md): the `<xlsx-editor>` element, its open worker and the shared controls.
- [`@christophervr/xlsx-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/core/CHANGELOG.md): the model API this package re-exports.
- [`@christophervr/xlsx-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/react/CHANGELOG.md): the same editor for React.
- [`xlsx-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/vue/CHANGELOG.md): the same editor for Vue.
- [`xlsx-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/angular/CHANGELOG.md): the same editor for Angular.
- [`xlsx-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/svelte/CHANGELOG.md): the same editor for Svelte.
- [`xlsx-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/vanilla/CHANGELOG.md): `mountEditor` and the `<xlsx-editor>` element without a framework.
- [`xlsx-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/CHANGELOG.md): the Excel tools for AI agents, which now refuse formulas that reach outside the workbook.

## [1.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.5.0) - 2026-11-13

### Features

- **xlsx:** Host-owned context menus and workflow operations ([18854c4](https://github.com/ChristopherVR/ooxml/commit/18854c444095dfaea70f087f3c868211316436a7))
- **xlsx:** Filter and extend the native context menu ([554d354](https://github.com/ChristopherVR/ooxml/commit/554d354d5f2ec8ce9146ce43645301b379f71d11))
- **xlsx:** Make replace validation configurable ([093f478](https://github.com/ChristopherVR/ooxml/commit/093f4782640506ab8274134f70ad4abb8a2e4e50))

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.4.1](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.4.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.4.0) - 2026-11-05

### Features

- **xlsx:** Open workbooks from cloud places in File > Open ([f1b19e4](https://github.com/ChristopherVR/ooxml/commit/f1b19e495f8ac1fe43ed1a540f76ce5c6cbc2710))
- **ui:** Expose openCloudFile on the word, excel and visio bindings ([b4837c7](https://github.com/ChristopherVR/ooxml/commit/b4837c76398c125755458861ce62679d7c179003))

## [1.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.3.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.2.2](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.2.2) - 2026-10-29

## [1.2.1](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.2.1) - 2026-10-23

### Other

- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))

### Refactor

- **docx:** Build the Word bindings on ooxml-ui/bindings ([90e8131](https://github.com/ChristopherVR/ooxml/commit/90e81310800081ad484e93a3bab25806a07bc3e4))
- **xlsx:** Build the Excel bindings on ooxml-ui/bindings ([31ca812](https://github.com/ChristopherVR/ooxml/commit/31ca812ba4513f0b0e0db643ab4b36643cbe62e5))

## [1.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.2.0) - 2026-10-22

### Features

- **xlsx:** Route save and open through host-owned storage ([ef768a6](https://github.com/ChristopherVR/ooxml/commit/ef768a684f558275071be347d68406503aef8c23))
- **xlsx:** Forward file-command, workbook, ribbon and sheet events ([5aaf6db](https://github.com/ChristopherVR/ooxml/commit/5aaf6dbdc767a1ff8c30b1b8619c0dcb09cfdc38))
- **xlsx:** Take the shared customization contract in <xlsx-editor> ([f96a544](https://github.com/ChristopherVR/ooxml/commit/f96a5449a6a36fba59997197b484ddefb4f2d8e8))
- **embedding:** Land the host customization, read-only and storage contracts ([822a6a0](https://github.com/ChristopherVR/ooxml/commit/822a6a073840b9e9d1671476542322b28332dc42))

### Bug Fixes

- **xlsx:** Report read-only with the lock and host switch ([782adb9](https://github.com/ChristopherVR/ooxml/commit/782adb95a4c8ad4dded13d73e8a5fa5aee2c8666))

### Other

- Integrate the customization contract into host storage ([629554b](https://github.com/ChristopherVR/ooxml/commit/629554b3030728bfe16e11ac8ae5b2530bb28a17))

### Styling

- **xlsx:** Format the react binding test ([a2e8b70](https://github.com/ChristopherVR/ooxml/commit/a2e8b70fe6f25b51b02cbd612675f492f247b2fc))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.1.0) - 2026-10-22

### Features

- **xlsx:** Let hosts own the data-validation alert in building blocks ([fcbe367](https://github.com/ChristopherVR/ooxml/commit/fcbe367f5cfe366cc425796e3fdff9c31aa99e25))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@1.0.0) - 2026-10-22

### Features

- **xlsx:** Complete the building-block toolbar and add vanilla ([af8114f](https://github.com/ChristopherVR/ooxml/commit/af8114ff92e023858cd5a6d2a9922d6dbb2ee111))
- **xlsx:** Make the building blocks read-only unless canEdit is true ([efe6158](https://github.com/ChristopherVR/ooxml/commit/efe6158de251e1e34694335292755cfec96d640b))

## [0.7.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.7.0) - 2026-10-21

### Features

- **xlsx:** Expose composable spreadsheet building blocks ([9aef499](https://github.com/ChristopherVR/ooxml/commit/9aef499f76196be6dd008ae873ac8f0d4caeeb58))
- **xlsx:** Add native framework building blocks ([9689a65](https://github.com/ChristopherVR/ooxml/commit/9689a6520682e9db18f8cbee78d8434ad114b097))
- **xlsx:** Share office themes across editor building blocks ([a791ec2](https://github.com/ChristopherVR/ooxml/commit/a791ec2e68956921006c719397803e603286a8c0))

### Bug Fixes

- **deps:** Align private bindings with their react test host ([f4152ef](https://github.com/ChristopherVR/ooxml/commit/f4152ef04ccdc3e5c9d21ecf59f485b56de2ec52))
- **ui:** Harden composition lifecycle and consumer checks ([6aa52be](https://github.com/ChristopherVR/ooxml/commit/6aa52be88175ca4f46f6f3b1c223ef06a8f1f915))
- **ui:** Preserve live formula drafts and runtime compatibility ([3a7d4f1](https://github.com/ChristopherVR/ooxml/commit/3a7d4f184658c4d69b04bbd36bab7afa3896dd9e))

## [0.6.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.6.0) - 2026-10-10

### Features

- **xlsx:** Make ribbon add-in tabs a prop and event of every binding ([691f553](https://github.com/ChristopherVR/ooxml/commit/691f5531b9f5b4a5a5ac0645b43c8c0fc55be5d8))

## [0.5.1](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.5.1) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.5.0) - 2026-10-08

### Features

- **xlsx:** Expose collaboration in every binding ([2983ccc](https://github.com/ChristopherVR/ooxml/commit/2983ccc1e5d72749a31d2540e8f57b637f282b3e))

## [0.4.6](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.4.6) - 2026-10-08

## [0.4.5](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.4.5) - 2026-10-07

## [0.4.4](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.4.4) - 2026-10-07

## [0.4.3](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.4.3) - 2026-10-07

## [0.4.2](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.4.2) - 2026-10-06

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.4.1](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-solid-viewer@0.4.1) - 2026-10-05

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

## [0.4.0](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-solid-viewer@0.4.0) - 2026-10-05

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

### Build & CI

- **deps:** Adopt ooxml-core 0.17.0 and ooxml-ui 0.25.0 (by @ChristopherVR) ([49705cd](https://github.com/ChristopherVR/xlsx-viewer/commit/49705cd85cae5c14275ef15a544c689a3867abe4))
- **deps:** Adopt ooxml-ui 0.26.0 (by @ChristopherVR) ([94984f9](https://github.com/ChristopherVR/xlsx-viewer/commit/94984f959c1c75265294e5c49e3c22071e1f54f4))
- **deps:** Adopt ooxml-ui 0.27.0 (by @ChristopherVR) ([f24314b](https://github.com/ChristopherVR/xlsx-viewer/commit/f24314b56e662a05c6095338f32779536cc0d465))
- **deps:** Adopt ooxml-core 0.18.0 and ooxml-ui 0.27.1 (by @ChristopherVR) ([6e66a43](https://github.com/ChristopherVR/xlsx-viewer/commit/6e66a43e8966eedc4bd19934e27085be76d33a95))
- **deps:** Adopt ooxml-core 0.19.0 (by @ChristopherVR) ([d981768](https://github.com/ChristopherVR/xlsx-viewer/commit/d981768e1f6c5f612ea6c69aaf9507933a5a694f))

## [0.3.4](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-solid-viewer@0.3.4) - 2026-10-05

### Refactor

- **web-component:** Use the shared office-ui-context-menu (by @ChristopherVR) ([6aa323c](https://github.com/ChristopherVR/xlsx-viewer/commit/6aa323ce664715552c7e35da4c8ef6611b1d1bc4))

### Build & CI

- **deps:** Adopt ooxml-core 0.14.1 and ooxml-ui 0.20 (by @ChristopherVR) ([2da396d](https://github.com/ChristopherVR/xlsx-viewer/commit/2da396d66cad5cb74915d665258ac8216c3cd602))
- **deps:** Adopt ooxml-core 0.15.0 and ooxml-ui 0.21.0 (by @ChristopherVR) ([6cbd2b5](https://github.com/ChristopherVR/xlsx-viewer/commit/6cbd2b5ad65734a3d9ce7f2c8feff93d975eafe5))

## [0.3.3](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-solid-viewer@0.3.3) - 2026-10-03

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([afbbf9c](https://github.com/ChristopherVR/xlsx-viewer/commit/afbbf9c9a88f70e0afeec47da60e6cd9339c76ba))

## [0.3.2](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-solid-viewer@0.3.2) - 2026-10-03

### Bug Fixes

- **release:** Stage planned package versions and changelogs (by @ChristopherVR) ([66e7c29](https://github.com/ChristopherVR/xlsx-viewer/commit/66e7c29b4c8d5109c358ac1df77a4afad91a86f1))

## [0.3.0](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-solid-viewer@0.3.0) - 2026-10-03

### Features

- **web-component:** Adopt core encryption and workbook features (by @ChristopherVR) ([857348f](https://github.com/ChristopherVR/xlsx-viewer/commit/857348fb817ce99d773c8d4dad6258ff1dc93384))

## [0.2.0](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-solid-viewer@0.2.0) - 2026-10-03

### Features

- **bindings:** Keep in-editor changes and add readOnly/ribbon callbacks (by @ChristopherVR) ([6196ce1](https://github.com/ChristopherVR/xlsx-viewer/commit/6196ce10a9145bb1a409bd694086b772740045fa))

### Bug Fixes

- **web-component:** Restore sheet selections and fire selection once (by @ChristopherVR) ([fdfe205](https://github.com/ChristopherVR/xlsx-viewer/commit/fdfe205cd0a3bdb3ad575447a5bd842591eae942))
- **deps:** Depend on ooxml-core ^0.9.0 and ooxml-ui ^0.1.8 (by @ChristopherVR) ([e2a8ae6](https://github.com/ChristopherVR/xlsx-viewer/commit/e2a8ae63533262f4fef4fb8c93feadd3ca3b15b8))

## [0.1.1](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-solid-viewer@0.1.1) - 2026-10-03

### Features

- Add the Excel viewer and editor (by @ChristopherVR) ([408fe71](https://github.com/ChristopherVR/xlsx-viewer/commit/408fe71adc283e887cca4e873673daa6934957d5))

