# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/@christophervr/xlsx-core@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

`@christophervr/xlsx-core` re-exports `ooxml-core/xlsx` and `ooxml-core/xlsx/load`, so this release brings the 2.0 workbook model to headless code: servers, workers and scripts that read, calculate and write workbooks. Calculation is bounded per formula and can be bounded per recalculation, so an untrusted workbook can no longer hang a process, and lookups and aggregates over large ranges scale with the edit. Signatures are reported as a verified state, hyperlinks keep every safe scheme, and Node.js 22.3 is the minimum. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) has the before and after for each breaking change.

### Breaking changes

- Node.js 22.3 or newer is required (`"engines": { "node": ">=22.3" }`).
- `Workbook.signatures` is now a `SignatureSummary` (`state` and `signatures`) instead of `{ count, parts }`, and loading a signed workbook no longer adds an English signature message to `warnings`. Read `signatures.state` and word it with `signatureStateText` from `ooxml-core/i18n`.
- Each formula now has its own work and text budget, next to an array budget that used to cover a whole recalculation and now covers one formula. A formula past it, such as `MMULT` of two 1,500 by 1,500 arrays or `COUNTIF` over 20,000 cells repeated for 20,000 criteria, returns `#NUM!`, usually before any of the work runs, and every other formula still computes. Workbooks Excel calculates in ordinary time stay well inside it.
- `CalcEngine` has a new `stopped` member, so code that implements the interface itself, such as a test double, must add it.
- The hyperlink policy now stores every well-formed scheme that cannot run script, such as `tel:`, `sip:`, `msteams:` and `onenote:`, and still opens only http, https and mailto unless a host opts in. The `unsupported-scheme` rejection reason is gone, so drop any code that checks for it.

### Security

- Calculation can no longer be made to hang or crash a process: each formula is bounded, `REGEXTEST`, `REGEXEXTRACT` and `REGEXREPLACE` run on the engine's own matcher, so no pattern can freeze a recalculation (one that would take too long returns `#NUM!`), and a recalculation stops between formulas rather than storing more than 50 million array cells or 512 million characters of text.
- Every package part is read through the guarded zip reader, including signature listing and embedded chart workbooks, so a small file can no longer inflate without limit. Signature listing ignores a signature part larger than 16 MiB.
- A workbook is reported as signed only after its signature is verified, never because it merely contains a signature part.
- A damaged or crafted legacy `.xls` workbook can no longer keep `loadWorkbook` busy for seconds (bundled `@christophervr/ole2` 0.14.1). A 9.7 KB workbook that took over 20 seconds now reads in about 5 ms.
- XML scans that took quadratic time on runs of empty or unclosed elements, such as header and footer pictures and the declared-font scan, are now linear.
- Chart point counts and indices, and ChartEx histogram bins (at most 1,000), no longer decide how much memory is allocated.
- For hosts that share workbooks through `ooxml-core/xlsx/collab`, a hyperlink a peer sends keeps every target the loader keeps and leaves out only a script or disguised one, and the left-out value is reported.
- Releases are published by a pipeline in which no third-party code can push, tag or publish, with npm provenance, and the signed tarballs on each GitHub release are checked against the integrity the registry records.

### Added

- `CalcEngineOptions.timeLimitMs` bounds how long one recalculation may run, and `CalcEngineOptions.workCap` how much work one formula may need. Past either, the recalculation stops between formulas, `CalcEngine.stopped` says why (`time`, `work` or `resources`), the formulas it did not reach keep their previous values, and the next recalculation computes everything.
- `calculateOnOpen(workbook, { timeLimitMs })` runs the calculation a workbook asks for on open (`calcPr fullCalcOnLoad`) and reports whether it ran and why it stopped. `needsCalculationOnOpen`, `calculateOnOpenCells` and `applyCalculatedCells` let you calculate in a worker and copy only the written cells back.
- `formulaPolicyIssue`, `cellInputPolicyIssue` and `RESTRICTED_FORMULA_FUNCTIONS` say why a formula reaches outside the workbook (web and external data functions, macro functions, add-in calls, AI functions, links to other workbooks). They read the formula with the formula parser, so a function passed by name to `MAP`, `LET` or `LAMBDA` or written with a storage prefix is found too.
- `describeCell` and `accessibleWindow` describe cells and the visible window for screen readers, and `ChartObject.description` holds a chart's alt text.

### Changed

- Charts keep the point counts and points their file declares, however large, through reading, editing and saving. Views past the point budget are reported with a `CHART_POINTS_LIMITED` load warning and stay fully editable.
- Built and tested against current releases of its dependencies, including `emf-converter` 5.

### Fixed

- Workbooks with `LET`, `LAMBDA` or the functions that take a `LAMBDA` (`REDUCE`, `MAP`, `SCAN`, `BYROW` and the others) save their names with the `_xlpm.` prefix Excel requires, so Excel opens them; reading removes the prefix again.
- Formulas compare text as Excel does: comparisons, criteria functions, `XMATCH`, `XLOOKUP`, `UNIQUE`, `SORT` and `SWITCH` treat "ß" as "ss" and ligatures as their letters while staying accent-, width- and kana-sensitive, and `MATCH`, `VLOOKUP`, `HLOOKUP`, wildcards and `SEARCH` compare character by character.
- Exact lookups compare numbers exactly, so `MATCH(0.3, ...)` no longer finds `0.1+0.2`, while `=` and `COUNTIF` still treat the two as equal.
- In criteria, `~` is plain text unless the criterion holds `*` or `?`, a blank criterion cell matches zeros, and `DCOUNTA` counts empty text. `&`, `CONCATENATE`, `CONCAT`, `TEXTJOIN`, `ARRAYTOTEXT`, `TEXT`, `BASE` and `LOWER` follow Excel's length and case rules.
- `REGEX` patterns with nested repeats, such as `(a+)+`, now work instead of returning `#VALUE!`, long texts are no longer refused, and `REGEXREPLACE` with an occurrence uses that match's own groups.
- In iterative calculation, `SUMIF`, `COUNTIF` and the other criteria functions read current values on every pass instead of keeping their first read.
- A chart keeps its alt text when the drawing is saved again after the chart was moved.
- For hosts using `ooxml-core/xlsx/collab`, a collaborator now receives `tel:`, `sip:`, `msteams:`, `onenote:` and other hyperlink targets a workbook holds.

### Performance

- Exact `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `XMATCH`, `MATCH` and `INDEX`/`MATCH` searches over the same range share an index within a recalculation, in both directions, with results identical to a scan.
- Identical `SUM`, `PRODUCT`, `COUNT`, `COUNTA`, `AVERAGE`, `MIN`, `MAX` and `MEDIAN` calls over the same large ranges are computed once per recalculation.
- An edit that changes a row's automatic height recalculates only what depends on the edited cells. Hiding rows, setting a height by hand and other structural edits still recalculate fully, so `SUBTOTAL` and `AGGREGATE` keep following hidden rows.
- `MMULT` runs about three and a half times faster, and 1,000 nearest-match `XLOOKUP` or `XMATCH` lookups in 100,000 rows take about a second instead of six.
- `forEachCellInRange` costs what the range covers, volatile formulas are tracked as a list, and selecting a whole column no longer lists every row to decide whether Ungroup is available.

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

