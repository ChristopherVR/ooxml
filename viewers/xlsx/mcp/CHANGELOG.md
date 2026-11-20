# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

For agents that read and edit workbooks through `xlsx-tools`, this release closes two ways a prompt-injected model could misuse the server. The write tools refuse formulas that would make Excel fetch data, run code or reach another file when the workbook is opened, and file paths that Windows would reinterpret, or that point at a network share, are refused before any filesystem call. Tool names are unchanged; the edit tools move to schema version 2. Upgrade together with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) has the before and after for each breaking change.

### Breaking changes

- `xlsx_set_cells`, `xlsx_propose_edits` and `xlsx_apply_proposal` (now schema version 2) refuse a formula that can reach outside the workbook: web and external data functions (`WEBSERVICE`, `IMAGE`, `HYPERLINK`, `RTD`, `STOCKHISTORY`, the `CUBE` functions, `PY`), XLL add-in calls (`_xll.`), the Excel 4.0 macro functions (`EVALUATE`, `EXEC`, `RUN`, `FOPEN` and similar), functions that send cell data to an AI service (`COPILOT`, `TRANSLATE`, `DETECTLANGUAGE`), references to another workbook or a file path, and formulas the parser cannot read. The error names the cell, and defined names are checked the same way. No tool parameter turns the check off: a server for trusted callers registers `createXlsxTools({ formulas: 'unrestricted' })` from `ooxml-core/automation` with `registerAutomationTools` instead of this package's `registerTools`.
- On Windows, file tools refuse path segments that contain `:`, `<`, `>`, `"`, `|`, `?` or `*`, that end in a dot or a space, or that are device names such as `NUL` or `COM1`, as well as `/x` (rooted on the current drive) and `C:x` (relative to a drive's current folder). On every platform they refuse UNC and device paths and paths that start with a backslash. Use a path relative to the root, or a full absolute path inside it, and rename files whose names Windows would reinterpret.
- Node.js 22.3 or newer is required (`"engines": { "node": ">=22.3" }`).

### Security

- A prompt-injected model can no longer write a formula that sends workbook data to a server when the file is opened in a spreadsheet application: such formulas are refused on every write path.
- Dangerous paths are refused on their text before any filesystem call, so a UNC path no longer reaches a remote host, every refused path returns the same error, and an alternate data stream name can no longer create a file past the extension check.
- Workbooks are read through the guarded zip reader and calculated under the per-formula budget of `ooxml-core`, so a small crafted file can no longer exhaust memory or stall a tool call.
- Releases are published by a pipeline in which no third-party code can push, tag or publish, with npm provenance, and the signed tarballs on each GitHub release are checked against the integrity the registry records.

### Changed

- An absolute path may reach the root through a link, a Windows short name or the real path of a linked root. On POSIX systems, names such as `a: b.xlsx`, `aux/` and `why?.xlsx` are ordinary files, and a backslash inside a name is not a separator.
- Recalculation after an edit uses the per-formula work budget of `ooxml-core`: a formula past it returns `#NUM!` while the others still compute.
- Built and tested against current releases of its dependencies.

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): the `xlsxTools` definitions, the formula policy and the path checks this server registers.
- [`ooxml-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/mcp/CHANGELOG.md): the combined Office MCP server, which registers these same tools.
- [`@christophervr/xlsx-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/core/CHANGELOG.md): the workbook model for your own headless code.
- [`@christophervr/xlsx-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/packages/react/CHANGELOG.md): the Excel editor for React; the other framework packages are listed in the [Excel viewer changelog](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/CHANGELOG.md).

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@1.1.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.0.1](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@1.0.1) - 2026-11-03

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([8e60b14](https://github.com/ChristopherVR/ooxml/commit/8e60b14e1d8ee3a1cb507670caa15a611632ef26))
- **deps:** Merge production dependency updates ([#236](https://github.com/ChristopherVR/ooxml/issues/236)) ([7eb137b](https://github.com/ChristopherVR/ooxml/commit/7eb137b03906ded043ce0e4adf83e01fcdaa44a9))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@1.0.0) - 2026-10-29

### Features

- **xlsx:** Commit xlsx-core and xlsx-viewer-mcp to a stable 1.0 api ([e458ecf](https://github.com/ChristopherVR/ooxml/commit/e458ecf368421096ccffb366cc1e0ba17f438b77))

## [0.2.2](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.2.2) - 2026-10-22

### Testing

- **xlsx:** Run the reviewed-edit tools end to end over mcp ([9f66379](https://github.com/ChristopherVR/ooxml/commit/9f66379b96ec2708d032294d1ef2f8557e24f284))

## [0.2.1](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.2.1) - 2026-10-22

### Refactor

- **mcp:** Register every format's tools through core automation ([9a28bfa](https://github.com/ChristopherVR/ooxml/commit/9a28bfa684167472fff27f04a371c37d494a5fd5))

## [0.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.2.0) - 2026-10-21

### Features

- **xlsx:** Add native framework building blocks ([9689a65](https://github.com/ChristopherVR/ooxml/commit/9689a6520682e9db18f8cbee78d8434ad114b097))

## [0.1.11](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.1.11) - 2026-10-20

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([02d7713](https://github.com/ChristopherVR/ooxml/commit/02d7713ceb301921fd6c7706138d64f4c10a9d38))
- **deps:** Merge dependency update ([#94](https://github.com/ChristopherVR/ooxml/issues/94)) ([e381aba](https://github.com/ChristopherVR/ooxml/commit/e381abaa11794233cdea85a9dc6d098d4cb338c4))

## [0.1.10](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.1.10) - 2026-10-15

### Bug Fixes

- **automation:** Add --root and untrusted-content notes to MCP servers ([0da2c7d](https://github.com/ChristopherVR/ooxml/commit/0da2c7d1ec6655f649a608ded1f36dea41f34962))

## [0.1.9](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.1.9) - 2026-10-10

### Styling

- Apply oxfmt to files drifted from the formatter ([daeabf2](https://github.com/ChristopherVR/ooxml/commit/daeabf26125af866f90359ac8a5716234d5b8b36))

## [0.1.8](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.1.8) - 2026-10-09

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

## [0.1.7](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.1.7) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.1.6](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.1.6) - 2026-10-08

### Bug Fixes

- **ci:** Recover unpublished releases and validate manifests ([ae190df](https://github.com/ChristopherVR/ooxml/commit/ae190df2352283303858719fe1db5e58b86903fb))

## [0.1.5](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.1.5) - 2026-10-07

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.1.4](https://github.com/ChristopherVR/ooxml/releases/tag/xlsx-viewer-mcp@0.1.4) - 2026-10-05

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs re-releases a package whenever any of its files, or an internal
dependency it ships against, changes, not only on conventional commits.

## [0.1.3](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-viewer-mcp@0.1.3) - 2026-10-03

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([afbbf9c](https://github.com/ChristopherVR/xlsx-viewer/commit/afbbf9c9a88f70e0afeec47da60e6cd9339c76ba))

## [0.1.2](https://github.com/ChristopherVR/xlsx-viewer/releases/tag/xlsx-viewer-mcp@0.1.2) - 2026-10-03

### Features

- **release:** Add core backed Excel MCP tools (by @ChristopherVR) ([775cc2b](https://github.com/ChristopherVR/xlsx-viewer/commit/775cc2b59d7a31437f7f528069f957e2ba165581))

### Bug Fixes

- **release:** Stage planned package versions and changelogs (by @ChristopherVR) ([66e7c29](https://github.com/ChristopherVR/xlsx-viewer/commit/66e7c29b4c8d5109c358ac1df77a4afad91a86f1))


