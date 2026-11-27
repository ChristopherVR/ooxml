# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@2.0.1) - Unreleased

<!-- release-notes: 2.5 -->

The combined server's tools, schemas and bounds are unchanged in this release; this version ships its rewritten README, which says what the server does, what it does not do and what it sends over the network. It is published against `ooxml-core` 2.5.0, whose `ooxml-core/automation` definitions it runs, so the documents its tools save gain that release's writer fixes. Install it with the packages under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.5.md) has nothing to change for MCP servers.

**[Introducing OOXML 2.5](https://christophervr.github.io/ooxml/releases/2.5/)** tells the story of this release. Every change to the shared UI and MCP packages is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.5/#changes-ui-mcp) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.5.md#changes-ui-mcp)).

### Released together

- [`ooxml-core` 2.5.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): defines every tool, its schema and its bounds in `ooxml-core/automation`.
- [`docx-viewer-mcp` 2.0.1](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/mcp/CHANGELOG.md): the Word tools.
- [`xlsx-viewer-mcp` 2.0.1](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/CHANGELOG.md): the Excel tools.
- [`visio-viewer-mcp` 2.0.1](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/mcp/CHANGELOG.md): the Visio tools.
- [`pptx-viewer-mcp` 3.0.1](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/tools/CHANGELOG.md): the PowerPoint tools, which now report their own package version.
- [`ooxml-ui` 2.5.0](https://github.com/ChristopherVR/ooxml/blob/main/src/ui/CHANGELOG.md): the editors whose assistants run the same tool definitions on an open document.

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@2.0.0) - 2026-11-21

<!-- release-notes: 2.0 -->

This release makes the combined server safer to hand to an agent you do not fully trust. Excel writes refuse formulas that could send workbook data out of the file, file tools refuse paths that could leave the document root before they touch the disk, and the PowerPoint theme and metadata tools now save what they change. The server needs Node.js 22.3 or newer. Install it with the format packages released alongside it, listed under "Released together"; the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md) covers each breaking change.

**[Introducing OOXML 2.0](https://christophervr.github.io/ooxml/releases/2.0/)** tells the story of this release. Every change to the shared UI and MCP packages is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.0/#changes-ui-mcp) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.0.md#changes-ui-mcp)).

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): defines every tool, its schema and its bounds in `ooxml-core/automation`.
- [`docx-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/docx/mcp/CHANGELOG.md): the Word tools.
- [`xlsx-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/xlsx/mcp/CHANGELOG.md): the Excel tools, with the formula check.
- [`visio-viewer-mcp` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/visio/mcp/CHANGELOG.md): the Visio tools.
- [`pptx-viewer-mcp` 3.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/pptx/packages/tools/CHANGELOG.md): the PowerPoint tools, with the theme and metadata fixes.
- [`ooxml-ui` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/ui/CHANGELOG.md): the editors whose assistants run the same tool definitions on an open document.

## [1.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@1.1.2) - 2026-11-13

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@1.1.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@1.1.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.0.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@1.0.1) - 2026-11-03

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([8e60b14](https://github.com/ChristopherVR/ooxml/commit/8e60b14e1d8ee3a1cb507670caa15a611632ef26))
- **deps:** Merge production dependency updates ([#236](https://github.com/ChristopherVR/ooxml/issues/236)) ([7eb137b](https://github.com/ChristopherVR/ooxml/commit/7eb137b03906ded043ce0e4adf83e01fcdaa44a9))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@1.0.0) - 2026-10-29

### Features

- **automation:** Release ooxml-mcp 1.0 with versioned tool schemas ([e07b773](https://github.com/ChristopherVR/ooxml/commit/e07b77322730c4b427f29ed6e3479364aea8f54f))

### Other

- Bring origin/main into the embedding work ([a3711a1](https://github.com/ChristopherVR/ooxml/commit/a3711a14049a2cd11b994e76eafec039db2dc2de))

## [0.1.19](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.19) - 2026-10-22

### Documentation

- Describe the in-repo viewers, mcp packages and real release flow ([35e0055](https://github.com/ChristopherVR/ooxml/commit/35e0055d016b0a8dd9ac0255fb5dcae98e4568c1))

## [0.1.18](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.18) - 2026-10-22

### Bug Fixes

- **deps:** Track the core and viewer mcp packages through the workspace ([4373ea1](https://github.com/ChristopherVR/ooxml/commit/4373ea1209961b54e812b537711ac789ebdafb24))

## [0.1.17](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.17) - 2026-10-21

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([02d7713](https://github.com/ChristopherVR/ooxml/commit/02d7713ceb301921fd6c7706138d64f4c10a9d38))
- **deps:** Merge dependency update ([#94](https://github.com/ChristopherVR/ooxml/issues/94)) ([e381aba](https://github.com/ChristopherVR/ooxml/commit/e381abaa11794233cdea85a9dc6d098d4cb338c4))

## [0.1.16](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.16) - 2026-10-15

### Bug Fixes

- **automation:** Add --root and untrusted-content notes to MCP servers ([0da2c7d](https://github.com/ChristopherVR/ooxml/commit/0da2c7d1ec6655f649a608ded1f36dea41f34962))

### Styling

- Fix the oxfmt and oxlint warnings outside the viewers ([cd66f32](https://github.com/ChristopherVR/ooxml/commit/cd66f324a45aad1d9bd680d1518f0dbe8baf8fc3))

## [0.1.15](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.15) - 2026-10-09

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

## [0.1.14](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.14) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.1.13](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.13) - 2026-10-08

## [0.1.12](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.12) - 2026-10-08

## [0.1.11](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.11) - 2026-10-07

## [0.1.10](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.10) - 2026-10-07

## [0.1.9](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.9) - 2026-10-07

## [0.1.8](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.8) - 2026-10-06

### Refactor

- Make src/core a workspace package ([d5b8d6d](https://github.com/ChristopherVR/ooxml/commit/d5b8d6d91161c9ae631aa84436afb35105ac9517))

### Build & CI

- **deps:** Declare zod for the mcp contract test ([0fef32c](https://github.com/ChristopherVR/ooxml/commit/0fef32ca39d40bd4c1d032964eef6341df247ab4))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.1.7](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.7) - 2026-10-05

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

## [0.1.6](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.6) - 2026-10-05

## [0.1.5](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.5) - 2026-10-05

## [0.1.4](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.4) - 2026-10-05

## [0.1.3](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.3) - 2026-10-05

### Bug Fixes

- **mcp:** Restore the file: link to the core (by @ChristopherVR) ([f0354eb](https://github.com/ChristopherVR/ooxml/commit/f0354eb185f46f2b35fefe7796352c643d11fe88))
- **mcp:** State that ooxml-core is a dependency (by @ChristopherVR) ([7c5475c](https://github.com/ChristopherVR/ooxml/commit/7c5475cdf1533ea6436e8bd6becbdb6d0a2d88b1))

## [0.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.2) - 2026-10-05

### Build & CI

- **deps:** Update the MCP SDK and the CI toolchain pins (by @ChristopherVR) ([67aedfe](https://github.com/ChristopherVR/ooxml/commit/67aedfe96f6ce6acacf62b5d27196b7b9d3247e4))

## [0.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-mcp@0.1.1) - 2026-10-03

### Features

- **automation:** Centralize document tools and compose MCP servers (by @ChristopherVR) ([8c586b5](https://github.com/ChristopherVR/ooxml/commit/8c586b56df6b5853cfb9dd971073a04687d985b0))

### Documentation

- **packages:** Align npm readmes with powerpoint structure (by @ChristopherVR) ([52d33d4](https://github.com/ChristopherVR/ooxml/commit/52d33d4c58f4cfac7e47df3c890168ca8542752b))

