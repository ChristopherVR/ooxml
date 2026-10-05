# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs releases whenever any published file changes, not only on
conventional commits.

## [0.21.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.21.0) - 2026-10-05

### Features

- **chart:** Add the automatic value-axis scale (by @ChristopherVR) ([007f5b7](https://github.com/ChristopherVR/ooxml/commit/007f5b77507f99b8caf837fe0c82ea12d10465d6))
- **chart:** Add histogram and category binning (by @ChristopherVR) ([f3d564c](https://github.com/ChristopherVR/ooxml/commit/f3d564c705dc2f2ab6da4df4dcf590bd0c2df023))
- **pptx:** Add the pptx/ui subpath with the theme colour palette (by @ChristopherVR) ([0b7cde2](https://github.com/ChristopherVR/ooxml/commit/0b7cde28ecacc39bb80fbbebe1e6e9cb84e24634))
- **pptx:** Write chart spacing, axis visibility and area format (by @sedrew) ([5db45d5](https://github.com/ChristopherVR/ooxml/commit/5db45d573f28986acb8116f615ac0807a11afc91))
- **pptx:** Add chart spacing and area options to the chart SDK (by @sedrew) ([0e3e4ed](https://github.com/ChristopherVR/ooxml/commit/0e3e4ed0689e1ddfde5fdf1d0946ea60fb4bfe34))

### Bug Fixes

- **pptx:** Read series line style, keep table cell geometry fractional ([#6](https://github.com/ChristopherVR/ooxml/issues/6)) (by @IHAGI-c) ([f731f59](https://github.com/ChristopherVR/ooxml/commit/f731f59d5fc48c13e2a3032bc1566f10bae917af))
- **pptx:** Write c:scaling children in schema order (by @sedrew) ([a530d5b](https://github.com/ChristopherVR/ooxml/commit/a530d5b3f37badd4e6b12a6e7b387c23f28d84bb))

### Testing

- **pptx:** Cover chart spacing, axis visibility and area format (by @sedrew) ([5b912b7](https://github.com/ChristopherVR/ooxml/commit/5b912b7da0db97c58d56e3d9403a93a52e9408a7))

### Build & CI

- **viewers:** Make the viewers root workspaces with one lockfile (by @ChristopherVR) ([129c615](https://github.com/ChristopherVR/ooxml/commit/129c615fcf49bb15e25b136019be5425f6c36bd1))

### Chores

- **viewers:** Resolve imported viewers against this checkout (by @ChristopherVR) ([fc79891](https://github.com/ChristopherVR/ooxml/commit/fc798917be9668674973df731e56263ab063af4a))

## [0.20.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.20.0) - 2026-10-05

### Features

- **docx:** Move section edits, page setup, unit attrs and case into core (by @ChristopherVR) ([5e668b7](https://github.com/ChristopherVR/ooxml/commit/5e668b767189c06a2d739cd5809099643939ef9e))
- **xlsx:** Move the DOM-free editor logic into core (by @ChristopherVR) ([d6c6be1](https://github.com/ChristopherVR/ooxml/commit/d6c6be13ec5a1aa5826ca810a8a98b020d5c0878))
- **docx:** Move the DOM-free editor modules into core (by @ChristopherVR) ([e3bd270](https://github.com/ChristopherVR/ooxml/commit/e3bd2705501d0d30c715878ff2da87ff407fb961))
- **visio:** Move the DOM-free viewer modules into core (by @ChristopherVR) ([956cb90](https://github.com/ChristopherVR/ooxml/commit/956cb90371477a2e97e9036caf6f7b562d2a598b))

### Testing

- **xlsx:** Give the row-insert timing test headroom on shared runners (by @ChristopherVR) ([9c81002](https://github.com/ChristopherVR/ooxml/commit/9c81002188098eb8ad6230da1f80f339ea1fb8f1))

### Chores

- Lint with oxlint and pin oxfmt (by @ChristopherVR) ([e6e9aad](https://github.com/ChristopherVR/ooxml/commit/e6e9aade4bd74b3587acb80bd341b3e9505caf0b))

## [0.19.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.19.0) - 2026-10-05

### Features

- **docx:** Move column, line spacing and section helpers into core (by @ChristopherVR) ([b3beeb2](https://github.com/ChristopherVR/ooxml/commit/b3beeb2577c4cc6bcefe664f8e59499b124a8651))

## [0.18.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.18.0) - 2026-10-05

### Features

- **xlsx:** Move the formula editor text helpers into core (by @ChristopherVR) ([07e1108](https://github.com/ChristopherVR/ooxml/commit/07e110868a4287f529a87bf1e5147122f76d88df))

### Build & CI

- Add a script that reports dependency state across every repository (by @ChristopherVR) ([4b604ba](https://github.com/ChristopherVR/ooxml/commit/4b604ba58a472444eca8f77b3f0667c8a8f293b7))

## [0.17.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.17.0) - 2026-10-05

### Features

- **xlsx:** Add bond, coupon, discount-security and Treasury bill functions (by @ChristopherVR) ([761673c](https://github.com/ChristopherVR/ooxml/commit/761673c65671f8e802c69289160ace8b95bc76cb))
- **xlsx:** Add VDB, ENCODEURL, REGEXTEST, REGEXEXTRACT and REGEXREPLACE (by @ChristopherVR) ([c2c3eac](https://github.com/ChristopherVR/ooxml/commit/c2c3eacd6bd7f67a2dfa2657d5bb81b6c1a866fa))
- **xlsx:** Support iterative calculation for circular references (by @ChristopherVR) ([25c9307](https://github.com/ChristopherVR/ooxml/commit/25c930726dbbcbc8d692662347be89e786eb8452))

## [0.16.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.16.0) - 2026-10-04

### Features

- **xlsx:** Add F, beta and gamma inverses and the legacy distribution names (by @ChristopherVR) ([7218339](https://github.com/ChristopherVR/ooxml/commit/721833979bb5b9c939e08a29a1694251bb212e70))
- **xlsx:** Add Z.TEST, T.TEST, F.TEST, CHISQ.TEST and PROB (by @ChristopherVR) ([f86c3d9](https://github.com/ChristopherVR/ooxml/commit/f86c3d9cee3bf6d99c99e9d294ee19147fb9255c))
- **xlsx:** Add LINEST, LOGEST, TREND and GROWTH (by @ChristopherVR) ([d61c910](https://github.com/ChristopherVR/ooxml/commit/d61c9105d7b2ffd14426ae1e4f1a18d5da7d51f8))
- **xlsx:** Add the database functions (DSUM, DCOUNT, DGET and the rest) (by @ChristopherVR) ([8062b9d](https://github.com/ChristopherVR/ooxml/commit/8062b9d38729d74ae9e10ea279963e5d21fa03f4))

## [0.15.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.15.0) - 2026-10-04

### Features

- **teams:** Add the team-workspace logic area (by @ChristopherVR) ([c3cca75](https://github.com/ChristopherVR/ooxml/commit/c3cca752c7833b1e1f318c0c67287c5d0bff3a9e))

### Bug Fixes

- **teams:** Open files through short-lived signed links, not the token (by @ChristopherVR) ([5dcbc38](https://github.com/ChristopherVR/ooxml/commit/5dcbc38076192690018d1b37cc17fc042dfc4b6c))

### Dependencies

- **deps:** Update all dependencies (by @ChristopherVR) ([d13b03a](https://github.com/ChristopherVR/ooxml/commit/d13b03ad9cb10c74655d699e7799d11aca4b30a0))

## [0.14.1](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.14.1) - 2026-10-03

### Bug Fixes

- **collab:** Accept broadcast bytes cloned from another realm (by @ChristopherVR) ([03eb043](https://github.com/ChristopherVR/ooxml/commit/03eb0431c304641ba4e3f4e235625bc54c545105))

## [0.14.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.14.0) - 2026-10-03

### Features

- **collab:** Add broadcast channel transport and package adapter (by @ChristopherVR) ([9aa3f4a](https://github.com/ChristopherVR/ooxml/commit/9aa3f4ab8bbca06ede53c779e72342f3c8cb51ef))

## [0.13.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.13.0) - 2026-10-03

### Features

- **visio:** Import legacy VSD using the shared binary model (by @ChristopherVR) ([9810823](https://github.com/ChristopherVR/ooxml/commit/9810823ded5e2b1d6268914b7c9bd55a5ba62d0a))

## [0.12.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.12.0) - 2026-10-03

### Features

- **visio:** Admit proven inherited master shape moves (by @ChristopherVR) ([6cb176f](https://github.com/ChristopherVR/ooxml/commit/6cb176f4b8c58cb21b7aed30259e1dfc94371fb9))

### Performance

- **xlsx:** Reuse spill cells and index short row footprints (by @ChristopherVR) ([c3b7d37](https://github.com/ChristopherVR/ooxml/commit/c3b7d37c3cf2596a381b3b3d6ebce7a56a1e9aec))
- **xlsx:** Reuse validated full recalculation ordering (by @ChristopherVR) ([490f124](https://github.com/ChristopherVR/ooxml/commit/490f124e215853515c34d3f2504a700a62a97ee1))

## [0.11.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.11.0) - 2026-10-03

### Features

- **visio:** Move independently proven master instances (by @ChristopherVR) ([d6fc328](https://github.com/ChristopherVR/ooxml/commit/d6fc32885af6a7afe704a22337057a12c215dd3e))
- **automation:** Centralize document tools and compose MCP servers (by @ChristopherVR) ([8c586b5](https://github.com/ChristopherVR/ooxml/commit/8c586b56df6b5853cfb9dd971073a04687d985b0))

### Chores

- **style:** Format existing docs and package configuration (by @ChristopherVR) ([09a7466](https://github.com/ChristopherVR/ooxml/commit/09a7466e8289e8f1cd0165e7830138f73b93fcf2))

## [0.10.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.10.0) - 2026-10-03

### Features

- **visio:** Prove independent edits around active masters (by @ChristopherVR) ([5590c8a](https://github.com/ChristopherVR/ooxml/commit/5590c8af893e81b8a3d9f3f8874c9434821b6abd))
- **visio:** Evaluate bounded numeric formulas (by @ChristopherVR) ([2343de2](https://github.com/ChristopherVR/ooxml/commit/2343de2d82d863b7d9f765c07b8484b97a96917a))
- **crypto:** Share OOXML encryption, signatures, links and properties (by @ChristopherVR) ([050df71](https://github.com/ChristopherVR/ooxml/commit/050df71849a7ceae5bef03334c03a5bd4aeadc38))

### Bug Fixes

- **xlsx:** Fix 50+ Excel mismatches found in a full core review (by @ChristopherVR) ([4b4c7a0](https://github.com/ChristopherVR/ooxml/commit/4b4c7a006b22fe915ab9e682c37b20621c4869db))

## [0.9.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.9.0) - 2026-10-03

### Features

- Share chart text and svg geometry algorithms (by @ChristopherVR) ([bc2a7cd](https://github.com/ChristopherVR/ooxml/commit/bc2a7cdf2e5c77a2c7b1d92c63ba2b5188a1cbee))

## [0.8.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.8.0) - 2026-10-03

### Features

- **docx:** Preserve edits to cell vertical alignment (by @ChristopherVR) ([939eddf](https://github.com/ChristopherVR/ooxml/commit/939eddf3eb08a9333c399a1e46ec9c2f39150cdd))
- **docx:** Map layout fragments to source text ranges (by @ChristopherVR) ([55cbd6b](https://github.com/ChristopherVR/ooxml/commit/55cbd6bff039de1d383ef9434d3d54051e90357c))

## [0.7.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.7.0) - 2026-10-03

### Features

- **visio:** Add atomic geometry editing and bounded recalculation (by @ChristopherVR) ([dc60e01](https://github.com/ChristopherVR/ooxml/commit/dc60e01ff97ae4db5e346c9942cd4e9052ba7b9d))

## [0.6.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.6.0) - 2026-10-03

### Features

- **digest:** Add a shared area for synchronous digests (by @ChristopherVR) ([5094d3d](https://github.com/ChristopherVR/ooxml/commit/5094d3dd74f4313c7312867aec20ef25a4a344a4))
- **docx:** Preserve edits to individual table cell margins (by @ChristopherVR) ([c413404](https://github.com/ChristopherVR/ooxml/commit/c41340440621e94628cd26728b5947d493b2e4e7))
- **docx:** Preserve imported equations as display-only runs (by @ChristopherVR) ([7114aed](https://github.com/ChristopherVR/ooxml/commit/7114aed5e79a051dba9060129c3938b420144497))

### Bug Fixes

- **xlsx:** Verify Excel's modern sheet password hash (by @ChristopherVR) ([b38a0b7](https://github.com/ChristopherVR/ooxml/commit/b38a0b79b65f5bd00484fe9060fe1d065bb6e5ee))
- **xlsx:** Keep removed passwords removed and check every hash (by @ChristopherVR) ([5ef0c50](https://github.com/ChristopherVR/ooxml/commit/5ef0c503f1c6b48bd3ac15bb805a587a7e485d6a))

### Testing

- **pptx:** Preserve smartart across repeated save and reload (by @ChristopherVR) ([0a94952](https://github.com/ChristopherVR/ooxml/commit/0a949525b65b879dd10bad2641658e68d4c48036))

## [0.5.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.5.0) - 2026-10-03

### Features

- **math:** Share equation converters across office formats (by @ChristopherVR) ([2763d2a](https://github.com/ChristopherVR/ooxml/commit/2763d2afe28f0e3d7dead91d25d2011d2f618dc1))

### Bug Fixes

- **docx:** Preserve edited shading and automatic table widths (by @ChristopherVR) ([1d15587](https://github.com/ChristopherVR/ooxml/commit/1d15587d642872bda18c468d6dba13a1b2c39c33))

## [0.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.4.0) - 2026-10-02

### Features

- **visio:** Round bounded open orthogonal connectors (by @codex) ([2c6e66a](https://github.com/ChristopherVR/ooxml/commit/2c6e66afa7520e882d89498bebd117eb966476ea))
- **opc:** Add SpreadsheetML namespaces and relationship types (by @ChristopherVR) ([adc12c1](https://github.com/ChristopherVR/ooxml/commit/adc12c17c0255f56443cd5aea45e307eb6e68ae1))
- **xlsx:** Add the SpreadsheetML area (by @ChristopherVR) ([981e75b](https://github.com/ChristopherVR/ooxml/commit/981e75bf9f84dbedb8aafa99fea236c580cba9c2))

### Testing

- **xlsx:** Restore the Office user name after generating fixtures (by @ChristopherVR) ([2d9d2f9](https://github.com/ChristopherVR/ooxml/commit/2d9d2f950e3b43bad6811212850f0ecf2f0eb92b))

## [0.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/ooxml-core@0.3.0) - 2026-10-02

### Features

- **visio:** Add bounded read-only VSDX scene core (by @codex) ([dee6d0d](https://github.com/ChristopherVR/ooxml/commit/dee6d0dd475d951e6fdb59994cc217674d496cd0))
- **visio:** Add bounded saves, fills and worker vector import (by @codex) ([7364222](https://github.com/ChristopherVR/ooxml/commit/7364222cc9687da5e10a5992c2bf373cb6d68d98))

## [0.2.0](https://github.com/ChristopherVR/ooxml-core/releases/tag/@christophervr/ooxml-core@0.2.0) - 2026-10-01

### Features

- **diagram:** Add a format-neutral SmartArt area (by @ChristopherVR) ([d5fc851](https://github.com/ChristopherVR/ooxml-core/commit/d5fc85134c512afc6109316b719659937c8c90da))
- **docx:** Read SmartArt drawings and preserve them on save (by @ChristopherVR) ([ab0e611](https://github.com/ChristopherVR/ooxml-core/commit/ab0e6113475918090d16249d3cbb850403d61339))
- **docx:** Add layout and load areas moved from docx-viewer (by @ChristopherVR) ([7350cc1](https://github.com/ChristopherVR/ooxml-core/commit/7350cc1ae92f5b682ebe809d173c3bd6252ea847))
- **collab:** Add the format-neutral collaboration area (by @ChristopherVR) ([f4cb051](https://github.com/ChristopherVR/ooxml-core/commit/f4cb0519a3d61cce7b6e0b829fe7ee6c064a6fd2))

### Refactor

- **pptx:** Re-import SmartArt parsers from the diagram area (by @ChristopherVR) ([f7a6f3a](https://github.com/ChristopherVR/ooxml-core/commit/f7a6f3acf856f42294ca3a9d1e051a79e22e6a0b))

### Testing

- **pptx:** Commit the e2e deck snapshot and stop reading from pptx-viewer (by @ChristopherVR) ([c3e0979](https://github.com/ChristopherVR/ooxml-core/commit/c3e09792c624155c9de416e844ba01b257218234))

### Build & CI

- Run the build steps in parallel and stop prepack rebuilding (by @ChristopherVR) ([4d25238](https://github.com/ChristopherVR/ooxml-core/commit/4d25238854ddf6169fbe4f72e535e6b78b27dfec))
- Automate releases from conventional commits with OIDC publishing (by @ChristopherVR) ([09d5a8b](https://github.com/ChristopherVR/ooxml-core/commit/09d5a8b7a09d045d525b1ce5b19b28844fcd2f6b))

### Chores

- **ui:** Scaffold the @christophervr/office-ui workspace package (by @ChristopherVR) ([6bb6ddd](https://github.com/ChristopherVR/ooxml-core/commit/6bb6ddd0e1bd3594996810639232070239b225a4))

## [0.1.0](https://github.com/ChristopherVR/ooxml-core/releases/tag/@christophervr/ooxml-core@0.1.0) - 2026-10-01

### Features

- **opc:** Add @christophervr/ooxml-opc (relationships, content types, zip helpers) (by @ChristopherVR) ([08e5427](https://github.com/ChristopherVR/ooxml-core/commit/08e5427cd24ab1ae03ba05fd0faef76a65ddeda1))
- **docx:** Move the Word core (model, parser, serializer, editing) into the docx area (by @ChristopherVR) ([b4f83e1](https://github.com/ChristopherVR/ooxml-core/commit/b4f83e123d492414258261095e72c8113459a630))
- **pptx:** Move the PowerPoint core from pptx-viewer into src/pptx (by @ChristopherVR) ([4638a2a](https://github.com/ChristopherVR/ooxml-core/commit/4638a2a991881d1e0b7fb9d89988bfd284911648))
- **pptx:** Write gradient fills for chart series and data points (by @ChristopherVR) ([3d80333](https://github.com/ChristopherVR/ooxml-core/commit/3d80333e33ed54240c488102b665fc3ef45a8377))

### Refactor

- One published package, @christophervr/ooxml-core, with an area per subpath (by @ChristopherVR) ([c8d5ef1](https://github.com/ChristopherVR/ooxml-core/commit/c8d5ef1acd2f7d545f8b5178a7101204064a0850))
- Treat docx and pptx symmetrically; publish via OIDC (by @ChristopherVR) ([1272e23](https://github.com/ChristopherVR/ooxml-core/commit/1272e2331f542bd01050303492fa9423aef359ee))

### Testing

- **pptx:** Declare group-drill.pptx in the fixture manifest; sync missing e2e decks every run (by @ChristopherVR) ([7a1b245](https://github.com/ChristopherVR/ooxml-core/commit/7a1b245f9d9bb342bcf6ea296fe3887da13c139c))

### Build & CI

- Compile ooxml-units to dist and export it for consumers (by @ChristopherVR) ([f7476df](https://github.com/ChristopherVR/ooxml-core/commit/f7476df1dc88fdaac213a44a2616afaf4fad4314))
- **pptx:** Subpath exports, relaxed tsconfig project and dual-format bundle for the pptx area (by @ChristopherVR) ([359be0b](https://github.com/ChristopherVR/ooxml-core/commit/359be0b60bb5f9b82df5d1a89fcb20d90dd4ad76))

### Chores

- Add CI, release workflow, package smoke test and community files (by @ChristopherVR) ([f1d4805](https://github.com/ChristopherVR/ooxml-core/commit/f1d48053fb86244b129de8d2bf96cc77ecf7f0eb))
