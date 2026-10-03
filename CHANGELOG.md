# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs releases whenever any published file changes, not only on
conventional commits.

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
