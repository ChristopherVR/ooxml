# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs releases whenever any published file changes, not only on
conventional commits.

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
