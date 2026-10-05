<div align="center">

# ooxml

**One repository for the Office Open XML libraries and the Word, Excel, Visio and OpenTeams viewers built on them.**

[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](../LICENSE)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](../CODE_OF_CONDUCT.md)

[**Try the apps**](https://christophervr.github.io/ooxml/) &nbsp;&middot;&nbsp;
[**What is here**](#what-is-in-this-repository) &nbsp;&middot;&nbsp;
[**Development**](#development) &nbsp;&middot;&nbsp;
[**Roadmap**](#roadmap) &nbsp;&middot;&nbsp;
[**Contributing**](../CONTRIBUTING.md)

</div>

This repository holds [`ooxml-core`](https://www.npmjs.com/package/ooxml-core), a TypeScript library that reads, edits, validates and writes Office Open XML documents, together with the shared UI elements ([`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui)), the MCP servers, the Word, Excel, Visio and OpenTeams viewers, and the site that launches their demos. The library has its own README, which is also its npm page: [README.md](../README.md).

## What is in this repository

The logic of every Office product lives in this repository, and so do the interfaces for Word, Excel, Visio and OpenTeams. They are all members of one Bun workspace with one lockfile.

| Path            | Published as                                                             | What it is                                                                                                                                                                                                                    |
| --------------- | ------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src/`          | [`ooxml-core`](https://www.npmjs.com/package/ooxml-core)                 | The library: every format's model, parser, writer, editing commands, layout and the shared areas under them (see below). No UI.                                                                                               |
| `src/ui`   | [`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui)                     | The shared browser elements (Lit web components): ribbons, dialogs, menus, the title and status bars, and the Word, Excel, Visio and OpenTeams editors as `ooxml-ui/<product>` subpaths. DOM only; the core never imports it. |
| `mcp/`          | [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp)                   | The combined MCP server that lets an AI assistant work with Office documents.                                                                                                                                                 |
| `viewers/docx`  | `docx-core`, `docx-<framework>-viewer`, `docx-viewer-mcp`                | The Word editor and its bindings for React, Vue, Angular, Svelte, Solid and vanilla JavaScript.                                                                                                                               |
| `viewers/xlsx`  | `@christophervr/xlsx-core`, `xlsx-<framework>-viewer`, `xlsx-viewer-mcp` | The Excel editor and its bindings.                                                                                                                                                                                            |
| `viewers/visio` | `visio-core`, `visio-<framework>-viewer`, `visio-viewer-mcp`             | The local-first Visio viewer and its bindings.                                                                                                                                                                                |
| `viewers/teams` | `openteams-<framework>-viewer`, `openteams-server`                       | OpenTeams: a bring-your-own-server team workspace (channels, chat, calls, shared Office files), its bindings and a reference server.                                                                                          |
| `site/`         | not published                                                            | The launcher at https://christophervr.github.io/ooxml/. Each viewer's documentation and demos are built into the same site, under `/docx/`, `/xlsx/`, `/visio/` and `/teams/`.                                                |

The viewers keep only what is specific to a framework: the bindings, demos, end-to-end tests and docs. Format logic stays in `src/`, and a viewer must not copy it. The PowerPoint viewer, [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer), is still its own repository and consumes `ooxml-core/pptx` and `ooxml-ui` from here. The docx, xlsx, visio and teams viewers used to be separate repositories; they were merged here with their history and release tags.

## Roadmap

The shared areas come first, then the formats and products on top of them.

**Done**

- Shared areas: `xml`, `opc`, `units`, `color`, `geometry`, `diagram` (SmartArt), `digest`, `crypto` (package encryption), `math` and `collab` (Yjs real-time collaboration).
- Formats as areas: `docx`, `xlsx`, `pptx` and `visio`, each with its own parser and writer (Visio edits are bounded, and legacy `.vsd` is preview only).
- Products: the Word, Excel, Visio and OpenTeams viewers and the shared `ooxml-ui` elements now live here, released from one pipeline.

**Next**

- **DrawingML and charts as shared areas:** move the DrawingML (fills, lines, effects, text, theme) and chart code out of `pptx` so every format uses one implementation. The order is in `docs/agnostic-core-plan.md`.
- **One XML model:** the `pptx` area still uses its own XML object model and is compiled with relaxed TypeScript flags while it is migrated onto the shared `xml` area and tightened. New code is strict.
- **Less logic in the viewers:** the Word editor's view code and the PowerPoint viewer's `shared` render logic still live in the viewers and move into areas over time.

## Development

You need [Bun](https://bun.sh/) and Node.js 22 or newer. One install at the root sets up the library, the UI and every viewer.

```bash
bun install
bun run typecheck      # strict project and the pptx project
bun run test           # the library's unit tests
bun run build          # the library
bun run test:package   # packs the build and imports every entry point from a clean install
```

The viewers resolve `ooxml-ui` to the copy in this workspace through its built `dist`, so build the library and the UI before working on one:

```bash
bun run build && bun run --cwd src/ui build
cd viewers/docx        # or xlsx, visio, teams
bun run typecheck && bun run test
bun run demo           # docx, xlsx: the demo app (visio, teams: bun run dev)
```

Each viewer has its own scripts and browser tests; the root `typecheck`, `test`, `fmt` and `lint` do not cover `viewers/`. CI (`.github/workflows/ci.yml`) decides from the files a change touches what to run, including which viewers to verify.

Releases are automated. Every published package has its own version and tag (`<npm-name>@<version>`), planned from conventional commits, with changelogs and npm publishing through trusted publishing (OIDC, with provenance). See [docs/releasing.md](../docs/releasing.md). Commits must follow [Conventional Commits](../CONTRIBUTING.md), and nobody publishes or tags by hand.

The working agreements are in [AGENTS.md](../AGENTS.md), and [PROVENANCE.md](../PROVENANCE.md) records where each module came from.

## Documentation and related projects

- The viewers, in this repository under `viewers/` ([docx](https://christophervr.github.io/ooxml/docx/), [xlsx](https://christophervr.github.io/ooxml/xlsx/), [visio](https://christophervr.github.io/ooxml/visio/), [teams](https://christophervr.github.io/ooxml/teams/)) and [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer): the editors and viewers built on this package.
- [OOXML Office](https://christophervr.github.io/ooxml/): the suite's launcher page (`site/`). It opens the demos of every viewer: the ones in this repository are built into the same site, and the PowerPoint demos come from the pptx-viewer site.
- [ole2](https://github.com/ChristopherVR/ole2): the compound-file container and legacy binary Office codecs.
- [emf-converter](https://github.com/ChristopherVR/emf-converter) and [mtx-decompressor](https://github.com/ChristopherVR/mtx-decompressor): EMF/WMF rendering and embedded-font (MicroType Express) decompression, used by the library.

## License

[Apache-2.0](../LICENSE). Third-party notices are in [NOTICE](../NOTICE) and [THIRD-PARTY-LICENSES](../THIRD-PARTY-LICENSES).
