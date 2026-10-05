<div align="center">

# ooxml

**One repository for the Office Open XML libraries and the Word, Excel, Visio and OpenTeams viewers built on them.**

[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](../LICENSE)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](../CODE_OF_CONDUCT.md)

[**Try the apps**](https://christophervr.github.io/ooxml/) &nbsp;&middot;&nbsp;
[**What is here**](#what-is-in-this-repository) &nbsp;&middot;&nbsp;
[**Development**](#development) &nbsp;&middot;&nbsp;
[**Roadmap**](../docs/roadmap.md) &nbsp;&middot;&nbsp;
[**Contributing**](../CONTRIBUTING.md)

</div>

This repository holds [`ooxml-core`](https://www.npmjs.com/package/ooxml-core), a TypeScript library that reads, edits, validates and writes Office Open XML documents, together with the shared UI elements ([`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui)), the MCP servers, the Word, Excel, Visio and OpenTeams viewers, and the site that launches their demos. The library has its own README, which is also its npm page: [README.md](../README.md).

## What is in this repository

The logic of every Office product lives in this repository, and so do the interfaces for Word, Excel, Visio and OpenTeams. They are all members of one Bun workspace with one lockfile. Each folder below links to its own README.

| Folder                                     | Published as                                             | What it is                                                                                                                                                                                                                    |
| ------------------------------------------ | -------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`src/core`](../README.md)                     | [`ooxml-core`](https://www.npmjs.com/package/ooxml-core) | The library: every format's model, parser, writer, editing commands, layout and the shared areas under them. No UI. Its README is the package's npm page.                                                                     |
| [`src/ui`](../src/ui#readme)               | [`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui)     | The shared browser elements (Lit web components): ribbons, dialogs, menus, the title and status bars, and the Word, Excel, Visio and OpenTeams editors as `ooxml-ui/<product>` subpaths. DOM only; the core never imports it. |
| [`mcp/`](../mcp#readme)                    | [`ooxml-mcp`](https://www.npmjs.com/package/ooxml-mcp)   | The combined MCP server that lets an AI assistant work with Office documents.                                                                                                                                                 |
| [`viewers/docx`](../viewers/docx#readme)   | Word packages (below)                                    | The Word editor (`<docx-editor>`) and its bindings for React, Vue, Angular, Svelte, Solid and vanilla JavaScript. [Docs and demos](https://christophervr.github.io/ooxml/docx/).                                              |
| [`viewers/xlsx`](../viewers/xlsx#readme)   | Excel packages (below)                                   | The Excel editor (`<xlsx-editor>`) and its bindings. [Docs and demos](https://christophervr.github.io/ooxml/xlsx/).                                                                                                           |
| [`viewers/visio`](../viewers/visio#readme) | Visio packages (below)                                   | The local-first Visio viewer and its bindings. [Docs and demos](https://christophervr.github.io/ooxml/visio/).                                                                                                                |
| [`viewers/teams`](../viewers/teams#readme) | OpenTeams packages (below)                               | A bring-your-own-server team workspace (channels, chat, calls, shared Office files), its bindings and a reference server. [Docs and demos](https://christophervr.github.io/ooxml/teams/).                                     |
| [`site/`](../site)                         | not published                                            | The launcher at https://christophervr.github.io/ooxml/. Each viewer's documentation and demos are built into the same site, under `/docx/`, `/xlsx/`, `/visio/` and `/teams/`.                                                |

### Every published package

- **Word**: [`docx-viewer-mcp`](../viewers/docx/mcp#readme), [`docx-core`](../viewers/docx/packages/core#readme), [`docx-react-viewer`](../viewers/docx/packages/react#readme), [`docx-vue-viewer`](../viewers/docx/packages/vue#readme), [`docx-angular-viewer`](../viewers/docx/packages/angular#readme), [`docx-svelte-viewer`](../viewers/docx/packages/svelte#readme), [`docx-solid-viewer`](../viewers/docx/packages/solid#readme), [`docx-vanilla-viewer`](../viewers/docx/packages/vanilla#readme)
- **Excel**: [`xlsx-viewer-mcp`](../viewers/xlsx/mcp#readme), [`@christophervr/xlsx-core`](../viewers/xlsx/packages/core#readme), [`@christophervr/xlsx-react-viewer`](../viewers/xlsx/packages/react#readme), [`xlsx-vue-viewer`](../viewers/xlsx/packages/vue#readme), [`xlsx-angular-viewer`](../viewers/xlsx/packages/angular#readme), [`xlsx-svelte-viewer`](../viewers/xlsx/packages/svelte#readme), [`xlsx-solid-viewer`](../viewers/xlsx/packages/solid#readme), [`xlsx-vanilla-viewer`](../viewers/xlsx/packages/vanilla#readme)
- **Visio**: [`visio-viewer-mcp`](../viewers/visio/mcp#readme), [`visio-core`](../viewers/visio/packages/core#readme), [`visio-react-viewer`](../viewers/visio/packages/react#readme), [`visio-vue-viewer`](../viewers/visio/packages/vue#readme), [`visio-angular-viewer`](../viewers/visio/packages/angular#readme), [`visio-svelte-viewer`](../viewers/visio/packages/svelte#readme), [`visio-solid-viewer`](../viewers/visio/packages/solid#readme), [`visio-vanilla-viewer`](../viewers/visio/packages/vanilla#readme)
- **OpenTeams**: [`openteams-react-viewer`](../viewers/teams/packages/react#readme), [`openteams-vue-viewer`](../viewers/teams/packages/vue#readme), [`openteams-angular-viewer`](../viewers/teams/packages/angular#readme), [`openteams-svelte-viewer`](../viewers/teams/packages/svelte#readme), [`openteams-solid-viewer`](../viewers/teams/packages/solid#readme), [`openteams-vanilla-viewer`](../viewers/teams/packages/vanilla#readme), [`openteams-server`](../viewers/teams/server#readme)
- **Library and shared**: [`ooxml-core`](../README.md), [`ooxml-ui`](../src/ui#readme), [`ooxml-mcp`](../mcp#readme)

The viewers keep only what is specific to a framework: the bindings, demos, end-to-end tests and docs. Format logic stays in `src/core/`, and a viewer must not copy it.

### pptx-viewer

The PowerPoint viewer, [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer), is still its own repository for now, and it consumes `ooxml-core/pptx` and `ooxml-ui` from here. **The plan is to merge it into this repository as `viewers/pptx`**, the way the Word, Excel, Visio and OpenTeams viewers were merged (with their history and release tags), so that one change can span the library and every viewer. Until then a scheduled workflow keeps it on the latest releases. This is intended but not scheduled; see the [roadmap](../docs/roadmap.md).

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
