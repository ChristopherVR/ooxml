# Contributing to ooxml-core

Thanks for wanting to help. Please read the [Code of Conduct](CODE_OF_CONDUCT.md) first. Found a security problem? Do **not** open a public issue: see [SECURITY.md](SECURITY.md).

## What belongs here

This repository is the single package `ooxml-core`: all the logic behind the Office viewers (`docx-viewer`, `pptx-viewer`, later `xlsx-viewer`), organised as **areas** under `src/<area>/`, each a subpath export (`ooxml-core/xml`). The viewers keep only their UI.

- Shared modern OOXML (units, colour, geometry, XML, OPC, and later DrawingML, charts, diagrams) goes in its own area.
- Format-specific code goes in `docx`, `pptx` or `xlsx`.
- Legacy binary formats (DOC, XLS, PPT, the compound-file container) belong in [`ole2`](https://github.com/ChristopherVR/ole2), not here. Modern OOXML never goes into `ole2`.
- UI (components, ribbons, dialogs, styling) belongs in the viewer repositories.
- Do not add a second package; add an area. Read [AGENTS.md](AGENTS.md) for the full working agreements.

## Getting set up

You need [Bun](https://bun.sh/) and Node.js 22 or newer.

```bash
git clone https://github.com/ChristopherVR/ooxml.git
cd ooxml
bun install
bun run typecheck   # strict project and the pptx project
bun run test
bun run build
bun run test:package   # packs the build and imports every entry point from a clean install
```

The pptx tests read real decks from `src/pptx/__tests__/fixtures`, including a committed snapshot of the pptx-viewer end-to-end decks under `fixtures/e2e`; everything needed is in this repository.

## Code rules

- TypeScript is **strict**, including `exactOptionalPropertyTypes` and `noUncheckedIndexedAccess`. The `src/pptx` area is compiled with relaxed flags while it is tightened; new code anywhere else must pass the strict project.
- Keep modules under about 300 lines. Tests live next to the code (`*.test.ts`).
- Add a regression test for every parsing, preservation, round-trip or editing change. Real-file fixtures beat hand-built XML.
- Unsupported features must be reported, never silently dropped, and nothing may claim Office parity or lossless export without evidence.
- When moving code in from another repository, add an entry to [PROVENANCE.md](PROVENANCE.md) (source, path, commit, what changed).
- Formatting is `oxfmt` (`bun run fmt`); the project uses tabs and single quotes.

## Commit conventions

Commits **must** follow [Conventional Commits](https://www.conventionalcommits.org). This is load-bearing, not cosmetic: the published package is versioned and released automatically, and **the bump level is derived from your commit type**. A mislabelled commit mis-versions the package, and a non-conforming commit is silently dropped from the changelog. The `Conventional Commits` check on every pull request validates the PR title and every commit subject (`scripts/check-conventional-commits.mjs`).

```
<type>(<scope>): <subject>

<body>

<footer>
```

- **type**: `feat` (minor bump); `fix`, `perf`, `refactor`, `docs`, `test`, `build`, `ci`, `style`, `chore`, `revert` (patch bump). A `!` after the type/scope, or a `BREAKING CHANGE:` footer, bumps major.
- **scope**: the area: `units`, `color`, `geometry`, `xml`, `opc`, `docx`, `pptx`, `ci`, `deps`, `docs`.
- **subject**: imperative, lower-case, no trailing period, header at most 72 characters.

Whether a commit releases at all is decided by the **paths it touches**, not its type: only changes to a package's published files release a new version of that package (for `ooxml-core`: `src/` outside tests, the bundler and declaration configs, the manifest's shipping fields, the licence files; for `ooxml-ui`: everything under `packages/ui/` outside tests). Tests, docs, CI and fixtures never do, whatever their type, and a test-only `feat` does not raise the bump level. See [docs/releasing.md](docs/releasing.md).

Examples: `feat(xml): add a streaming serializer`, `fix(opc): resolve relative part names`, `feat(docx)!: rename the section model`.

The release workflow writes `chore(release): bump versions and update changelogs [skip ci]` commits to `main`; do not edit the version or `CHANGELOG.md` by hand.

## Pull requests

- Keep pull requests focused; describe what changed and how you checked it. CI runs typecheck, tests, the release-script tests, build and the package smoke test, and must be green (`ci-success`).
- Releases are cut automatically from `main` (hourly, plus manual dispatch); do not bump versions in pull requests.

## License

By contributing you agree that your contributions are licensed under the [Apache License 2.0](LICENSE).
