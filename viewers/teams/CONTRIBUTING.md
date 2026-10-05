# Contributing to OpenTeams (teams-viewer)

Read [AGENTS.md](AGENTS.md) for the working agreements (all logic in `ooxml-core/teams`, visual
primitives in `ooxml-ui`, honest reporting of unsupported features, no em-dashes).

## Getting set up

You need [Bun](https://bun.sh/) 1.3 or later and Node.js 22 or later (24 in CI).

```sh
bun install
bun run typecheck
bun run test                 # vitest (jsdom) and the server's node:test suite
bun run test:scripts         # release planner, publish guards, commit and changelog checks
bun run build:packages && bun run check:published && bun run pack:smoke
```

With the `ooxml` repository checked out next to this one (`../ooxml-core`), tests and demos use its
source; without it they use the published `ooxml-core` and `ooxml-ui`. Formatting is `oxfmt`
(`bun run fmt`, `bun run fmt:check`): tabs, single quotes, width 100.

## Commit conventions

Commits **must** follow [Conventional Commits](https://www.conventionalcommits.org). This is
load-bearing, not cosmetic: each of the seven published packages is versioned and released
independently, and **the bump level is derived from your commit type**. A mislabelled commit
mis-versions a published package, and a non-conforming commit is silently dropped from the
changelog. The `Conventional Commits` check on every pull request validates the PR title and every
commit subject (`scripts/check-conventional-commits.mjs`).

```
<type>(<scope>): <subject>

<body>

<footer>
```

- **type**: `feat` (minor bump); `fix`, `perf`, `refactor`, `docs`, `test`, `build`, `ci`, `style`,
  `chore`, `revert` (patch bump). A `!` after the type/scope, or a `BREAKING CHANGE:` footer, bumps
  major.
- **scope**: the package or area: `react`, `vue`, `angular`, `svelte`, `solid`, `vanilla`,
  `server`, the internal `web-component`, `demo`, `release`, `ci`, `deps`, `docs`.
- **subject**: imperative, lower-case, no trailing period, header at most 72 characters.

Which package a commit versions is decided by the **paths it touches**, not by its scope, so keep a
commit within one package where practical. A change to `packages/web-component` re-releases every
framework package, because it is bundled into each one; a change under `server/` releases only
`openteams-server` (see [docs/releasing.md](docs/releasing.md)). Changes that only touch tests,
docs or the demos never release anything, whatever their type.

Examples: `feat(web-component): add a channel topic editor`, `fix(vue): forward the openers
prop`, `feat(server)!: require a token by default`.

The release workflow also writes `chore(release): bump versions and update changelogs [skip ci]`
commits to `main`; do not edit versions or `CHANGELOG.md` files by hand.

## Pull requests

Keep them focused, describe how you checked the change, and add a regression test for behaviour
changes. Do not bump versions in a pull request. CI must be green (`ci-success`).

## License

By contributing you agree that your contributions are licensed under the
[Apache License 2.0](LICENSE).
