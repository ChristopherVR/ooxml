# Contributing to visio-viewer

Read [AGENTS.md](AGENTS.md) for the working agreements (all Visio logic in `ooxml-core/visio`, the
viewer element in `ooxml-ui/visio`, honest reporting of unsupported features, no em-dashes).

## Getting set up

You need Node.js 22 or later (24 in CI). This repository uses npm.

```sh
npm ci
npm run typecheck
npm test
npm run test:scripts         # release planner, publish guards, commit checks
npm run fmt:check
```

Formatting is `oxfmt` (`npm run fmt`): tabs, single quotes, width 100. Lint is `oxlint`.

## Commit conventions

Commits **must** follow [Conventional Commits](https://www.conventionalcommits.org). This is
load-bearing, not cosmetic: each published viewer package is versioned and released independently,
and **the bump level is derived from your commit type**. A mislabelled commit mis-versions a
package, and a non-conforming commit is silently dropped from the changelog. The `Conventional
Commits` check on every pull request validates the PR title and every commit subject
(`scripts/check-conventional-commits.mjs`).

```
<type>(<scope>): <subject>

<body>

<footer>
```

- **type**: `feat` (minor bump); `fix`, `perf`, `refactor`, `docs`, `test`, `build`, `ci`, `style`,
  `chore`, `revert` (patch bump). A `!` after the type/scope, or a `BREAKING CHANGE:` footer, bumps
  major.
- **scope**: the package or area: `react`, `vue`, `angular`, `svelte`, `solid`, `vanilla`,
  `bindings`, `demo`, `release`, `ci`, `deps`, `docs`.
- **subject**: imperative, lower-case, no trailing period, header at most 72 characters.

Which package a commit versions is decided by the **paths it touches**, not by its scope. Changes
that only touch tests, docs or the demos never release anything, whatever their type. Write
multi-line messages with `git commit -F <file>`, and do not edit versions or changelogs by hand.

Examples: `feat(bindings): expose the selection event`, `fix(vue): forward the zoom prop`.

## Pull requests

Keep them focused, describe how you checked the change, and add a regression test for behaviour
changes. Do not bump versions in a pull request. CI must be green.

## License

By contributing you agree that your contributions are licensed under the
[Apache License 2.0](LICENSE).
