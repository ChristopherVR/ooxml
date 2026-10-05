# Releasing packages

Seven unscoped packages are published to npm under the product name **OpenTeams** (`openteams`
itself belongs to an unrelated project, so every name is longer). Each has its **own version
line**, bumped only when it actually changes, and its own git tag `<npm-name>@<version>` (for
example `openteams-server@0.2.0`). The flow, the scripts and the workflows are ported from
[docx-viewer](https://github.com/ChristopherVR/docx-viewer) (itself adapted from pptx-viewer); each
ported file says so in its header, and the deviations are listed at the end.

| Package                    | Dir                | Depends on (registry)                               | Bundled in               |
| -------------------------- | ------------------ | --------------------------------------------------- | ------------------------ |
| `openteams-react-viewer`   | `packages/react`   | `ooxml-core`, `ooxml-ui`, `lit`; peer `react`       | `packages/web-component` |
| `openteams-vue-viewer`     | `packages/vue`     | `ooxml-core`, `ooxml-ui`, `lit`; peer `vue`         | `packages/web-component` |
| `openteams-angular-viewer` | `packages/angular` | `ooxml-core`, `ooxml-ui`, `lit`; peer `@angular/core` | `packages/web-component` |
| `openteams-svelte-viewer`  | `packages/svelte`  | `ooxml-core`, `ooxml-ui`, `lit`; peer `svelte`      | `packages/web-component` |
| `openteams-solid-viewer`   | `packages/solid`   | `ooxml-core`, `ooxml-ui`, `lit`; peer `solid-js`    | `packages/web-component` |
| `openteams-vanilla-viewer` | `packages/vanilla` | `ooxml-core`, `ooxml-ui`, `lit`                     | `packages/web-component` |
| `openteams-server`         | `server`           | `ws`, `yjs`, `y-protocols`, `lib0`                  | nothing (ships its source) |

`teams-viewer` (`packages/web-component`, the `<teams-app>` element) and the demos are `private`
and never published. `scripts/build-packages.mjs` inlines the web component, CSS included, into
every framework bundle (vite library build) and flattens its declarations into each package's
`dist/index.d.ts` (rollup-plugin-dts), so a tarball imports only what its manifest declares.
Svelte ships `dist/Teams.svelte` as source next to a bundled `dist/runtime.js` (the store and the
helpers, under `openteams-svelte-viewer/runtime`); Angular is a bundled ES module with its
decorators applied at runtime (it needs the JIT compiler in the consuming app). The server is plain
ESM and is not built; its `openteams-server` command is `server/bin.mjs`.

The team-workspace logic is the published `ooxml-core` (`ooxml-core/teams`) and the visual
primitives the published `ooxml-ui`, both caret ranges (`^x.y.z`) in every framework manifest so
the scheduled `sync-ooxml.yml` workflow can move them. A range change releases that package only.

## How a release is decided

`scripts/release-plan.mjs` computes everything from git history; run it at any time (it writes the
git-ignored `release-plan.json` and changes nothing else):

```sh
bun run release:plan            # or: node scripts/release-plan.mjs --no-npm   (offline)
```

For each package it finds the newest `<npm-name>@x.y.z` tag that is an ancestor of `HEAD` and
looks at the files changed since:

- **Own files.** Any non-test file under the package directory releases it. Tests, `CHANGELOG.md`
  and manifest edits that only touch `version`, `scripts` or `devDependencies` never do.
- **Bundled internal package (trigger).** A non-test change under `packages/web-component` releases
  every framework package.
- **Shared build pipeline.** `scripts/build-packages.mjs` and `tsconfig.release.json` release every
  framework package. The server opts out (`global: false`): only `server/**` releases it.
- **Never published.** A package with no tag that is not on npm releases at its manifest version
  (`0.1.0` for all seven today).

The bump level is the highest Conventional Commit level among the commits that touch published
files in the package's scope: `!` / `BREAKING CHANGE:` is major, `feat` minor, anything else patch.
`--write` stamps the new versions (the workflow uses it).

## The release workflow

`.github/workflows/release.yml` runs hourly and on manual dispatch:

1. Check out `main` with `RELEASE_TOKEN`, run the planner with `--write`; no changed package ends
   the run.
2. Fail early unless the `NPM_PUBLISH` repository variable is `true`.
3. `bun install`, `bun run build:packages`, `bun run check:published`, `bun run pack:smoke`.
4. Prepend a section to each released package's `CHANGELOG.md` (git-cliff, `cliff.toml`) and a
   dated section to the root `CHANGELOG.md` (`cliff-root.toml`). Changelogs are prepend-only.
5. Commit `chore(release): bump versions and update changelogs [skip ci]` to `main`.
6. Create a GitHub release and tag per package (`scripts/release-notes.mjs`) and prune superseded
   releases (`scripts/prune-releases.mjs --keep 1`, also run daily by `prune-releases.yml`).
7. `publish` job (environment `npm`, `id-token: write`): build again and run
   `scripts/publish-released.mjs --plan release-plan.json`.

Re-publish one existing tag with `gh workflow run release.yml -f tag=openteams-server@0.1.0`.

`publish-released.mjs` publishes in dependency order with `npm publish --access public`, after
checking that the manifest on disk is the version being published, that no range uses
`workspace:`, `file:` or `link:`, that no dependency names a private package (`teams-viewer` or
any other `teams-*` name), and that the version is not already on npm (re-runs are safe). A version
older than `latest` goes out under the `old` dist-tag.

### Authentication and provenance

The `publish` job uses npm **trusted publishing** (GitHub OIDC): no npm token is stored anywhere.

**Provenance is conditional, because this repository is private for now.** npm can only attest a
build from a public repository. The `Decide on provenance` step sets `NPM_PROVENANCE`: the
`NPM_PROVENANCE` repository variable (`true` / `false`) wins when set; otherwise it is `true`
exactly when the repository is public (read through the GitHub API). `publish-released.mjs` adds
`--provenance` only when `NPM_PROVENANCE=true` **and** it runs in GitHub Actions with an OIDC token,
so it is never added on a laptop. Once the repository is public, provenance switches on by itself.

## Publishing from your machine

`bun run publish:local` publishes from a maintainer's machine, for example the very first version of
each package (npm only lets you add a trusted publisher to a package that already exists):

```sh
bun run publish:local -- --dry-run      # everything except the upload (npm publish --dry-run)
bun run publish:local                   # asks you to type "yes" before publishing
bun run publish:local -- --package server --yes
```

What it does, in order:

1. `npm whoami` against `https://registry.npmjs.org/`. If you are not logged in it runs
   `npm login --auth-type=web` in your terminal (npm opens the browser) and checks again. A dry run
   does not need a login.
2. Refuses a dirty work tree or a branch other than `main` unless `--allow-dirty`.
3. Computes the release plan and runs the CI checks from `publish-released.mjs` on each package
   (version on disk equals the planned version, no `file:` / `workspace:` ranges, nothing private),
   then `bun run build:packages` (skip with `--skip-build`), `check:published` and `pack:smoke`.
4. Prints the plan (each `package@version` and whether it is already on npm) and waits for `yes`
   (or `--yes`).
5. Publishes in dependency order with `npm publish --access public`, with stdin detached so npm can
   never prompt. It never passes `--provenance` locally.
6. Prints the `git tag` commands for what it published. Push those tags: without them the next
   scheduled run would see an untagged, already-published package and release it again as a patch.

**It never asks for or accepts a one-time password** (there is no `--otp`). If npm answers that the
publish needs a 2FA code (`EOTP`), it stops with a non-zero exit. Publish with a granular access
token that has "bypass 2FA" enabled (log in with it, or put it in your user `.npmrc` yourself), or
through the CI trusted-publishing path. The script itself never reads, writes or prints a token;
`npm login` stores its own credentials in your user `.npmrc` as usual.

## One-time setup (by hand)

1. **Ruleset on `main`**: require the `ci-success` check and the `Conventional Commits` check; add a
   bypass for repository admins (the release commit is pushed straight to `main`).
2. **`RELEASE_TOKEN` secret**: a fine-grained personal access token with `Contents: Read and write`
   on this repository, owned by a repository admin.
3. **`NPM_PUBLISH` repository variable** set to `true`; optionally `NPM_PROVENANCE`.
4. **`npm` environment** (Settings, Environments), optionally with required reviewers.
5. **First publish of each package name**, with `bun run publish:local`, then push the
   `<npm-name>@0.1.0` tags it prints.
6. **npm trusted publisher for each of the seven packages** (package settings on npmjs.com,
   "Trusted Publisher", GitHub Actions): user `ChristopherVR`, repository `teams-viewer`, workflow
   `release.yml`, environment `npm`.
7. Make the repository public when you want provenance attestations.

## Deviations from docx-viewer

- `--provenance` is conditional (above) instead of always on, because this repository is private.
- The planner table has a `global: false` flag so the server is not released by changes to the
  framework build pipeline.
- `publish-local.mjs` is new (docx-viewer bootstrapped its first versions by hand).
- `check-published-refs.mjs` also scans `.mjs` files and the server directory.
- `check-conventional-commits.test.mjs` checks history only after `Create bun.lock`, the one
  commit made before the convention was enforced.
- TypeScript is 7 here; `rollup-plugin-dts` reaches the TypeScript 6 API through
  `@typescript/typescript6`.
- There is no browser (Playwright) or MCP job in CI, and no docs site workflow.
