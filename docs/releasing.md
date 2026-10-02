# Releasing

This repository publishes two packages to npm, each with its own version line, tag, changelog and GitHub release:

| Key    | npm name                    | Directory     | Changelog                  | Depends on |
| ------ | --------------------------- | ------------- | -------------------------- | ---------- |
| `core` | `ooxml-core` | `.` (root)    | `CHANGELOG.md`             | nothing    |
| `ui`   | `ooxml-ui`  | `packages/ui` | `packages/ui/CHANGELOG.md` | `core`     |

The UI package depends on the core, never the reverse. Viewers depend on both, so users never install the UI separately.

Releases are automatic: an hourly workflow (and a manual dispatch) looks at the conventional commits on `main` since each package's last release tag and, for every package whose published files changed, bumps the version, writes the changelog, tags, creates a GitHub release and publishes with provenance. The flow is the one used by [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer) and [docx-viewer](https://github.com/ChristopherVR/docx-viewer); scripts that are copies say so in their header comments.

## How a release is decided

`scripts/release-plan.mjs` computes everything from git history. Run it locally at any time; it writes the git-ignored `release-plan.json` and changes nothing else (`--write` also stamps versions into the manifests, which CI does):

```sh
bun run release:plan             # or: node scripts/release-plan.mjs --no-npm   (offline)
```

A package whose `package.json` does not exist yet (the UI before it is merged) is simply left out of the plan.

### Baseline and what counts as a change

Each package's baseline is the newest tag `<npm-name>@x.y.z` that is an ancestor of `HEAD` (`ooxml-core@0.1.0`, `ooxml-ui@0.1.0`). A tag on `HEAD` yields an empty diff, so an already released `HEAD` is a no-op. A package releases when one of **its published files** changed since its baseline:

- core: `src/` except tests (`*.test.ts`, `__tests__/`, fixtures), `scripts/pptx/merge-declarations.mjs`, `tsconfig*.json`, `tsup*.config.ts`, `tsdown.pptx.config.ts`, `package.json`, `LICENSE`, `NOTICE`, `THIRD-PARTY-LICENSES`. Nothing under `packages/` counts for core.
- ui: everything under `packages/ui/` except tests and its `CHANGELOG.md`, including its `package.json`.

`package.json` changes are ignored when they only touch `version`, `scripts`, `devDependencies`, `workspaces` or the range on the other package. Docs, CI, tests, `README.md`, `CHANGELOG.md` and the build orchestration scripts never release anything, and neither does the release commit itself.

### Bump level

The level is the highest Conventional Commit level among the commits since that package's baseline that touch **its own** published files: `!` or a `BREAKING CHANGE:` footer is **major**, `feat` is **minor**, anything else **patch**. A `feat` that only touches tests does not raise the level, and a core commit never raises the ui level. The new version is that bump applied to the highest of the tag version, the npm `latest` and the manifest version. The rule is the same while the version is `0.x` (a breaking change goes to `1.0.0`).

### How core releases affect ui

The ui manifest declares `"ooxml-core": "*"` (the repository root is the core and cannot be a member of its own Bun workspace, so `workspace:*` cannot link; tsconfig paths point the UI at the core source in development). In the published tarball that becomes `^<core version in the repository at publish time>` (see below). A core release therefore does not force a ui release as long as the new core version still satisfies the range ui last shipped with:

| Core change                                 | Core release | ui release                                                              |
| ------------------------------------------- | ------------ | ----------------------------------------------------------------------- |
| Only core files, patch (or minor at >= 1.0) | yes          | no (range still satisfied)                                              |
| Core minor while core is `0.x`              | yes          | yes, patch (`^0.1.0` does not admit `0.2.0`; reason "dependency range") |
| Core major                                  | yes          | yes, patch unless ui's own commits raise it                             |
| Only ui files                               | no           | yes, at the level of ui's own commits                                   |
| Both                                        | yes          | yes, at the level of ui's own commits                                   |
| Neither (tests, docs, CI)                   | no           | no                                                                      |

### First release and baselines

- **Untagged and not on npm (E404)** (the UI before its first publish): released as `initial` at its manifest version (no bump), tagged `ooxml-ui@<manifest version>`. This is what the scheduled run does if nobody bootstraps the package by hand.
- **Published by hand, never tagged** (the manual bootstrap below): the registry version is a baseline, not a reason to release. The planner uses the commit npm recorded as `gitHead` for that version when it is in this history (so UI changes made since the manual publish still release), otherwise `HEAD`, and the workflow pushes the tag `<npm-name>@<published version>` there (the "Adopt baseline tags" step). If the tag lands on `HEAD` because npm had no usable `gitHead`, any change between the manual publish and that run is treated as already released: move the tag by hand if that matters.
- If a baseline tag is deleted for a package that is on npm and was never adopted, the planner sees an untagged published package. Restore the tag rather than rely on that.

`ooxml-core@0.1.0` points at `1272e23`, the commit npm records as `gitHead` of the published `0.1.0`.

## Commit conventions

See [CONTRIBUTING.md](../CONTRIBUTING.md#commit-conventions). The `PR hygiene / Conventional Commits` workflow validates the PR title and every commit subject (`scripts/check-conventional-commits.mjs`): a missing or unknown type fails, header length, casing and a trailing period only warn. Make that check required in the ruleset.

## The release workflow

`.github/workflows/release.yml` has two jobs and two ways to start (hourly schedule, manual dispatch).

**Scheduled or manual run without input** (`release` job, then `publish` job):

1. Check out `main` with full history using `RELEASE_TOKEN` and run the planner with `--write`. If nothing changed the run ends here, which is what a quiet hour looks like.
2. Push the baseline tag of any package that was published by hand (see above).
3. Fail early if the `NPM_PUBLISH` repository variable is not `true`, so nothing is tagged that cannot be published.
4. `bun install`, `bun run build`, `bun run test:package` (pack, install into a clean project, import every entry point), then `scripts/build-released.mjs --smoke`, which runs `build` and `test:package` of each released workspace package that defines them. Typecheck and the unit tests are not repeated: `main` is protected by the `ci-success` check. A failure here stops the run before anything is committed or tagged and the next run retries.
5. Prepend the new section to each released package's changelog with [git-cliff](https://git-cliff.org) (`cliff.toml`, scoped by `--include-path` to that package's published paths and by `--tag-pattern` to its tags). Changelogs are prepend-only: old tags and releases are pruned, so history is never regenerated.
6. Commit `chore(release): bump versions and update changelogs [skip ci]` straight to `main` (released manifests, changelogs, lockfile), retrying with a rebase if `main` moved. `[skip ci]` keeps it from starting CI.
7. Create each tag and GitHub release at that commit, core first (`scripts/release-notes.mjs <key>` writes the body), upload the plan, and prune superseded releases (`scripts/prune-releases.mjs --keep 1`; tags are kept).
8. `publish` job (environment `npm`): check out the release commit, `bun install --frozen-lockfile`, `bun run build`, build the workspace packages, then `node scripts/publish-released.mjs --plan release-plan.json`.

**Manual dispatch with `tag`** re-publishes one existing tag of either package (skips the `release` job), for when a publish failed or was skipped:

```sh
gh workflow run release.yml -f tag=ooxml-core@0.2.0
gh workflow run release.yml -f tag=ooxml-ui@0.1.0
```

`scripts/publish-released.mjs` publishes the released packages in dependency order (core, then ui), each with `npm publish --provenance --access public` run **from that package's own directory**. Before uploading it checks that the manifest on disk is the version being published, that no dependency uses `file:`/`link:` or a `workspace:` range on anything but the sibling, that a sibling required by the package is already on npm, and that the version is not already published (it is skipped if so, so re-runs are safe). A version older than the registry's `latest` is published under the `old` dist-tag. `--dry-run` prints the commands without publishing.

`npm publish` does not rewrite `workspace:` ranges (Bun and pnpm do, npm does not), so for the upload only, the script rewrites the ui manifest's `*` (or `workspace:*`) on core to `^<core version on disk>` and restores the file afterwards. The repository keeps `*`, so `bun install` never has to resolve a core version that is published later in the same release; the tarball carries a real caret range. Because core is stamped and published before ui in the same run, that range always points at a version that exists on npm.

### Authentication: trusted publishing, no secrets

There is no npm token anywhere. The `publish` job requests an OIDC token (`id-token: write`), npm >= 11.5.1 exchanges it for a short-lived publish credential, and `--provenance` attaches the attestation. Each package needs its own trusted publisher on npmjs.com (package Settings, Trusted Publisher, GitHub Actions). Both use the same repository, workflow and environment; renaming the workflow or the environment breaks publishing:

| Package                     | Organization or user | Repository   | Workflow filename | Environment |
| --------------------------- | -------------------- | ------------ | ----------------- | ----------- |
| `ooxml-core` | `ChristopherVR`      | `ooxml-core` | `release.yml`     | `npm`       |
| `ooxml-ui`  | `ChristopherVR`      | `ooxml-core` | `release.yml`     | `npm`       |

## One-time setup

1. **Rulesets on `main`**: require the `ci-success` and `Conventional Commits` checks; add a bypass for repository admins (the release commit is pushed straight to `main`).
2. **`RELEASE_TOKEN` secret**: a fine-grained personal access token with `Contents: Read and write` on this repository, owned by a repository admin. The default `GITHUB_TOKEN` cannot be granted ruleset bypass, so without it the release commit and baseline tags cannot be pushed.
3. **`NPM_PUBLISH` repository variable** set to `true`.
4. **`npm` environment** (Settings, Environments), optionally with required reviewers.
5. The trusted publisher for `ooxml-core` and for `ooxml-ui`. npm only lets you add a trusted publisher to a package that exists, so each new package name needs one manual publish first.

### First publish of a new package name (manual bootstrap)

The packages were first published as `@christophervr/ooxml-core` and `@christophervr/office-ui`; from 2026-10 they are the unscoped `ooxml-core` and `ooxml-ui`. Publish core before the UI.

Trusted publishing cannot create a package. Either publish the first version by hand, or let the workflow try (it will fail at the `npm publish` step with a 404/permission error until the publisher exists, and the retry is `gh workflow run release.yml -f tag=ooxml-ui@<version>`). By hand, from a clean checkout of the commit you want to publish, logged in to npm with the account that will own the packages:

```sh
git pull
bun install
bun run build                         # core first: the UI resolves it through dist
bun run --cwd packages/ui build
npm login                             # 2FA
node scripts/publish-released.mjs --tag ooxml-core@<version> --manual
node scripts/publish-released.mjs --tag ooxml-ui@<version> --manual
```

(`publish-released.mjs` does what the workflow does: it rewrites the ui manifest's `*` on core to `^<core version on disk>` for the upload only and restores the file, refuses to publish unless that core version is already on npm, and publishes from `packages/ui`. `--manual` only omits provenance, which needs a CI OIDC token; that is expected for the first publish. `--dry-run` prints the command first.)

Then configure the trusted publisher (table above) on npmjs.com. Nothing else is needed: the next scheduled run sees `ooxml-core@<version>` and `ooxml-ui@<version>` on npm with no tag, treats it as the baseline (see "First release and baselines"), pushes the tag, and later UI changes release normally. If you want the tag immediately: `git tag ooxml-ui@<version> <commit> && git push origin ooxml-ui@<version>`.

## Housekeeping

- `prune-releases.yml` runs daily and keeps only the newest GitHub release per package (tags and npm versions are untouched). Run it manually with `dry_run` to preview.
- `scripts/check-changelog-sections.mjs` (CI) fails if any changelog (the root one and `packages/*/CHANGELOG.md`) contains a stub section, because the files are prepend-only and cannot be regenerated.
- `bun run test:scripts` (CI) runs the planner, publish-guard, commit and changelog tests; the planner tests build throwaway git repositories shaped like this one (core only, ui only, both, initial ui, manual bootstrap).
- Scripts that are candidates to move into a shared package: `release-plan.mjs` (only the package table differs), `check-conventional-commits.mjs`, `check-changelog-sections.mjs`, `release-notes.mjs`, `prune-releases.mjs`, `publish-released.mjs`. They are copied between this repository, `docx-viewer` and `pptx-viewer`.
