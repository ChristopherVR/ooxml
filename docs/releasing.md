# Releasing

This repository publishes one package, `@christophervr/ooxml-core`, to npm. Releases are automatic: an hourly workflow (and a manual dispatch) looks at the conventional commits on `main` since the last release tag and, if published files changed, bumps the version, writes the changelog, tags, creates a GitHub release and publishes with provenance. The flow is the one used by [pptx-viewer](https://github.com/ChristopherVR/pptx-viewer) and [docx-viewer](https://github.com/ChristopherVR/docx-viewer), in its single-package form; scripts that are copies say so in their header comments.

## How a release is decided

`scripts/release-plan.mjs` computes everything from git history. Run it locally at any time; it writes the git-ignored `release-plan.json` and changes nothing else:

```sh
bun run release:plan             # or: node scripts/release-plan.mjs --no-npm   (offline)
```

The baseline is the newest tag `@christophervr/ooxml-core@x.y.z` that is an ancestor of `HEAD`. A tag on `HEAD` yields an empty diff, so an already released `HEAD` is a no-op. The package releases when a **published file** changed since the baseline:

- `src/` except tests (`*.test.ts`, `__tests__/`, fixtures),
- `scripts/pptx/merge-declarations.mjs`, `tsconfig*.json`, `tsup*.config.ts`, `tsdown.pptx.config.ts`,
- `package.json`, ignoring changes that only touch `version`, `scripts` or `devDependencies`,
- `LICENSE`, `NOTICE`, `THIRD-PARTY-LICENSES`.

Docs, CI, tests, `README.md`, `CHANGELOG.md` and the build orchestration scripts (`scripts/build.mjs`, `scripts/ensure-built.mjs`) never release anything, and neither does the release commit itself.

The bump level is the highest Conventional Commit level among the commits since the baseline that touch published files: `!` or a `BREAKING CHANGE:` footer is **major**, `feat` is **minor**, anything else **patch**. A `feat` that only touches tests does not raise the level. The new version is that bump applied to the highest of the tag version, the npm `latest` and the manifest version. The rule is the same while the version is `0.x` (a breaking change goes to `1.0.0`); change `bumpVersion` deliberately if you want 0.x breaking changes to bump the minor.

### Baseline

`@christophervr/ooxml-core@0.1.0` is the baseline tag. It points at `1272e23`, the commit npm records as `gitHead` of the published `0.1.0`. Commits made after it that touched no published file (build orchestration, CI, docs, test fixtures) do not release anything, so the first scheduled run is a no-op. If the tag is ever missing, the planner sees the package on npm without a tag and releases a patch bump from the npm version, so restore the tag rather than let that happen.

## Commit conventions

See [CONTRIBUTING.md](../CONTRIBUTING.md#commit-conventions). The `PR hygiene / Conventional Commits` workflow validates the PR title and every commit subject (`scripts/check-conventional-commits.mjs`): a missing or unknown type fails, header length, casing and a trailing period only warn. Make that check required in the ruleset.

## The release workflow

`.github/workflows/release.yml` has two jobs and two ways to start (hourly schedule, manual dispatch).

**Scheduled or manual run without input** (`release` job, then `publish` job):

1. Check out `main` with full history using `RELEASE_TOKEN` and run the planner with `--write`. If nothing changed the run ends here, which is what a quiet hour looks like.
2. Fail early if the `NPM_PUBLISH` repository variable is not `true`, so nothing is tagged that cannot be published.
3. `bun install`, `bun run build`, `bun run test:package` (pack, install into a clean project, import every entry point). Typecheck and the unit tests are not repeated: `main` is protected by the `ci-success` check, so what is on `main` has passed CI. A failure here stops the run before anything is committed or tagged and the next run retries.
4. Prepend the new section to `CHANGELOG.md` with [git-cliff](https://git-cliff.org) (`cliff.toml`, scoped to the published paths and the tag pattern). The changelog is prepend-only: old tags and releases are pruned, so history is never regenerated.
5. Commit `chore(release): bump versions and update changelogs [skip ci]` straight to `main` (version, changelog, lockfile), retrying with a rebase if `main` moved. `[skip ci]` keeps it from starting CI.
6. Create the tag and GitHub release at that commit (`scripts/release-notes.mjs` writes the body), upload the plan, and prune superseded releases (`scripts/prune-releases.mjs --keep 1`; tags are kept).
7. `publish` job (environment `npm`): check out the release commit, `bun install --frozen-lockfile`, `bun run build`, then `node scripts/publish-released.mjs --plan release-plan.json`.

**Manual dispatch with `tag`** re-publishes one existing tag (skips the `release` job), for when a publish failed or was skipped:

```sh
gh workflow run release.yml -f tag=@christophervr/ooxml-core@0.2.0
```

`scripts/publish-released.mjs` runs `npm publish --provenance --access public` after checking that the manifest on disk is the version being published, that no dependency uses `workspace:` or `file:`, and that the version is not already on npm (it is skipped if so, so re-runs are safe). A version older than the registry's `latest` is published under the `old` dist-tag. `--dry-run` prints the command without publishing.

### Authentication: trusted publishing, no secrets

There is no npm token anywhere. The `publish` job requests an OIDC token (`id-token: write`), npm >= 11.5.1 exchanges it for a short-lived publish credential, and `--provenance` attaches the attestation. The trusted publisher on npmjs.com is configured for `ChristopherVR/ooxml-core`, workflow `release.yml`, environment `npm`; renaming the workflow or the environment breaks publishing.

## One-time setup

1. **Rulesets on `main`**: require the `ci-success` and `Conventional Commits` checks; add a bypass for repository admins (the release commit is pushed straight to `main`).
2. **`RELEASE_TOKEN` secret**: a fine-grained personal access token with `Contents: Read and write` on this repository, owned by a repository admin. The default `GITHUB_TOKEN` cannot be granted ruleset bypass, so without it the release commit cannot be pushed.
3. **`NPM_PUBLISH` repository variable** set to `true`.
4. **`npm` environment** (Settings, Environments), optionally with required reviewers.
5. The npm trusted publisher described above (already configured).

## Housekeeping

- `prune-releases.yml` runs daily and keeps only the newest GitHub release per package (tags and npm versions are untouched). Run it manually with `dry_run` to preview.
- `scripts/check-changelog-sections.mjs` (CI) fails if the changelog contains a stub section, because the file is prepend-only and cannot be regenerated.
- Scripts that are candidates to move into a shared package: `release-plan.mjs` (only the package table differs), `check-conventional-commits.mjs`, `check-changelog-sections.mjs`, `release-notes.mjs`, `prune-releases.mjs`, `publish-released.mjs`. They are copied between this repository, `docx-viewer` and `pptx-viewer`.
