# Releasing packages

Seven unscoped packages are published to npm under the product name **OpenTeams** (`openteams`
itself belongs to an unrelated project, so every name is longer). Each has its **own version
line**, bumped only when it actually changes, and its own git tag `<npm-name>@<version>` (for
example `openteams-server@0.4.0`).

Releases are run from the **repository root**, not from this folder. The planner
(`scripts/release-plan.mjs`), the viewer package table (`scripts/viewer-packages.mjs`), the
workflow (`.github/workflows/release.yml`), trusted publishing and the one-time setup are described
in the root [`docs/releasing.md`](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md).
This folder has no release scripts, no workflow of its own (its `.github/` is inert) and no
`release:plan` or `publish:local` command.

| Package                    | Dir                | Depends on (registry)                                 | Bundled in                 |
| -------------------------- | ------------------ | ----------------------------------------------------- | -------------------------- |
| `openteams-react-viewer`   | `packages/react`   | `ooxml-core`, `ooxml-ui`, `lit`; peer `react`         | `packages/web-component`   |
| `openteams-vue-viewer`     | `packages/vue`     | `ooxml-core`, `ooxml-ui`, `lit`; peer `vue`           | `packages/web-component`   |
| `openteams-angular-viewer` | `packages/angular` | `ooxml-core`, `ooxml-ui`, `lit`; peer `@angular/core` | `packages/web-component`   |
| `openteams-svelte-viewer`  | `packages/svelte`  | `ooxml-core`, `ooxml-ui`, `lit`; peer `svelte`        | `packages/web-component`   |
| `openteams-solid-viewer`   | `packages/solid`   | `ooxml-core`, `ooxml-ui`, `lit`; peer `solid-js`      | `packages/web-component`   |
| `openteams-vanilla-viewer` | `packages/vanilla` | `ooxml-core`, `ooxml-ui`, `lit`                       | `packages/web-component`   |
| `openteams-server`         | `server`           | `ws`, `yjs`, `y-protocols`, `lib0`                    | nothing (ships its source) |

`teams-viewer` (`packages/web-component`, the re-export and bind layer of `<teams-app>`) and the
demos are `private` and never published. `scripts/build-packages.mjs` inlines the web component,
CSS included, into every framework bundle (vite library build) and flattens its declarations into
each package's `dist/index.d.ts` (rollup-plugin-dts), so a tarball imports only what its manifest
declares. Svelte ships `dist/Teams.svelte` as source next to a bundled `dist/runtime.js` (the store
and the helpers, under `openteams-svelte-viewer/runtime`); Angular is a bundled ES module with its
decorators applied at runtime (it needs the JIT compiler in the consuming app). The server is plain
ESM and is not built; its `openteams-server` command is `server/bin.mjs`.

In the monorepo the framework manifests depend on `ooxml-core` and `ooxml-ui` as `*`; the root
release flow rewrites those ranges to caret ranges on the versions released with them. A change
that needs both the core and this viewer lands in one push and is released together.

## What decides a release here

The root planner reads `bundled` and `globals` from the `teams` entry of
`scripts/viewer-packages.mjs`:

- **Own files.** Any non-test file under a package directory releases it. Tests, `CHANGELOG.md`
  and manifest edits that only touch `version`, `scripts` or `devDependencies` never do.
- **Bundled internal package.** A non-test change under `packages/web-component` releases every
  framework package.
- **Shared build pipeline.** `scripts/build-packages.mjs` and `tsconfig.release.json` release every
  framework package. The server opts out (`global: false`): only `server/**` releases it.

The bump level is the highest Conventional Commit level among the commits that touch published
files in the package's scope: `!` / `BREAKING CHANGE:` is major, `feat` minor, anything else patch.
Use `docs(teams): ...`, `fix(server): ...` and similar scopes (see the root `AGENTS.md`).

## Checks that run for this viewer

`bun run check:published` (no `workspace:`, `file:` or `link:` range and no private package in a
published manifest; it also scans `.mjs` files and the server directory) and `bun run pack:smoke`
(install and import the tarballs). `bun run test:scripts` runs the viewer's own script tests
(`scripts/*.test.mjs`).

Never run `npm publish` or push release tags by hand, and never edit versions or `CHANGELOG.md`
files; the release commit does. A new package name needs the manual bootstrap and trusted-publisher
setup described in the root page.
