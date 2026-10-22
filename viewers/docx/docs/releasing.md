# Releasing packages

The Word viewer packages are released by the `ooxml` repository's own flow, together with `ooxml-core`
and `ooxml-ui`. There is no separate release workflow in `viewers/docx`. The whole flow (the planner, the hourly `release.yml`, trusted publishing, changelogs and the
one-time setup) is described in the repository's
[releasing guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md). Never run
`npm publish` or push release tags by hand.

## What this viewer publishes

Each package has its own version line and its own `<npm-name>@<version>` git tag. The mapping from
directory to npm name is `scripts/viewer-packages.mjs` at the repository root.

| Package               | Dir                | Notes                                                       |
| --------------------- | ------------------ | ----------------------------------------------------------- |
| `docx-core`           | `packages/core`    | Thin re-export of `ooxml-core/docx` and `/docx/embedded`.   |
| `docx-react-viewer`   | `packages/react`   | React entry (`WordEditor`).                                 |
| `docx-vue-viewer`     | `packages/vue`     | Vue 3 entry.                                                |
| `docx-angular-viewer` | `packages/angular` | Angular entry (`WordEditorComponent`).                      |
| `docx-svelte-viewer`  | `packages/svelte`  | Svelte 5 entry.                                             |
| `docx-solid-viewer`   | `packages/solid`   | Solid entry.                                                |
| `docx-vanilla-viewer` | `packages/vanilla` | Plain `<docx-editor>` registration and `mountEditor`.       |
| `docx-viewer-mcp`     | `mcp`              | The Word MCP server (see its README and the root `mcp.md`). |

`web-component` and `bindings` are `private: true` and are never published.
`scripts/build-packages.mjs` bundles them into every framework package, so a tarball imports only
`docx-core`, `ooxml-core`, `ooxml-ui`, the ProseMirror libraries and its framework peer.
`ooxml-ui` and `docx-core` stay real registry dependencies (never bundled), and the legacy `.doc`
codecs from `ole2` are inlined inside `ooxml-core`, so `ole2` is not a dependency of anything here.

## What triggers a release

The root planner (`scripts/release-plan.mjs`, using the table in `scripts/viewer-packages.mjs`)
looks at the conventional commits since each package's last tag:

- A non-test change under `packages/<name>/` releases that package.
- A non-test change under `packages/web-component` or `packages/bindings` releases every framework
  package, because each ships that code. `scripts/build-packages.mjs` and `tsconfig.release.json`
  release everything in this viewer.
- A release of `docx-core` re-releases the framework packages that depend on it.
- Tests, docs, demos and CI never release anything. A change under `demos/docx/` or `e2e/docx/` only
  runs this viewer's checks in CI.

The editor's own code lives in `src/ui/src/docx` and ships in `ooxml-ui`; a change there releases
`ooxml-ui`, and the viewer packages pick it up through their `ooxml-ui` range.

## Checks before pushing

Run `bun run verify` at the repository root. For this viewer the equivalent local checks, from
`viewers/docx`, are:

```sh
bun run typecheck && bun run test && bun run test:scripts
bun run build:packages && bun run check:published && bun run pack:smoke
```

`check:published` fails if a tarball imports an internal package or `ole2`, and `pack:smoke` installs
the tarballs into a clean consumer and loads DOCX and legacy `.doc` through every framework package.
