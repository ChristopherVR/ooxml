# Releasing packages

The Excel packages are released by the repository's single automated flow, not by a flow of their own. This viewer is `viewers/xlsx` in the [ooxml](https://github.com/ChristopherVR/ooxml) repository, together with `ooxml-core` (`src/core`) and `ooxml-ui` (`src/ui`). The whole process (how a release is decided, the bump level, the hourly workflow, trusted publishing and the one-time repository setup) is described in [docs/releasing.md](https://github.com/ChristopherVR/ooxml/blob/main/docs/releasing.md) at the repository root.

## What is published from here

Each package has its own version line, git tag `<npm-name>@<version>` and changelog. They are declared in `scripts/viewer-packages.mjs` at the repository root.

| Package                            | Dir                | Notes                                                                      |
| ---------------------------------- | ------------------ | -------------------------------------------------------------------------- |
| `@christophervr/xlsx-core`         | `packages/core`    | Workbook model entry point over `ooxml-core`; no logic of its own.         |
| `@christophervr/xlsx-react-viewer` | `packages/react`   | Self-contained editor package; bundles the web component and the bindings. |
| `xlsx-vue-viewer`                  | `packages/vue`     | Same shape.                                                                |
| `xlsx-angular-viewer`              | `packages/angular` | Same shape.                                                                |
| `xlsx-svelte-viewer`               | `packages/svelte`  | Same shape.                                                                |
| `xlsx-solid-viewer`                | `packages/solid`   | Same shape.                                                                |
| `xlsx-vanilla-viewer`              | `packages/vanilla` | Same shape; also owns the plain `<xlsx-editor>` entry.                     |
| `xlsx-viewer-mcp`                  | `mcp`              | Standalone MCP server, see its README.                                     |

The internal packages `packages/web-component` (`xlsx-web-component`) and `packages/bindings` (`xlsx-bindings`) are `private` and never published: `scripts/build-packages.mjs` bundles them into every framework package, so a non-test change in either releases all six. A published tarball imports only `@christophervr/xlsx-core`, `ooxml-core`, `ooxml-ui` and its framework peer; `ole2` is inlined inside `ooxml-core`. `ooxml-core` and `ooxml-ui` are real registry dependencies, never bundled.

## Checks from this folder

Run these from `viewers/xlsx` after a root `bun install`:

```sh
bun run build:packages
bun run check:published   # no tarball imports an internal package or ole2
bun run pack:smoke        # install the tarballs into a clean consumer and load .xlsx and .xls
```

The release planner, the publish guards and the workflow run from the repository root (`bun run release:plan`, `node scripts/publish-released.mjs --dry-run`, `.github/workflows/release.yml`). Never run `npm publish` or push release tags by hand.

## Commit conventions

The bump level comes from the commit type (`!` or `BREAKING CHANGE:` is major, `feat` is minor, anything else patch), and the files a commit touches decide whether a package releases. Rules and examples are in [CONTRIBUTING.md](https://github.com/ChristopherVR/ooxml/blob/main/CONTRIBUTING.md#commit-conventions).
