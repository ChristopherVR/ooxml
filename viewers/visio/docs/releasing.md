# Releasing

The Visio packages are released by the `ooxml` repository's own flow, described in the repository's `docs/releasing.md`: the root `scripts/release-plan.mjs` versions each package from the conventional commits since its last `<npm-name>@<version>` tag, and the root `.github/workflows/release.yml` tags, writes changelogs and publishes with npm trusted publishing (OIDC and provenance, no token). The Visio package table is `viewer('visio', ...)` in the root `scripts/viewer-packages.mjs` (`viewers/visio/scripts/viewer-packages.mjs` is a view of it for the build scripts). The `.github/` folder of this viewer is inert.

Eight packages ship: `visio-core` and `visio-{react,vue,angular,svelte,solid,vanilla}-viewer` from `packages/{core,react,vue,angular,svelte,solid,vanilla}`, and `visio-viewer-mcp` from `mcp`. A change to `src/` or `packages/bindings/src/` here, or to `scripts/build-packages.mjs` or `tsconfig.build.json`, releases all six viewer packages; a change to one package's own directory releases that package. The viewer implementation itself is `ooxml-ui/visio` (`src/ui/src/visio`), released with `ooxml-ui`, and the packages depend on `ooxml-core` and `ooxml-ui` by workspace range: the release flow rewrites each to a caret range on the version released with it. Documentation and test-only changes do not release packages unless they change a shipped package README.

`bun run check` in `viewers/visio` verifies source, adapters, docs, builds and all packed consumers (`bun run test:packages`), and `bun run test:browser` verifies browser behavior. The build includes both browser workers, emits TypeScript declarations and supplies native Svelte source and Solid/Svelte server entries.

Never run `npm publish` or push release tags by hand. A failed publish is retried through the release workflow's `tag` input, which skips versions already on npm.

Release tooling provenance is recorded in `scripts/RELEASE-PROVENANCE.md`.
