# Releasing

Seven packages ship from `packages/{core,react,vue,angular,svelte,solid,vanilla}`. They use the hourly `release.yml` workflow, conventional commits and independent package tags. Viewer changes in `src/` and adapter changes in `packages/bindings/src/` release all six viewer packages. A core release also releases dependent viewers. Documentation and test-only changes do not release packages unless they change a shipped package README.

The build uses released `ooxml-core/visio`, includes both browser workers, emits TypeScript declarations and supplies native Svelte source and Solid/Svelte server entries. `npm run check` verifies source, adapters, docs, builds, release planning and all seven packed consumers. `npm run test:browser` verifies browser behavior. Both run before release tags are created. The publish job rebuilds the exact release revision and validates its tarballs before publishing.

The publish job runs in GitHub environment `npm`, requests `id-token: write`, and publishes with npm OIDC and provenance. Each npm trusted publisher must name `ChristopherVR/visio-viewer`, `release.yml`, and environment `npm`. No npm token is required. `NPM_PUBLISH=true` enables publishing.

Run `gh workflow run release.yml` to release now. Retry a failed publish with `-f tag=<npm-name>@<version>`; the workflow checks out that exact tag and skips versions already on npm. Never publish release tags or packages manually. Historical placeholder tags remain retryable from their own revisions.

Release tooling provenance and its license are recorded in `scripts/RELEASE-PROVENANCE.md`.
