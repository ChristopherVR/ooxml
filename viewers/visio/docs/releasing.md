# Releasing

The seven placeholder packages use the hourly `release.yml` workflow, conventional commits and independent package tags. Changes to `npm-placeholders/<name>` release that package. Viewer, demo and binding changes do not release placeholders.

The publish job runs in the GitHub environment `npm`, requests `id-token: write`, and publishes with npm OIDC and provenance. Each npm trusted publisher must name `ChristopherVR/visio-viewer`, `release.yml`, and environment `npm`. No npm token is required. `NPM_PUBLISH=true` enables publishing.

Run `gh workflow run release.yml` to release now. To retry a failed publish, dispatch with `-f tag=<npm-name>@<version>`; the workflow checks out that exact tag and skips versions already present on npm. Never publish release tags or packages manually after the initial bootstrap.

The existing packages are explicit placeholders. Functional releases need a separate package build and consumer validation workflow before expanding the release table. Release tooling provenance and its license are recorded in `scripts/RELEASE-PROVENANCE.md`.
