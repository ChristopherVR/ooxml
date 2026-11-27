# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@2.0.1) - 2026-11-27

### Bug Fixes

- **teams:** Accept workspace ids up to the 128-character room limit ([66d9e54](https://github.com/ChristopherVR/ooxml/commit/66d9e54ae76010beef52939d857cacf4da95ee9d))

### Documentation

- Rewrite the OpenTeams viewer and server READMEs ([15e0b11](https://github.com/ChristopherVR/ooxml/commit/15e0b11b2dcf26edfe165577124e6a47b0ecc79e))

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@2.0.0) - 2026-11-21

<!-- release-notes: 2.0 -->

For pages without a framework, `mountTeams`, `createTeams`, `defineTeamsApp` and the `<teams-app>` element keep their API. What changes is underneath them: the client speaks Teams protocol version 2, private chats get new ids and their membership is decided by the server, and a crafted update can no longer stall the page. Deploy this release and `openteams-server` 2.0.0 (or your own server, updated to protocol 2) in one step, and let clients finish syncing first, because changes still unsent at the upgrade are discarded. Upgrade together with the packages under "Released together", and read the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#openteams) before you deploy.

**[Introducing OOXML 2.0](https://christophervr.github.io/ooxml/releases/2.0/)** tells the story of this release. Every change to OpenTeams is listed in [the release notes](https://christophervr.github.io/ooxml/releases/2.0/#changes-openteams) ([Markdown on GitHub](https://github.com/ChristopherVR/ooxml/blob/main/docs/releases/2.0.md#changes-openteams)).

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): `ooxml-core/teams`, the client behind this package: protocol 2, the new chat ids and the update budget.
- [`ooxml-ui` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/ui/CHANGELOG.md): the `<teams-app>` element and the Teams primitives this package renders.
- [`openteams-server` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/server/CHANGELOG.md): the reference server that speaks protocol 2; deploy it with this release.
- [`openteams-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/react/CHANGELOG.md): the same release for React apps.
- [`openteams-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/vue/CHANGELOG.md): the same release for Vue apps.
- [`openteams-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/angular/CHANGELOG.md): the same release for Angular apps.
- [`openteams-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/svelte/CHANGELOG.md): the same release for Svelte apps.
- [`openteams-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/solid/CHANGELOG.md): the same release for Solid apps.

## [1.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@1.1.2) - 2026-11-13

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@1.1.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@1.1.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@1.0.0) - 2026-10-29

### Features

- **teams:** Release openteams 1.0 on the versioned sync protocol ([e035417](https://github.com/ChristopherVR/ooxml/commit/e035417be1285095f9bbe1b9655743dcd059b55a))

## [0.2.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.2.1) - 2026-10-23

### Other

- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))

### Refactor

- **teams:** Move bindTeams into ooxml-ui/teams on the binding kit ([fd417a5](https://github.com/ChristopherVR/ooxml/commit/fd417a515f76b580af69acc5b1e319d15d41fee5))

## [0.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.2.0) - 2026-10-16

### Features

- **teams:** Pass a locale through every binding and demo ([bcea500](https://github.com/ChristopherVR/ooxml/commit/bcea500df4863fc296ff5020a6b6af0f8631d072))

### Testing

- **teams:** Cover chats, the calendar and reminders in the browser ([544d1f0](https://github.com/ChristopherVR/ooxml/commit/544d1f0adb31802033102e009a66c338bb0c3df6))

## [0.1.8](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.1.8) - 2026-10-11

### Styling

- Fix the oxfmt and oxlint warnings in the viewers ([db0314f](https://github.com/ChristopherVR/ooxml/commit/db0314f319e3d0508960872260792215477d9ca5))

## [0.1.7](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.1.7) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.1.6](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.1.6) - 2026-10-08

### Bug Fixes

- **teams:** Accept a host class in every binding ([629bb11](https://github.com/ChristopherVR/ooxml/commit/629bb118e6adc0ad21c78a34a60c884586b09869))

## [0.1.5](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.1.5) - 2026-10-07

## [0.1.4](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.1.4) - 2026-10-07

## [0.1.3](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.1.3) - 2026-10-07

## [0.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.1.2) - 2026-10-06

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-vanilla-viewer@0.1.1) - 2026-10-05

### Features

- **web-component:** Add the team-workspace app, bindings and reference server (by @ChristopherVR) ([36916ee](https://github.com/ChristopherVR/ooxml/commit/36916eeef144e9260144b5f0e87ddee6c8021b75))
- **release:** Package the bindings as openteams-*-viewer (by @ChristopherVR) ([2d5101e](https://github.com/ChristopherVR/ooxml/commit/2d5101e55f477834b591962409a37b078960dd1a))

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

### Refactor

- **web-component:** Take <teams-app> from ooxml-ui/teams and add oxlint (by @ChristopherVR) ([c8175e9](https://github.com/ChristopherVR/ooxml/commit/c8175e98a795206693c247a499112fb9ee079d15))

### Build & CI

- **release:** Release the imported viewers from the root planner (by @ChristopherVR) ([86984f1](https://github.com/ChristopherVR/ooxml/commit/86984f1bbc4f4df828e8f1760586a3f0b7321b0b))
- **pages:** Serve the launcher and every viewer's docs and demos from one site (by @ChristopherVR) ([1048feb](https://github.com/ChristopherVR/ooxml/commit/1048febca7d914d11ca08a0c7a6b4c80622ca867))
- Fix the viewer checks that failed on the first run (by @ChristopherVR) ([2445d3a](https://github.com/ChristopherVR/ooxml/commit/2445d3a58688ea0917b15b48088afa03fc3562b7))

### Chores

- **viewers:** Import the teams-viewer repository with its history (by @ChristopherVR) ([7d29cec](https://github.com/ChristopherVR/ooxml/commit/7d29cecf88ac551cc5a80d2e9e35c65e130ab0de))

