# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

For Svelte 5 apps, the `Teams.svelte` component and the `teamsStore` store from `openteams-svelte-viewer/runtime` are unchanged. What changes is underneath them: the client speaks Teams protocol version 2, private chats get new ids and their membership is decided by the server, and a crafted update can no longer stall the page. Deploy this release and `openteams-server` 2.0.0 (or your own server, updated to protocol 2) in one step, and let clients finish syncing first, because changes still unsent at the upgrade are discarded. Upgrade together with the packages under "Released together", and read the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#openteams) before you deploy.

### Breaking changes

- The client speaks Teams protocol version 2. `openteams-server` 2.0.0 closes a client that announces another version, with a message that names both versions, so a 1.x server and this client do not work together. Deploy the app and the server together.
- A new 1:1 chat id is `dm-` and 32 hex digits, and a new group chat id is `g-` with a random part bound to its creator. Meeting chats are group chats and no longer use `mt-<meeting id>`. A private chat's call uses the signaling room `call-` followed by the chat's room name, which the server opens only to the chat's members. Existing chats keep working for the people their rooms list, and 1:1 chats keep their history; a chat that never got a room on the server under an old id cannot be started any more, so start a new one.
- Local snapshots are kept per protocol version, so changes that were still unsent when the app was upgraded are discarded.
- An update that holds more than one struct per 4 bytes, or more than one deleted range per 2 bytes, is refused before it is decoded, in the browser and on the server. The budget grows with the update, so the first sync of a real workspace stays under half of it. If you also call `ooxml-core/collab` yourself, `DEFAULT_UPDATE_LIMITS` is replaced by `defaultUpdateLimits(bytes)` ([migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#collaboration-updates-are-bounded-and-a-refusal-stops-the-sync)).
- The package requires Node.js 22.3 or newer wherever it runs under Node (builds, tests, server rendering), and declares `"engines": { "node": ">=22.3" }` (it declared Node 20 before). What runs in the browser is unaffected.

### Security

- Private chat membership is decided by the server. Editing the shared chat directory no longer gets anyone into a chat: only a chat's members can change its row, a chat's kind and creator are fixed, and clients no longer copy the directory into the room's member list.
- One small crafted update could make the browser (or the server) allocate millions of objects and stop. Such updates are now refused before they are decoded.
- The package is released by a pipeline in which no third-party code can push, tag or publish, and the signed tarball attached to each GitHub release is checked against the integrity the npm registry records.

### Added

- When sharing stops because an update was refused, Teams shows a notice that says why, in the display language.

### Changed

- `--office-teams-brand-hover` now derives from `--office-teams-brand`, so a host that rebrands Teams gets a matching hover colour.
- Avatars (`office-ui-avatar`) pick black or white initials by computed contrast against their colour, and generated avatar colours are hex values.
- The supported browsers are the last two major versions of Chrome, Edge, Firefox and Safari, and Safari 16.4 at least. [Browser support](https://github.com/ChristopherVR/ooxml/blob/main/docs/browser-support.md) lists them with the Safari behaviours the packages handle.
- The package is built and tested against current releases of its dependencies.

### Fixed

- People added to or leaving a group at the same time are all kept, on every client. These changes no longer overwrite each other, and they can no longer stop a person's workspace sync.
- A change made offline by someone who has since been removed from a group is undone by the server instead of stopping that person's sync.
- If a group's creator closes the app right after creating it, the people listed are let into the chat the next time the creator opens it.
- Leaving a group while offline takes effect in the chat itself when the app next connects, even if it was closed in between.
- A refusal from the server no longer reaches the person as a bare close code, and the client no longer reconnects in a loop after one.
- The channel a new workspace starts with is shown in the display language ("Allgemein" in German, and so on) while it keeps its original name and topic. A renamed channel shows its name as typed.

### Accessibility

- F6 and Shift+F6 move between the app bar, the chat list, the header, the messages and the message box, as in the Word, Excel and PowerPoint editors.
- Teams search and the chat composer have control boundaries with at least 3:1 contrast.
- Selected, current and pressed states in Teams lists keep a system highlight in forced-colours mode instead of disappearing.
- Every change is checked in a real browser before it can be released: Tab never gets stuck, F6 reaches every region, dialogs keep focus and give it back, and axe-core runs on the main screens in light mode, dark mode and on a phone.

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): `ooxml-core/teams`, the client behind this package: protocol 2, the new chat ids and the update budget.
- [`ooxml-ui` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/ui/CHANGELOG.md): the `<teams-app>` element and the Teams primitives this package renders.
- [`openteams-server` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/server/CHANGELOG.md): the reference server that speaks protocol 2; deploy it with this release.
- [`openteams-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/react/CHANGELOG.md): the same release for React apps.
- [`openteams-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/vue/CHANGELOG.md): the same release for Vue apps.
- [`openteams-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/angular/CHANGELOG.md): the same release for Angular apps.
- [`openteams-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/solid/CHANGELOG.md): the same release for Solid apps.
- [`openteams-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/vanilla/CHANGELOG.md): the same release for pages without a framework.

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@1.1.1) - 2026-11-06

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@1.1.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([8e60b14](https://github.com/ChristopherVR/ooxml/commit/8e60b14e1d8ee3a1cb507670caa15a611632ef26))
- **deps:** Merge production dependency updates ([#236](https://github.com/ChristopherVR/ooxml/issues/236)) ([7eb137b](https://github.com/ChristopherVR/ooxml/commit/7eb137b03906ded043ce0e4adf83e01fcdaa44a9))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@1.0.0) - 2026-10-29

### Features

- **teams:** Release openteams 1.0 on the versioned sync protocol ([e035417](https://github.com/ChristopherVR/ooxml/commit/e035417be1285095f9bbe1b9655743dcd059b55a))

## [0.2.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.2.1) - 2026-10-23

### Other

- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))

### Refactor

- **teams:** Move bindTeams into ooxml-ui/teams on the binding kit ([fd417a5](https://github.com/ChristopherVR/ooxml/commit/fd417a515f76b580af69acc5b1e319d15d41fee5))

## [0.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.2.0) - 2026-10-16

### Features

- **teams:** Pass a locale through every binding and demo ([bcea500](https://github.com/ChristopherVR/ooxml/commit/bcea500df4863fc296ff5020a6b6af0f8631d072))

### Testing

- **teams:** Cover chats, the calendar and reminders in the browser ([544d1f0](https://github.com/ChristopherVR/ooxml/commit/544d1f0adb31802033102e009a66c338bb0c3df6))

## [0.1.8](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.1.8) - 2026-10-11

### Styling

- Fix the oxfmt and oxlint warnings in the viewers ([db0314f](https://github.com/ChristopherVR/ooxml/commit/db0314f319e3d0508960872260792215477d9ca5))

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

## [0.1.7](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.1.7) - 2026-10-08

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

## [0.1.6](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.1.6) - 2026-10-08

### Bug Fixes

- **teams:** Accept a host class in every binding ([629bb11](https://github.com/ChristopherVR/ooxml/commit/629bb118e6adc0ad21c78a34a60c884586b09869))

## [0.1.5](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.1.5) - 2026-10-07

### Refactor

- Remove remaining relative typescript js suffixes ([e567a4a](https://github.com/ChristopherVR/ooxml/commit/e567a4a8a9abdea5b58d88a7dda7814e311787c9))

## [0.1.4](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.1.4) - 2026-10-07

## [0.1.3](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.1.3) - 2026-10-07

## [0.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.1.2) - 2026-10-06

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-svelte-viewer@0.1.1) - 2026-10-05

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


