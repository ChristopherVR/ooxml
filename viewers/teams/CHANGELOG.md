# Changelog

All notable changes to this viewer's packages, newest first. Each section names the packages
released that day; the full history of each package is in its own `CHANGELOG.md`
(`packages/<name>/`, `mcp/`, `server/`).

## Unreleased

<!-- release-notes: 2.0 -->

_Releases: openteams-react-viewer@2.0.0, openteams-vue-viewer@2.0.0, openteams-angular-viewer@2.0.0, openteams-svelte-viewer@2.0.0, openteams-solid-viewer@2.0.0, openteams-vanilla-viewer@2.0.0, openteams-server@2.0.0_

OpenTeams moves every package to its next major version on the same day. Most of the release hardens the reference server and the protocol against hostile peers: private chat membership is decided by the server, calls are bound to the signed-in user, and crafted updates, presence and file names can no longer stall a server or grow its storage. The client and the server now speak Teams protocol version 2 and do not work with 1.x, so deploy them together, after backing up the server's data directory, which the new server migrates on first start. The six bindings keep their component and hook APIs. OpenTeams stays in beta. The [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#openteams) gives the before and after of each step.

### Breaking changes

- Teams protocol version 2: a 2.0 server closes clients that announce another version. Deploy the app and the server together.
- New private chat ids (`dm-` and 32 hex digits, `g-` bound to its creator) and call rooms (`call-` and the chat's room name). Existing chats keep working for the people their rooms list; local changes still unsent at the upgrade are discarded.
- `openteams-server` migrates its data directory to new stored names on first start, refuses session secrets shorter than 32 characters or with fewer than 8 distinct ones, offers sync alone when it has only `authorizeSync`, and enforces new default limits on files, stored rooms, presence, sync message size and repeated refusals.
- Updates over a struct budget are refused before they are decoded, in the browser and on the server.
- Every package requires Node.js 22.3 or newer.

### Security

- Private chat membership is decided by the server, and nobody can take over a chat room by connecting first.
- Call signaling is bound to the signed-in user, and a private chat's call admits only its members.
- Small crafted messages can no longer stall the server or a browser, and storage quotas can no longer be sidestepped with new names.
- Weak session secrets are refused, and an open socket is closed when its session expires.

### Added

- Teams shows why sharing stopped when an update was refused, in the display language.
- New server limits and their environment variables (`TEAMS_USER_QUOTA_MB`, `TEAMS_ROOMS_QUOTA_MB`, `TEAMS_MAX_STORED_ROOMS`, presence caps, `TEAMS_MAX_REFUSALS_PER_MIN`, `TEAMS_MAX_UPDATE_STRUCTS`).

### Fixed

- Concurrent and offline group membership changes are all kept, and no longer stop a person's sync.
- A refusal is shown to the person instead of a bare close code, without a reconnect loop.
- The Solid binding renders during server rendering and hydrates in the browser.
- The starting channel is shown in the display language.

### Accessibility

- F6 and Shift+F6 move between the app bar, chat list, header, messages and message box.
- Search and composer boundaries reach 3:1 contrast, and selected states survive forced-colours mode.
- Keyboard and axe-core checks run in a real browser before every release.

### Packages

- [`openteams-react-viewer` 2.0.0](packages/react/CHANGELOG.md)
- [`openteams-vue-viewer` 2.0.0](packages/vue/CHANGELOG.md)
- [`openteams-angular-viewer` 2.0.0](packages/angular/CHANGELOG.md)
- [`openteams-svelte-viewer` 2.0.0](packages/svelte/CHANGELOG.md)
- [`openteams-solid-viewer` 2.0.0](packages/solid/CHANGELOG.md)
- [`openteams-vanilla-viewer` 2.0.0](packages/vanilla/CHANGELOG.md)
- [`openteams-server` 2.0.0](server/CHANGELOG.md)
- Built on [`ooxml-core` 2.0.0](../../src/core/CHANGELOG.md) and [`ooxml-ui` 2.0.0](../../src/ui/CHANGELOG.md)

## 2026-11-16

_Releases: openteams-server@1.1.3_

### Dependencies

- **deps:** Take the open dependency updates in one change ([8eab074](https://github.com/ChristopherVR/ooxml/commit/8eab0749af462b944cda96a92e723132e9031f7a))

## 2026-11-13

_Releases: openteams-react-viewer@1.1.2, openteams-angular-viewer@1.1.2, openteams-vanilla-viewer@1.1.2, openteams-server@1.1.2_

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## 2026-11-09

_Releases: openteams-server@1.1.1_

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 25 updates ([4046ca3](https://github.com/ChristopherVR/ooxml/commit/4046ca399665d485da35696889c0212cc6526f63))

## 2026-11-06

_Releases: openteams-react-viewer@1.1.1, openteams-vue-viewer@1.1.1, openteams-angular-viewer@1.1.1, openteams-svelte-viewer@1.1.1, openteams-solid-viewer@1.1.1, openteams-vanilla-viewer@1.1.1_

### Documentation

- Disclose every outbound request in PRIVACY.md ([eb1810b](https://github.com/ChristopherVR/ooxml/commit/eb1810bd3dd1ede67a8c4872f3fe5daa64472fe1))

## 2026-11-04

_Releases: openteams-react-viewer@1.1.0, openteams-vue-viewer@1.1.0, openteams-angular-viewer@1.1.0, openteams-svelte-viewer@1.1.0, openteams-solid-viewer@1.1.0, openteams-vanilla-viewer@1.1.0, openteams-server@1.1.0_

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

### Dependencies

- **deps:** Bump the angular group with 8 updates ([b233fe0](https://github.com/ChristopherVR/ooxml/commit/b233fe0bd00783876e01b5f0c7488a8f9e139fe6))
- **deps:** Merge angular dependency updates ([#234](https://github.com/ChristopherVR/ooxml/issues/234)) ([5c642cf](https://github.com/ChristopherVR/ooxml/commit/5c642cf7e6fc39265778fb062e4b13772af47a8e))
- **deps:** Bump the production-minor-and-patch group with 22 updates ([8e60b14](https://github.com/ChristopherVR/ooxml/commit/8e60b14e1d8ee3a1cb507670caa15a611632ef26))
- **deps:** Merge production dependency updates ([#236](https://github.com/ChristopherVR/ooxml/issues/236)) ([7eb137b](https://github.com/ChristopherVR/ooxml/commit/7eb137b03906ded043ce0e4adf83e01fcdaa44a9))

## 2026-10-29

_Releases: openteams-react-viewer@1.0.0, openteams-vue-viewer@1.0.0, openteams-angular-viewer@1.0.0, openteams-svelte-viewer@1.0.0, openteams-solid-viewer@1.0.0, openteams-vanilla-viewer@1.0.0, openteams-server@1.0.0_

### Features

- **teams:** Release openteams 1.0 on the versioned sync protocol ([e035417](https://github.com/ChristopherVR/ooxml/commit/e035417be1285095f9bbe1b9655743dcd059b55a))
- **teams:** Refuse to start the server without authentication ([d06e213](https://github.com/ChristopherVR/ooxml/commit/d06e213bf2eeba74e087204d768b12f96f31c070))
- **teams:** Enforce authorship, moderators and protocol in the server ([2840145](https://github.com/ChristopherVR/ooxml/commit/28401455bcc11a2342f1a5bdeb9038ca6c31bc69))
- **teams:** Validate meetings, roster, chat members and content size ([05565a0](https://github.com/ChristopherVR/ooxml/commit/05565a0dc83c0a292fa20202fa1be1b54623b7d4))
- **teams:** Scope files, compact rooms and add metrics to the server ([9c2f821](https://github.com/ChristopherVR/ooxml/commit/9c2f821d160f158176b426b9a5a2c7bf3212c083))

### Documentation

- **teams:** Document backups and hardening ([4ffdd3b](https://github.com/ChristopherVR/ooxml/commit/4ffdd3bb3ee1a5b4a5c965edbf1798352e58f352))

### Styling

- Format files from the stability work ([6ebb965](https://github.com/ChristopherVR/ooxml/commit/6ebb965ec54f2424db7297ff89a1fda5066ea280))
- Format visio evidence and teams server files ([3f5e2e0](https://github.com/ChristopherVR/ooxml/commit/3f5e2e0d46da8d8ffac7d61a088b0adb815b293a))

## 2026-10-23

_Releases: openteams-react-viewer@0.2.1, openteams-vue-viewer@0.3.1, openteams-angular-viewer@0.3.1, openteams-svelte-viewer@0.2.1, openteams-solid-viewer@0.3.1, openteams-vanilla-viewer@0.2.1, openteams-server@0.4.1_

### Features

- **server:** Add per-room sync permissions for host applications ([33643ec](https://github.com/ChristopherVR/ooxml/commit/33643ec02d475f4599f21c7ed7f6ac64cdb69acf))

### Refactor

- **teams:** Move bindTeams into ooxml-ui/teams on the binding kit ([fd417a5](https://github.com/ChristopherVR/ooxml/commit/fd417a515f76b580af69acc5b1e319d15d41fee5))

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 22 updates ([02d7713](https://github.com/ChristopherVR/ooxml/commit/02d7713ceb301921fd6c7706138d64f4c10a9d38))
- **deps:** Merge dependency update ([#94](https://github.com/ChristopherVR/ooxml/issues/94)) ([e381aba](https://github.com/ChristopherVR/ooxml/commit/e381abaa11794233cdea85a9dc6d098d4cb338c4))
- **deps:** Bump the angular group with 7 updates ([c3aee14](https://github.com/ChristopherVR/ooxml/commit/c3aee1459d6f86422e51c20819d6203343993ef0))
- **deps:** Merge angular updates ([#90](https://github.com/ChristopherVR/ooxml/issues/90)) ([119c279](https://github.com/ChristopherVR/ooxml/commit/119c279502424ffc0e14a0543c334fc828fccdd5))

### Chores

- **teams:** Commit the server's bin scripts as executable ([#180](https://github.com/ChristopherVR/ooxml/issues/180)) ([0491277](https://github.com/ChristopherVR/ooxml/commit/0491277232a67f83c75ea233a86af1a17bd30ee2))

### Other

- Shared binding kit (wave6-1) ([98ab396](https://github.com/ChristopherVR/ooxml/commit/98ab396a11b31675c0266acd5242060412642bd3))
- Bring origin/main into wave 6 ([2fedfaa](https://github.com/ChristopherVR/ooxml/commit/2fedfaae708c279044fdf86ea192315e366f80da))

## 2026-10-21

_Releases: openteams-server@0.4.0_

### Features

- **server:** Add per-room sync permissions for host applications ([d28fe46](https://github.com/ChristopherVR/ooxml/commit/d28fe46a899c9c8a7a858e87c75346a74308a532))

### Testing

- **teams:** Stop the password hash test failing on a random salt ([e9d36c9](https://github.com/ChristopherVR/ooxml/commit/e9d36c98fc2793d775f072c88cfc2f66cb173f29))

## 2026-10-16

_Releases: openteams-react-viewer@0.2.0, openteams-vue-viewer@0.3.0, openteams-angular-viewer@0.3.0, openteams-svelte-viewer@0.2.0, openteams-solid-viewer@0.3.0, openteams-vanilla-viewer@0.2.0, openteams-server@0.3.0_

### Features

- **teams:** Pass a locale through every binding and demo ([bcea500](https://github.com/ChristopherVR/ooxml/commit/bcea500df4863fc296ff5020a6b6af0f8631d072))
- **teams:** Rate limit, hash and revoke sessions on the server ([32f9921](https://github.com/ChristopherVR/ooxml/commit/32f9921325769c74493517f89a6f0ce8a9b37583))

### Testing

- **teams:** Cover chats, the calendar and reminders in the browser ([544d1f0](https://github.com/ChristopherVR/ooxml/commit/544d1f0adb31802033102e009a66c338bb0c3df6))

### Dependencies

- **deps:** Bump the angular group with 7 updates ([#57](https://github.com/ChristopherVR/ooxml/issues/57)) ([7d7881b](https://github.com/ChristopherVR/ooxml/commit/7d7881bfd9615f32b32fbd5c856402926498a492))
- **deps:** Update production dependencies ([bbd7b53](https://github.com/ChristopherVR/ooxml/commit/bbd7b530c7a15ae48c0c0a2ff4ea7dcdfe2aa9d3))

## 2026-10-15

_Releases: openteams-server@0.2.0_

### Features

- **server:** Admit only the members of a private chat to its room ([e616ed3](https://github.com/ChristopherVR/ooxml/commit/e616ed3287fc3ced6833d139125d85c47c6b5557))
- **teams:** Authenticate users with signed sessions on the server ([8450cb0](https://github.com/ChristopherVR/ooxml/commit/8450cb0cff962c9a1b625fbab980dc6290eb33ef))

### Bug Fixes

- **teams:** Harden the reference server against abuse ([45ea3f7](https://github.com/ChristopherVR/ooxml/commit/45ea3f782957e85e323d0844c179393a5f4080f7))

### Documentation

- **teams:** Document private chats, the calendar and their limits ([02aa753](https://github.com/ChristopherVR/ooxml/commit/02aa75378403f2a8e537276c37f7a539091bfd8a))

## 2026-10-11

_Releases: openteams-react-viewer@0.1.8, openteams-vue-viewer@0.2.4, openteams-angular-viewer@0.2.4, openteams-svelte-viewer@0.1.8, openteams-solid-viewer@0.2.4, openteams-vanilla-viewer@0.1.8, openteams-server@0.1.3_

### Styling

- Fix the oxfmt and oxlint warnings in the viewers ([db0314f](https://github.com/ChristopherVR/ooxml/commit/db0314f319e3d0508960872260792215477d9ca5))

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 5 updates ([#42](https://github.com/ChristopherVR/ooxml/issues/42)) ([fda493e](https://github.com/ChristopherVR/ooxml/commit/fda493e2697908c477ce845019d9c2fe8b591ab2))

## 2026-10-08

_Releases: openteams-react-viewer@0.1.7, openteams-react-viewer@0.1.6, openteams-vue-viewer@0.2.3, openteams-vue-viewer@0.2.2, openteams-angular-viewer@0.2.3, openteams-angular-viewer@0.2.2, openteams-svelte-viewer@0.1.7, openteams-svelte-viewer@0.1.6, openteams-solid-viewer@0.2.3, openteams-solid-viewer@0.2.2, openteams-vanilla-viewer@0.1.7, openteams-vanilla-viewer@0.1.6, openteams-server@0.1.2_

### Bug Fixes

- **teams:** Accept a host class in every binding ([629bb11](https://github.com/ChristopherVR/ooxml/commit/629bb118e6adc0ad21c78a34a60c884586b09869))

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## 2026-10-07

_Releases: openteams-react-viewer@0.1.5, openteams-react-viewer@0.1.4, openteams-react-viewer@0.1.3, openteams-vue-viewer@0.2.1, openteams-vue-viewer@0.2.0, openteams-vue-viewer@0.1.3, openteams-angular-viewer@0.2.1, openteams-angular-viewer@0.2.0, openteams-angular-viewer@0.1.3, openteams-svelte-viewer@0.1.5, openteams-svelte-viewer@0.1.4, openteams-svelte-viewer@0.1.3, openteams-solid-viewer@0.2.1, openteams-solid-viewer@0.2.0, openteams-solid-viewer@0.1.3, openteams-vanilla-viewer@0.1.5, openteams-vanilla-viewer@0.1.4, openteams-vanilla-viewer@0.1.3_

### Features

- **teams:** Add embedded content previews ([c127224](https://github.com/ChristopherVR/ooxml/commit/c127224861720dd2d37e24390f940e642588b559))

### Refactor

- Remove remaining relative typescript js suffixes ([e567a4a](https://github.com/ChristopherVR/ooxml/commit/e567a4a8a9abdea5b58d88a7dda7814e311787c9))

## 2026-10-06

_Releases: openteams-react-viewer@0.1.2, openteams-vue-viewer@0.1.2, openteams-angular-viewer@0.1.2, openteams-svelte-viewer@0.1.2, openteams-solid-viewer@0.1.2, openteams-vanilla-viewer@0.1.2_

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## 2026-10-05

_Releases: openteams-react-viewer@0.1.1, openteams-vue-viewer@0.1.1, openteams-angular-viewer@0.1.1, openteams-svelte-viewer@0.1.1, openteams-solid-viewer@0.1.1, openteams-vanilla-viewer@0.1.1, openteams-server@0.1.1_

### Features

- **web-component:** Add the team-workspace app, bindings and reference server (by @ChristopherVR) ([36916ee](https://github.com/ChristopherVR/ooxml/commit/36916eeef144e9260144b5f0e87ddee6c8021b75))
- **release:** Package the bindings as openteams-*-viewer (by @ChristopherVR) ([2d5101e](https://github.com/ChristopherVR/ooxml/commit/2d5101e55f477834b591962409a37b078960dd1a))
- **server:** Ship the reference server as openteams-server (by @ChristopherVR) ([9b49f4d](https://github.com/ChristopherVR/ooxml/commit/9b49f4db57f60fd2c7a2979b32caa2e6fa0010d2))

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

### Refactor

- **web-component:** Take <teams-app> from ooxml-ui/teams and add oxlint (by @ChristopherVR) ([c8175e9](https://github.com/ChristopherVR/ooxml/commit/c8175e98a795206693c247a499112fb9ee079d15))

### Build & CI

- **deps:** Update all dependencies to latest (by @ChristopherVR) ([e49671d](https://github.com/ChristopherVR/ooxml/commit/e49671d9f2e360b85136b9604c7adaa191b5b288))
- **release:** Release the imported viewers from the root planner (by @ChristopherVR) ([86984f1](https://github.com/ChristopherVR/ooxml/commit/86984f1bbc4f4df828e8f1760586a3f0b7321b0b))
- **pages:** Serve the launcher and every viewer's docs and demos from one site (by @ChristopherVR) ([1048feb](https://github.com/ChristopherVR/ooxml/commit/1048febca7d914d11ca08a0c7a6b4c80622ca867))
- Fix the viewer checks that failed on the first run (by @ChristopherVR) ([2445d3a](https://github.com/ChristopherVR/ooxml/commit/2445d3a58688ea0917b15b48088afa03fc3562b7))

### Chores

- **viewers:** Import the teams-viewer repository with its history (by @ChristopherVR) ([7d29cec](https://github.com/ChristopherVR/ooxml/commit/7d29cecf88ac551cc5a80d2e9e35c65e130ab0de))
