# Changelog

All notable changes to this package, newest first. Major releases are written by hand; the other
entries are generated from [Conventional Commits](https://www.conventionalcommits.org).

## [2.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@2.0.0) - Unreleased

<!-- release-notes: 2.0 -->

This release closes a set of security defects found in a review of the server, and nearly every fix changes what an operator configures or stores. The server speaks Teams protocol version 2, renames the data directory on its first start, refuses weak session secrets, closes file storage and call signaling when it cannot tell who is asking, and enforces new default limits on files, stored rooms, presence and repeated refusals. To upgrade: back up the data directory, check the session secret, the authentication hooks and the limits below, then deploy the server and the OpenTeams 2.0 app together. Upgrade together with the packages under "Released together", and follow the [migration guide](https://github.com/ChristopherVR/ooxml/blob/main/docs/migration-2.0.md#openteams) for the before and after of each step.

### Breaking changes

- **Teams protocol version 2.** A client that announces another version is accepted at the handshake and then closed with 1008 and a reason that names both versions (`unsupported protocol version 1: this server speaks 2`). Deploy the 2.0 OpenTeams app and this server together.
- **The data directory is migrated on first start.** Rooms, file scopes and uploader records are stored under `=` and the lower-case base32 of their exact name, so names that differ only by letter case stay apart. The new server renames what an older one stored and continues on the next start if it is interrupted; an older server cannot read the result. Copy the data directory before the first start (`cp -r data data-backup-1.x`). Existing chats keep working for the people their rooms list, and 1:1 chats keep their history. The migration cannot tell who should be in an existing room, so check the member list of any private chat you suspect someone joined uninvited. Data stored under a chat id that only this release issues cannot be genuine: it moves to a `.quarantine` folder inside `rooms`, `files` or `file-meta`, each entry is logged, and the server never reads it. Review or delete those folders.
- **Weak session secrets stop the server.** It refuses to start when `TEAMS_SESSION_SECRET` (`sessionSecret`) has fewer than 32 characters or fewer than 8 distinct characters. Generate one with `openssl rand -hex 32` and keep it outside the repository. Leaving it unset still gives a random secret per process.
- **A server with only `authorizeSync` offers sync alone.** With an `authorizeSync` hook and no `users` or `authenticate` hook, `POST /auth/session`, every `/files/...` request and every `/signal/...` socket answer 404, because the server cannot tell who a request comes from. If your clients use file storage or calls, add an `authenticate` hook. For development only, pass `insecureDev: true` or `--insecure-dev`.
- **New default limits.** Files have a ceiling of 1 GiB per uploader across every scope (`TEAMS_USER_QUOTA_MB`) and 20 GiB for the whole directory (`TEAMS_FILES_QUOTA_MB`, off before). Stored rooms have a ceiling of 10 GiB (`TEAMS_ROOMS_QUOTA_MB`) and 10,000 rooms (`TEAMS_MAX_STORED_ROOMS`); a socket past either is closed with `1009 server storage full`. Presence is capped at 64 KiB per client state, 4 MiB per room and 256 KiB per second per socket; a socket past a presence limit is closed with `1009 presence too large`. Large installations should raise the limits they need before upgrading.
- **Repeated costly refusals are turned away.** A client with more than 30 costly refusals in a minute (a chat join or chat file request refused after a room was read from disk, or a refused update) gets 429 for its sync sockets and chat file requests until the minute is over (`TEAMS_MAX_REFUSALS_PER_MIN`, `maxRefusals`).
- **The largest sync message is sized from the heap.** It defaults to 32 MiB or a 64th of the heap, whichever is smaller (8 MiB on a 512 MiB heap). A server with a small heap that hosts very large documents needs a larger heap (`NODE_OPTIONS=--max-old-space-size=...`) or an explicit `maxSyncBytes` to keep accepting their first sync; applying a dense first sync takes up to about 35 MiB of memory per MiB of message.
- **Updates have a struct budget.** An update that holds more than one struct per 4 bytes, or more than one deleted range per 2 bytes, is refused before it is decoded and its socket closed. The first sync of a real document stays under half of that budget. To set an absolute limit instead, use `maxUpdateStructs` (`TEAMS_MAX_UPDATE_STRUCTS`).
- **A sync message that cannot be read or applied closes its socket** with `1008 update rejected: invalid-update`. It used to be logged and dropped.
- **Call signaling is bound to the signed-in user.** With accounts, a signal whose `from` is not the signed-in user closes the socket with `1008 peer id is not the signed-in user`, and a second socket of the same user in one call closes with `1008 peer id in use`. A private chat's call uses the signaling room `call-` followed by the chat's room name, admits only the chat's members, and the `call-<workspace>-<chat id>` form that earlier clients used for a current chat id is refused (403).
- **New private chat rooms are seeded by the server.** `?members=` no longer sets the members of a new room; it is read only for a 1:1 room, to check its two people against the chat id. A group room is created only for the person its id is bound to, who adds the others through checked updates.
- **With accounts, an upload to a workspace scope that names no existing room** is answered with 404.
- **Node.js 22.3 or newer.** The package declares `"engines": { "node": ">=22.3" }` (Node 22 before).

### Security

- A server with only an `authorizeSync` hook signed a session for any claimed user and left file storage and call signaling open to anyone. Those routes are now closed on such a server.
- A signed-in user could join any call, speak in a call under another person's id and receive signals meant for someone else. Call sockets are now bound to their signed-in user, addressed signals reach only their peer, and a private chat's call admits only its members.
- Private chat membership is decided by the server. Editing the chat directory no longer gets anyone into a chat, only a chat's members can change its row, a chat's kind and creator are fixed, a member can no longer remove someone else, and nobody can take over a chat room by connecting first. A room that has been emptied stays closed.
- Several kinds of small message could make the server do work in proportion to a whole room, hold hundreds of megabytes of presence, or allocate millions of objects, stalling it for everyone. Update checks now cost the update, refused joins are answered from a member list kept beside each room, presence is capped, over-budget updates are refused before they are decoded, and repeated costly refusals are turned away.
- Storage quotas could be sidestepped by choosing new workspace and room names. Uploads now also count per uploader and against a directory ceiling, need an existing workspace when accounts are on, stored rooms have ceilings, and a room nobody wrote to is not stored.
- Room and file scope names that differ only by letter case are kept apart on every disk, and data prepared under the name of a future private chat does not become that chat when the server is upgraded.
- A session secret copied from an example let anyone sign sessions for any user and role. Short and repetitive secrets are now refused, and the documentation shows how to generate one.
- An open socket kept its write access after its session expired. It is now closed with `1008 session expired` at expiry.
- The server applies an update and the corrections it needs as one step. If checking or applying fails, nothing is stored or sent to anyone and the sender's connection closes, and nothing a connection sends after a refusal is applied.
- The package is released by a pipeline in which no third-party code can push, tag or publish, and the signed tarball attached to each GitHub release is checked against the integrity the npm registry records.

### Added

- Options `userQuotaBytes`, `roomsQuotaBytes`, `maxStoredRooms`, `maxAwarenessBytes`, `maxRoomAwarenessBytes`, `awarenessByteRate`, `maxRefusals` and `maxUpdateStructs`, with the environment variables `TEAMS_USER_QUOTA_MB`, `TEAMS_ROOMS_QUOTA_MB`, `TEAMS_MAX_STORED_ROOMS`, `TEAMS_MAX_AWARENESS_KB`, `TEAMS_MAX_ROOM_AWARENESS_KB`, `TEAMS_AWARENESS_RATE_KB`, `TEAMS_MAX_REFUSALS_PER_MIN` and `TEAMS_MAX_UPDATE_STRUCTS` ([reference](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/server/README.md#configuration)).

### Changed

- The supported Node.js versions and browsers follow the rest of the repository: Node.js 22.3 or newer, and the last two major versions of Chrome, Edge, Firefox and Safari for the clients.

### Fixed

- People added to or leaving a group at the same time are all kept, on every client, and these changes can no longer stop a person's workspace sync.
- A change made offline by someone who has since been removed from a group is undone by the server instead of stopping that person's sync.
- Leaving a group while offline takes effect in the chat when the app next connects, even if it was closed in between. The server used to keep admitting that person.
- The update budget no longer refuses the first sync of a large legitimate document. It refused a shared workbook of about 500,000 mixed cells before.
- A refusal no longer reaches a Teams user as a bare close code, and the client no longer reconnects in a loop after one.

### Performance

- Checking an update costs time in proportion to the update instead of the room. Each loaded room keeps one copy of its document that updates are tried on, so plan for up to two copies of each loaded room in memory.
- A refused chat join or chat file request for a room that is not loaded is answered from a small member list kept beside the room's snapshot (`rooms/=<name>.members.json`) instead of reading the room from disk.

### Released together

- [`ooxml-core` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/src/core/CHANGELOG.md): `ooxml-core/teams` and `ooxml-core/collab`, whose update checks, protocol constants and struct budget this server bundles; read it if you run your own server on `validateTeamsUpdate`.
- [`openteams-react-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/react/CHANGELOG.md): the React client that speaks protocol 2.
- [`openteams-vue-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/vue/CHANGELOG.md): the Vue client that speaks protocol 2.
- [`openteams-angular-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/angular/CHANGELOG.md): the Angular client that speaks protocol 2.
- [`openteams-svelte-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/svelte/CHANGELOG.md): the Svelte client that speaks protocol 2.
- [`openteams-solid-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/solid/CHANGELOG.md): the Solid client that speaks protocol 2.
- [`openteams-vanilla-viewer` 2.0.0](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/packages/vanilla/CHANGELOG.md): the framework-free client and `<teams-app>` element that speak protocol 2.

## [1.1.3](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@1.1.3) - 2026-11-16

### Dependencies

- **deps:** Take the open dependency updates in one change ([8eab074](https://github.com/ChristopherVR/ooxml/commit/8eab0749af462b944cda96a92e723132e9031f7a))

## [1.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@1.1.2) - 2026-11-13

### Documentation

- Correct outdated statements across the repository ([6ea98f5](https://github.com/ChristopherVR/ooxml/commit/6ea98f59c35894f7c2dc3e1037d82b17a482a80c))

## [1.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@1.1.1) - 2026-11-09

### Dependencies

- **deps:** Bump the production-minor-and-patch group with 25 updates ([4046ca3](https://github.com/ChristopherVR/ooxml/commit/4046ca399665d485da35696889c0212cc6526f63))

## [1.1.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@1.1.0) - 2026-11-04

### Features

- **embedding:** Connect office products to shared onedrive storage ([b8620c1](https://github.com/ChristopherVR/ooxml/commit/b8620c13a4903aa27fed365d113153cad76ed29f))

## [1.0.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@1.0.0) - 2026-10-29

### Features

- **teams:** Refuse to start the server without authentication ([d06e213](https://github.com/ChristopherVR/ooxml/commit/d06e213bf2eeba74e087204d768b12f96f31c070))
- **teams:** Enforce authorship, moderators and protocol in the server ([2840145](https://github.com/ChristopherVR/ooxml/commit/28401455bcc11a2342f1a5bdeb9038ca6c31bc69))
- **teams:** Validate meetings, roster, chat members and content size ([05565a0](https://github.com/ChristopherVR/ooxml/commit/05565a0dc83c0a292fa20202fa1be1b54623b7d4))
- **teams:** Scope files, compact rooms and add metrics to the server ([9c2f821](https://github.com/ChristopherVR/ooxml/commit/9c2f821d160f158176b426b9a5a2c7bf3212c083))
- **teams:** Release openteams 1.0 on the versioned sync protocol ([e035417](https://github.com/ChristopherVR/ooxml/commit/e035417be1285095f9bbe1b9655743dcd059b55a))

### Documentation

- **teams:** Document backups and hardening ([4ffdd3b](https://github.com/ChristopherVR/ooxml/commit/4ffdd3bb3ee1a5b4a5c965edbf1798352e58f352))

### Styling

- Format files from the stability work ([6ebb965](https://github.com/ChristopherVR/ooxml/commit/6ebb965ec54f2424db7297ff89a1fda5066ea280))
- Format visio evidence and teams server files ([3f5e2e0](https://github.com/ChristopherVR/ooxml/commit/3f5e2e0d46da8d8ffac7d61a088b0adb815b293a))

## [0.4.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@0.4.1) - 2026-10-23

### Features

- **server:** Add per-room sync permissions for host applications ([33643ec](https://github.com/ChristopherVR/ooxml/commit/33643ec02d475f4599f21c7ed7f6ac64cdb69acf))

### Other

- Bring origin/main into wave 6 ([2fedfaa](https://github.com/ChristopherVR/ooxml/commit/2fedfaae708c279044fdf86ea192315e366f80da))

### Chores

- **teams:** Commit the server's bin scripts as executable ([#180](https://github.com/ChristopherVR/ooxml/issues/180)) ([0491277](https://github.com/ChristopherVR/ooxml/commit/0491277232a67f83c75ea233a86af1a17bd30ee2))

## [0.4.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@0.4.0) - 2026-10-21

### Features

- **server:** Add per-room sync permissions for host applications ([d28fe46](https://github.com/ChristopherVR/ooxml/commit/d28fe46a899c9c8a7a858e87c75346a74308a532))

### Testing

- **teams:** Stop the password hash test failing on a random salt ([e9d36c9](https://github.com/ChristopherVR/ooxml/commit/e9d36c98fc2793d775f072c88cfc2f66cb173f29))

## [0.3.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@0.3.0) - 2026-10-16

### Features

- **teams:** Rate limit, hash and revoke sessions on the server ([32f9921](https://github.com/ChristopherVR/ooxml/commit/32f9921325769c74493517f89a6f0ce8a9b37583))

## [0.2.0](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@0.2.0) - 2026-10-15

### Features

- **server:** Admit only the members of a private chat to its room ([e616ed3](https://github.com/ChristopherVR/ooxml/commit/e616ed3287fc3ced6833d139125d85c47c6b5557))
- **teams:** Authenticate users with signed sessions on the server ([8450cb0](https://github.com/ChristopherVR/ooxml/commit/8450cb0cff962c9a1b625fbab980dc6290eb33ef))

### Bug Fixes

- **teams:** Harden the reference server against abuse ([45ea3f7](https://github.com/ChristopherVR/ooxml/commit/45ea3f782957e85e323d0844c179393a5f4080f7))

### Documentation

- **teams:** Document private chats, the calendar and their limits ([02aa753](https://github.com/ChristopherVR/ooxml/commit/02aa75378403f2a8e537276c37f7a539091bfd8a))

## [0.1.3](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@0.1.3) - 2026-10-11

### Styling

- Fix the oxfmt and oxlint warnings in the viewers ([db0314f](https://github.com/ChristopherVR/ooxml/commit/db0314f319e3d0508960872260792215477d9ca5))

## [0.1.2](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@0.1.2) - 2026-10-08

### Refactor

- Move the library into src/core, next to src/ui ([3933a88](https://github.com/ChristopherVR/ooxml/commit/3933a88e36784ce5f4bfab0af1ca12eb65d3cd0c))

### Documentation

- Refresh the readmes with webp images and the office suite ([6a800d0](https://github.com/ChristopherVR/ooxml/commit/6a800d062bc911d6c4ea682bdbffa7fc437d1556))

### Chores

- Merge the release commit into the restructure ([efdb30a](https://github.com/ChristopherVR/ooxml/commit/efdb30a34873f71211e11a39f315b5dc3c3cd82c))

## [0.1.1](https://github.com/ChristopherVR/ooxml/releases/tag/openteams-server@0.1.1) - 2026-10-05

### Features

- **web-component:** Add the team-workspace app, bindings and reference server (by @ChristopherVR) ([36916ee](https://github.com/ChristopherVR/ooxml/commit/36916eeef144e9260144b5f0e87ddee6c8021b75))
- **server:** Ship the reference server as openteams-server (by @ChristopherVR) ([9b49f4d](https://github.com/ChristopherVR/ooxml/commit/9b49f4db57f60fd2c7a2979b32caa2e6fa0010d2))

### Bug Fixes

- **viewers:** Point every package's homepage, bugs and README links at ooxml (by @ChristopherVR) ([e40592e](https://github.com/ChristopherVR/ooxml/commit/e40592e1cc1331c8223e8fcc366041a933435dee))

### Build & CI

- **deps:** Update all dependencies to latest (by @ChristopherVR) ([e49671d](https://github.com/ChristopherVR/ooxml/commit/e49671d9f2e360b85136b9604c7adaa191b5b288))
- **pages:** Serve the launcher and every viewer's docs and demos from one site (by @ChristopherVR) ([1048feb](https://github.com/ChristopherVR/ooxml/commit/1048febca7d914d11ca08a0c7a6b4c80622ca867))
- Fix the viewer checks that failed on the first run (by @ChristopherVR) ([2445d3a](https://github.com/ChristopherVR/ooxml/commit/2445d3a58688ea0917b15b48088afa03fc3562b7))

### Chores

- **viewers:** Import the teams-viewer repository with its history (by @ChristopherVR) ([7d29cec](https://github.com/ChristopherVR/ooxml/commit/7d29cecf88ac551cc5a80d2e9e35c65e130ab0de))

