# Changelog

All notable changes to this project are documented here.
This file is generated from [Conventional Commits](https://www.conventionalcommits.org)
by [git-cliff](https://git-cliff.org); do not edit it by hand.
A release listed with no entries carried no Conventional Commit in this package's
scope: scripts/release-plan.mjs releases whenever any published file changes, not only on
conventional commits.

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


