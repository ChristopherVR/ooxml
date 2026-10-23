# AGENTS.md

Guidance for coding agents (Claude Code, Codex, and others) working in this repository. This file is
canonical: `CLAUDE.md` only imports it, so edit this file and never fork the two.

## READ FIRST: this repository is UI, bindings and a reference server

All team-workspace **logic** lives in the `teams` area of `ooxml-core` (source in the
`ChristopherVR/ooxml` repository, `src/core/teams`, documented in its `docs/teams-area.md`): the chat
CRDT, presence, calls, signaling, server configuration and the `createTeamsClient` store. The
**visual primitives** (avatar, app rail, channel list, chat list, composer, pre-join, call grid and
controls) live in `ooxml-ui` (`src/ui/src/teams`). This repository holds:

- `packages/web-component`: a re-export of `ooxml-ui/teams`, where `<teams-app>` (the app that composes those primitives over the store) now lives
  (the private workspace package `teams-viewer`, never published, inlined into every binding).
- `packages/{react,vue,solid,svelte,angular,vanilla}`: thin bindings, published as
  `openteams-<framework>-viewer`.
- `server/`: the reference bring-your-own server (sync, signaling relay, file storage), published as
  `openteams-server` (command `openteams-server`).
- `demos/teams/` at the repository root: runnable demos (private workspace packages).
- `e2e/teams/` at the repository root: the browser tests (`bun run test:browser`, Playwright config in this folder, `playwright.config.ts`, whose `testDir` points there), run against the vanilla demo in local mode. Add specs there.

The product name is **OpenTeams**. The plain `openteams` npm name belongs to an unrelated project,
so never publish or document that name.

Never add or fork logic here that belongs in the core. If the core lacks something, change it there
in this repository (`src/core/teams`), then use it here: this folder is part of the `ooxml` monorepo, so
it is the same tree.

## How the code is written

- **Elements read like components.** Each element is a Lit class: reactive properties declared with
  `declare` and set in the constructor (class fields would shadow Lit's accessors), a `render()`
  template that reads like JSX, and its styles in a **real `.css` file** beside it, imported with
  `?raw`. No hand-built DOM trees, no CSS in template strings.
- **Colours, spacing and radii come from the ooxml-ui tokens** (`var(--office-...)`). No raw colours
  or pixel lengths in component CSS; relative units (`em`) are fine.
- **Bindings are lifecycle adapters.** The component binding creates `<teams-app>`, forwards props
  with `bindTeams`/`pickTeamsProps` (`src/ui/src/teams/app/bind.ts` in ooxml-ui, built on `TEAMS_CONTRACT` and the shared `ooxml-ui/bindings` kit) and re-emits events with it; `bindings-contract.test.ts` checks every binding against the contract. Each
  binding also exposes the **raw hook** over `createTeams` (React `useTeams`, Vue `useTeams`, Solid
  `createTeamsClient`, Svelte `teamsStore`, Angular `TeamsService`). A behaviour bug is fixed once,
  in `web-component` or the core; a hook or wiring bug is fixed in every binding in the same change.
- **State in, actions out.** UI reads `TeamsState` and calls client actions; it never touches Yjs,
  `RTCPeerConnection` or the server directly.
- Anything a peer sent is rendered as text (`textContent` semantics), never as markup.

## Commands

```bash
bun install            # run at the repository root; ooxml-core and ooxml-ui are workspace ranges (`*`),
                       # so tests and demos use the source in this tree (TEAMS_USE_DIST=1 uses dist)
bun run dev            # reference server :8787, vanilla demo :5173, React demo :5174
bun run typecheck      # strict; the web component and all six bindings (.svelte files excepted)
bun run test           # vitest (jsdom) then the server's node:test suite
bun run test:scripts   # this folder's script tests (published-reference and workspace-package checks)
bun run build:packages # dist/ for the six bindings (web component inlined); the server needs no build
bun run check:published && bun run pack:smoke   # what the tarballs import; install and import them
```

## Releases

Seven packages are published, each with its own version and tag `<npm-name>@<version>`:
`openteams-{react,vue,angular,svelte,solid,vanilla}-viewer` and `openteams-server`. Releases are
run by the repository root's flow (`scripts/release-plan.mjs`, `scripts/viewer-packages.mjs`,
`.github/workflows/release.yml`; this folder's `.github/` is inert), which plans from Conventional
Commits, writes the changelogs, tags and publishes through npm trusted publishing. This folder has
no `release:plan` or `publish:local`. Details: `docs/releasing.md` here and the root
`docs/releasing.md`.

Never run `npm publish` by hand, never bump versions or edit `CHANGELOG.md` files by hand (the
release commit does), and never add `file:` or `link:` ranges, or a `teams-viewer` dependency, to a
published manifest (`check:published` and `pack:smoke` fail on them; the `*` ranges on `ooxml-core` and
`ooxml-ui` are rewritten by the release flow).

## Style and commits

TypeScript strict (`exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`), tabs, single quotes.
**No em-dashes** (U+2014) anywhere: source, comments, docs, commits, UI copy. Commits follow
Conventional Commits with the area as scope (`web-component`, `react`, `server`, `demo`, `docs`,
`ci`, `release`, `deps`); the type sets the version bump, so see `CONTRIBUTING.md`. Trunk-based:
commit to `main`.

## Honesty

Report unsupported features honestly. This is not Microsoft Teams and there is no end-to-end
encryption; do not write copy or docs that imply either.
