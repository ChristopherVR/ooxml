# AGENTS.md

Guidance for coding agents (Claude Code, Codex, and others) working in this repository. This file is
canonical: `CLAUDE.md` only imports it, so edit this file and never fork the two.

## READ FIRST: this repository is UI, bindings and a reference server

All team-workspace **logic** lives in the `teams` area of `ooxml-core` (source in the
`ChristopherVR/ooxml` repository, `src/teams`, documented in its `docs/teams-area.md`): the chat
CRDT, presence, calls, signaling, server configuration and the `createTeamsClient` store. The
**visual primitives** (avatar, app rail, channel list, chat list, composer, pre-join, call grid and
controls) live in `ooxml-ui` (`packages/ui/src/teams`). This repository holds:

- `packages/web-component`: `<teams-app>`, the app that composes those primitives over the store.
- `packages/{react,vue,solid,svelte,angular,vanilla}`: thin bindings.
- `server/`: the reference bring-your-own server (sync, signaling relay, file storage).
- `demos/`: runnable demos.

Never add or fork logic here that belongs in the core. If the core lacks something, change it there
(the sibling `../ooxml-core` checkout), then use it here.

## How the code is written

- **Elements read like components.** Each element is a Lit class: reactive properties declared with
  `declare` and set in the constructor (class fields would shadow Lit's accessors), a `render()`
  template that reads like JSX, and its styles in a **real `.css` file** beside it, imported with
  `?raw`. No hand-built DOM trees, no CSS in template strings.
- **Colours, spacing and radii come from the ooxml-ui tokens** (`var(--office-...)`). No raw colours
  or pixel lengths in component CSS; relative units (`em`) are fine.
- **Bindings are lifecycle adapters.** The component binding creates `<teams-app>`, forwards props
  with `applyTeamsProps` and re-emits events with `listenTeamsEvents` (both in `bind.ts`). Each
  binding also exposes the **raw hook** over `createTeams` (React `useTeams`, Vue `useTeams`, Solid
  `createTeamsClient`, Svelte `teamsStore`, Angular `TeamsService`). A behaviour bug is fixed once,
  in `web-component` or the core; a hook or wiring bug is fixed in every binding in the same change.
- **State in, actions out.** UI reads `TeamsState` and calls client actions; it never touches Yjs,
  `RTCPeerConnection` or the server directly.
- Anything a peer sent is rendered as text (`textContent` semantics), never as markup.

## Commands

```bash
bun install          # ../ooxml-core must be built (bun run build there); file: dependencies are COPIES,
                     # so after changing it run `bun install --force` here (demos and tests alias to its source)
bun run dev          # reference server :8787, vanilla demo :5173, React demo :5174
bun run typecheck    # strict; React, Vue, Solid, vanilla and the web component (Svelte/Angular are source-only)
bun run test         # vitest (jsdom) then the server's node:test suite
```

## Style and commits

TypeScript strict (`exactOptionalPropertyTypes`, `noUncheckedIndexedAccess`), tabs, single quotes.
**No em-dashes** (U+2014) anywhere: source, comments, docs, commits, UI copy. Commits follow
Conventional Commits with the area as scope (`web-component`, `react`, `server`, `demo`, `docs`).
Trunk-based: commit to `main`. Do not publish packages by hand; nothing here is published yet.

## Honesty

Report unsupported features honestly. This is not Microsoft Teams and there is no end-to-end
encryption; do not write copy or docs that imply either.
