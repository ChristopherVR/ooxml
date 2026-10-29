<div align="center">

# OpenTeams

[![npm version](https://img.shields.io/npm/v/openteams-react-viewer.svg)](https://www.npmjs.com/package/openteams-react-viewer)
[![license](https://img.shields.io/npm/l/openteams-react-viewer.svg)](https://github.com/ChristopherVR/ooxml/blob/main/viewers/teams/LICENSE)
[![types](https://img.shields.io/npm/types/openteams-react-viewer.svg)](https://www.npmjs.com/package/openteams-react-viewer)

**An open-source, bring-your-own-server team workspace: channels and chat, presence and WebRTC meetings, one web component and thin adapters for your framework.**
A beta: not Microsoft Teams, not affiliated with Microsoft, and no end-to-end encryption.

[![docs](https://img.shields.io/badge/docs-christophervr.github.io-0e8f8f.svg)](https://christophervr.github.io/ooxml/teams/)
[![license](https://img.shields.io/badge/license-Apache--2.0-blue.svg)](LICENSE)
[![CI](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml/badge.svg)](https://github.com/ChristopherVR/ooxml/actions/workflows/ci.yml)

[**Live demo**](https://christophervr.github.io/ooxml/teams/demo/) &nbsp;&middot;&nbsp;
[**Documentation**](https://christophervr.github.io/ooxml/teams/) &nbsp;&middot;&nbsp;
[**Getting started**](#getting-started) &nbsp;&middot;&nbsp;
[**Packages**](#packages)

Part of [ooxml](https://github.com/ChristopherVR/ooxml#readme): [Office](https://github.com/ChristopherVR/ooxml#ooxml-office-the-whole-suite-in-one-app) &middot; [Word](https://github.com/ChristopherVR/ooxml/tree/main/viewers/docx#readme) &middot; [Excel](https://github.com/ChristopherVR/ooxml/tree/main/viewers/xlsx#readme) &middot; [PowerPoint](https://github.com/ChristopherVR/ooxml/tree/main/viewers/pptx#readme) &middot; [Visio](https://github.com/ChristopherVR/ooxml/tree/main/viewers/visio#readme) &middot; **[OpenTeams](https://github.com/ChristopherVR/ooxml/tree/main/viewers/teams#readme)** &middot; [Core](https://github.com/ChristopherVR/ooxml/tree/main/src/core#readme)

![The OpenTeams workspace: a channel list, a conversation with reactions and a quoted reply, and the composer](docs/public/hero-workspace.webp)

</div>

## Why OpenTeams?

- **Your server, your data.** No hosted service: run the small reference server (`openteams-server`) or any server that speaks three documented contracts (y-websocket sync, a JSON signaling relay, optional file storage).
- **One app, every framework.** A Lit `<teams-app>` element owns the UI. React, Vue, Angular, Svelte, Solid and vanilla bindings only handle lifecycle and events, and each also exposes a raw hook so you can build your own UI.
- **Logic in one place.** Chat, presence, calls and the client store live in [`ooxml-core/teams`](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-area.md), the visual primitives in [`ooxml-ui`](https://github.com/ChristopherVR/ooxml/tree/main/src/ui). This repository holds the app, the bindings, the reference server and the demos.
- **Honest limits.** What is not supported is written down, not hidden. See [limitations](https://christophervr.github.io/ooxml/teams/limitations).

## Features and limitations

|                           |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| ------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Channels and chat**     | Channels, posts with replies, 1:1 and group chats, reactions, author-only edit and delete, search and unread counts, in one Yjs document that merges concurrent edits. Everything a peer sends is re-validated and rendered as text, never markup.                                                                                                                                                                                                                                                                                                                          |
| **Presence**              | Who is online, typing, and a status you set (available, busy, do not disturb, be right back, away, offline), a status message and automatic away when idle. Advisory and unauthenticated, like any Yjs awareness.                                                                                                                                                                                                                                                                                                                                                           |
| **Calls**                 | Meetings in a channel over WebRTC: pre-join screen, audio, video, screen share and raise hand. Full mesh with a signaling relay, no media server: honest up to about a dozen people. TURN for strict NATs is yours to run.                                                                                                                                                                                                                                                                                                                                                  |
| **Files in conversation** | Attachments go to your server's `/files` endpoint or your own `uploadFile` function and are posted as links; Word, Excel, PowerPoint and Visio files are recognised so a host can open them. Without a file server, files are shared by name only.                                                                                                                                                                                                                                                                                                                          |
| **Server**                | Bring your own. The reference `openteams-server` (one Node process) does sync, signaling and file storage with an optional shared token, accounts with signed sessions (required unless started with `--insecure-dev`), an origin allowlist, per-room permissions (`authorizeSync`) and short-lived signed download links.                                                                                                                                                                                                                                                  |
| **Local mode**            | `mode: 'local'` needs no server: tabs of one browser share chat, presence and calls over `BroadcastChannel`. The live demos run this way.                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| **Not supported**         | No end-to-end encryption (your server can read chat and files). The reference server is secure by default: it will not start without accounts, an `authenticate` hook or `authorizeSync` (unless `--insecure-dev`), and its accounts are minimal (no SSO, password reset or team and channel access control). With `--insecure-dev` there is one shared token and user ids are claims; authorization is the server's job. Notifications are in-page only (no push or email). No recording, background blur or live captions. Markdown is a subset (no full CommonMark/GFM). |

## How it works

1. **Start a workspace.** Open the [in-browser demo](https://christophervr.github.io/ooxml/teams/demo/) (no server: it says so at the top), or run `npx openteams-server --insecure-dev` (local development only) and point the app at `ws://127.0.0.1:8787/sync` and `/signal`.

   ![The vanilla demo on GitHub Pages with its notice: this demo runs entirely in your browser](docs/public/in-browser-demo.webp)

2. **Open a second tab as someone else.** The notice's link (or `?name=Bob`) opens the same workspace as another person; in server mode, any browser that reaches your server can join.
3. **Talk.** Post in a channel, reply, react, and attach files. Each message is a Yjs update that reaches the other tabs (local mode) or every client of the room (server mode).

   ![A conversation with Word, Excel and PowerPoint attachments shown as cards](docs/public/files-in-conversation.webp)

4. **Meet.** **Meet now** opens the pre-join screen (camera and microphone toggles); **Join now** connects everyone in the channel's call peer to peer, with signaling relayed by the server (or `BroadcastChannel` locally).

   ![The pre-join screen with camera and microphone toggles](docs/public/call-prejoin.webp)

## Getting started

### 1. Install

```bash
npm install openteams-react-viewer      # or -vue-, -angular-, -svelte-, -solid-, -vanilla-viewer
npx openteams-server --insecure-dev     # local development only: 127.0.0.1:8787 (Node 22+)
```

Each binding is self-contained (the `<teams-app>` element is bundled in) and installs [`ooxml-core`](https://www.npmjs.com/package/ooxml-core) and [`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui) as regular dependencies. Use one binding per page: each registers the same custom elements.

### 2. Mount the workspace

```tsx
import { Teams } from 'openteams-react-viewer';

const config = {
	mode: 'server' as const,
	syncUrl: 'wss://teams.example.com/sync',
	signalingUrl: 'wss://teams.example.com/signal',
	iceServers: [{ urls: 'stun:stun.example.com:3478' }],
};

export function Workspace() {
	return <Teams workspaceId="acme" userName="Ada" config={config} />;
}
```

Or build your own UI over the raw hook (state in, plain actions out):

```tsx
const { client, state } = useTeams({ workspaceId: 'acme', user, config });
state?.channels.map((c) => <button onClick={() => client!.select(c.id)}>{c.name}</button>);
```

<details>
<summary><strong>Vanilla JavaScript</strong></summary>

```ts
import { mountTeams } from 'openteams-vanilla-viewer';

const teams = mountTeams(document.getElementById('teams')!, {
	workspaceId: 'acme',
	userName: 'Ada',
	config,
	onOpenFile: (detail) => console.log(detail.attachment.kind, detail.url),
});
teams.destroy();
```

Or plain markup after `defineTeamsApp()`:

```html
<teams-app
	workspace-id="acme"
	user-name="Ada"
	server-config='{"mode":"server","syncUrl":"wss://teams.example.com/sync","signalingUrl":"wss://teams.example.com/signal","iceServers":[{"urls":"stun:stun.example.com:3478"}]}'
></teams-app>
```

</details>

| Binding  | Component                       | Raw hook                                                 |
| -------- | ------------------------------- | -------------------------------------------------------- |
| React    | `<Teams />`                     | `useTeams(options)` (`useTeamsClient`, `useTeamsState`)  |
| Vue 3    | `<Teams />`                     | `useTeams(() => options)` returns shallow refs           |
| Solid    | `Teams(props)`                  | `createTeamsClient(() => options)` returns signals       |
| Svelte 5 | `Teams.svelte` (default export) | `teamsStore(options)` from `/runtime` (a readable store) |
| Angular  | `<teams-workspace>`             | `TeamsService` (signals)                                 |
| Vanilla  | `mountTeams(el, props)`         | `createTeams(options)`                                   |

See the [framework guides](https://christophervr.github.io/ooxml/teams/frameworks/react) and `demos/teams/react/src/App.tsx` for both styles side by side.

### 3. Run the reference server

```bash
TEAMS_USERS_FILE=users.json TEAMS_SESSION_SECRET=change-me TEAMS_ORIGINS=https://app.example.com TEAMS_DATA=/var/lib/openteams npx openteams-server
npx openteams-server --insecure-dev   # local development only: no accounts, listens on 127.0.0.1:8787
```

The server refuses to start without accounts (`TEAMS_USERS_FILE`), an `authenticate` hook or `authorizeSync`; `--insecure-dev` is the explicit, warned opt-out. It is configured by environment (`PORT`, `HOST`, `TEAMS_USERS_FILE`, `TEAMS_SESSION_SECRET`, `TEAMS_ORIGINS`, `TEAMS_DATA`, and more for limits; see [server/README.md](server/README.md)) and serves `ws /sync/<room>`, `ws /signal/<room>`, `POST|GET|DELETE /files/<scope>/<name>` (workspace and private-chat scopes), `GET /files/<scope>`, `POST /files/link/<scope>/<name>`, `GET /health` and, with `TEAMS_METRICS_TOKEN`, `GET /metrics`. It compacts room snapshots, enforces per-workspace file quotas and can log JSON (`--log-format=json`). Put it behind TLS. See [bring your own server](https://christophervr.github.io/ooxml/teams/server) for the contracts and what it does not do, and [docs/deploy.md](docs/deploy.md) for TLS, tokens, TURN and Docker.

## Packages

Seven packages are published on npm, each versioned independently. The names are `openteams-*` because the plain `openteams` name on npm belongs to an unrelated project.

| Package                    | What it is                                                             | Peer dependency |
| -------------------------- | ---------------------------------------------------------------------- | --------------- |
| `openteams-react-viewer`   | React 18+ component and hooks (`packages/react`).                      | `react`         |
| `openteams-vue-viewer`     | Vue 3.4+ component and composable (`packages/vue`).                    | `vue`           |
| `openteams-angular-viewer` | Angular 17+ standalone component and service (`packages/angular`).     | `@angular/core` |
| `openteams-svelte-viewer`  | Svelte 5 component and store (`packages/svelte`).                      | `svelte`        |
| `openteams-solid-viewer`   | SolidJS 1.9 component and primitive (`packages/solid`).                | `solid-js`      |
| `openteams-vanilla-viewer` | `mountTeams` and the plain `<teams-app>` element (`packages/vanilla`). | none            |
| `openteams-server`         | The reference server and its `openteams-server` command (`server`).    | none (Node 22+) |

The workspace also holds **private** packages that are never published: `teams-viewer` (`packages/web-component`, a re-export of `ooxml-ui/teams`, inlined into every binding at build time) and the demos.

```
packages/web-component   re-exports ooxml-ui/teams (the element and the binding wiring in bind.ts)
packages/{react,vue,solid,svelte,angular,vanilla}   thin bindings: a component AND raw hooks each
server/                  reference BYO server: Yjs sync + signaling relay + file storage + accounts
docs/                    the VitePress site and the GitHub Pages build (separate install)
```

`<teams-app>` itself (Lit: `teams-app.ts` + `teams-app.css`, the controller) and the binding wiring (`bindTeams`, `pickTeamsProps`, `TEAMS_CONTRACT` in `bind.ts`) are `src/ui/src/teams/app` at the repository root, shipped as `ooxml-ui/teams`. The demos are `demos/teams/*` and the browser tests `e2e/teams/`, both at the repository root.

## Development

Needs Node 22+ and Bun. This folder is `viewers/teams` of the [`ooxml`](https://github.com/ChristopherVR/ooxml) repository; run `bun install` at the repository root, then the commands below here. The demos and tests use the source of `ooxml-core` and `ooxml-ui` in the same tree.

```bash
bun install
bun run dev            # reference server :8787, vanilla demo :5173, React demo :5174
bun run typecheck      # strict; the web component and all six bindings (.svelte files excepted)
bun run test           # vitest (jsdom) then the server's node:test suite
bun run test:scripts   # this folder's script tests (published-reference and workspace-package checks)
bun run build:packages # dist/ for the six bindings (web component inlined); the server needs no build
bun run check:published && bun run pack:smoke
bun run pack:smoke --workspace-runtime   # also packs this checkout's ooxml-core and ooxml-ui (what CI runs)
bun install --cwd docs && bun run build:pages   # the docs site and both demos, as deployed
bun run docs:dev       # the docs site with hot reload
```

Open `http://127.0.0.1:5173/?name=Ada` and, in another tab, `...?name=Bob`: two people in one workspace. `?local=1` runs without the server. `TEAMS_USE_DIST=1` makes the demos use the built packages (`bun run build:packages` first) instead of the source. Cameras and microphones need `localhost` or HTTPS.

## Releasing

Releases are automated from Conventional Commits: each package has its own version and tag (`<npm-name>@<version>`), and the repository root's hourly release workflow plans, tags and publishes through npm trusted publishing. See [docs/releasing.md](docs/releasing.md) and the root [docs/releasing.md](../../docs/releasing.md).

## Documentation

[Getting started](https://christophervr.github.io/ooxml/teams/getting-started) &middot; [Architecture](https://christophervr.github.io/ooxml/teams/architecture) &middot; [Framework guides](https://christophervr.github.io/ooxml/teams/frameworks/react) &middot; [Bring your own server](https://christophervr.github.io/ooxml/teams/server) &middot; [Deploying](https://christophervr.github.io/ooxml/teams/deploy) &middot; [Limitations](https://christophervr.github.io/ooxml/teams/limitations) &middot; [Live demos](https://christophervr.github.io/ooxml/teams/demos)

## Contributing and license

Contributions follow [CONTRIBUTING.md](../../CONTRIBUTING.md). [Apache License 2.0](LICENSE); see [`NOTICE`](NOTICE) for attributions.
