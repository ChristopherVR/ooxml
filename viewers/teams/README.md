# OpenTeams (teams-viewer)

**OpenTeams** is an open-source, bring-your-own-server **team workspace** that sits beside the Open Office suite
(Word, Excel, PowerPoint, Visio): channels and chat, presence, meetings with video and screen
share (WebRTC), and Office files shared in the conversation. The logic lives in
[`ooxml-core/teams`](https://github.com/ChristopherVR/ooxml/blob/main/docs/teams-area.md), the visual
pieces in [`ooxml-ui`](https://github.com/ChristopherVR/ooxml/tree/main/packages/ui); this repository
holds the **app, the framework bindings, the reference server and the demos**.

```
packages/web-component   <teams-app> (Lit: teams-app.ts + teams-app.css), the controller, the raw store
packages/{react,vue,solid,svelte,angular,vanilla}   thin bindings: a component AND raw hooks each
server/                  reference BYO server: Yjs sync + signaling relay + file storage (~300 lines)
demos/{vanilla,react}    run against your local ooxml checkout, hot reload
```

## Install

Pick the binding for your framework; each one is self-contained (the `<teams-app>` element is
bundled in) and installs [`ooxml-core`](https://www.npmjs.com/package/ooxml-core) and
[`ooxml-ui`](https://www.npmjs.com/package/ooxml-ui) as regular dependencies.

| Package                    | For                         | Peer dependency       |
| -------------------------- | --------------------------- | --------------------- |
| `openteams-react-viewer`   | React 18+                   | `react`               |
| `openteams-vue-viewer`     | Vue 3.4+                    | `vue`                 |
| `openteams-angular-viewer` | Angular 17+ (standalone)    | `@angular/core`       |
| `openteams-svelte-viewer`  | Svelte 5                    | `svelte`              |
| `openteams-solid-viewer`   | SolidJS 1.9                 | `solid-js`            |
| `openteams-vanilla-viewer` | no framework (`<teams-app>`) | none                 |
| `openteams-server`         | the reference server (Node 22+) | none              |

```bash
npm install openteams-react-viewer      # or -vue-, -angular-, -svelte-, -solid-, -vanilla-viewer
npx openteams-server                    # the reference server on 127.0.0.1:8787
```

Use one binding per page: each registers the same custom elements. The package names are
`openteams-*` because the plain `openteams` name on npm belongs to an unrelated project.

## Run it from source

Needs Node 22+ and Bun. With the [`ooxml`](https://github.com/ChristopherVR/ooxml) repository checked
out next to this one (`../ooxml-core`), the demos and tests use its source; without it they use the
published `ooxml-core` and `ooxml-ui`.

```bash
bun install
bun run dev        # server on :8787, vanilla demo on :5173, React demo on :5174
```

Open `http://127.0.0.1:5173/?name=Ada` and, in another tab, `...?name=Bob`: two people in one
workspace. `?local=1` runs without the server (tabs of one browser share state over
`BroadcastChannel`). `TEAMS_USE_DIST=1` makes the demos use the built packages
(`bun run build:packages` first) and the published ooxml packages instead of any source. Cameras and
microphones need `localhost` or HTTPS.

## Use it

One element, full UI (`openteams-vanilla-viewer` exports `defineTeamsApp()`, which registers it):

```html
<teams-app workspace-id="acme" user-name="Ada" server-config='{"mode":"server","syncUrl":"wss://teams.example.com/sync","signalingUrl":"wss://teams.example.com/signal","iceServers":[{"urls":"stun:stun.example.com:3478"}]}'></teams-app>
```

Each binding gives you the same component **and** raw hooks over the core client, so you can build
your own UI instead (state in, plain actions out):

| Binding     | Component                    | Raw hook                                           |
| ----------- | ---------------------------- | -------------------------------------------------- |
| React       | `<Teams />`                  | `useTeams(options)` (`useTeamsClient`, `useTeamsState`) |
| Vue 3       | `<Teams />`                  | `useTeams(() => options)` returns shallow refs     |
| Solid       | `Teams(props)`               | `createTeamsClient(() => options)` returns signals |
| Svelte 5    | `Teams.svelte` (default export) | `teamsStore(options)` from `/runtime` (a readable store) |
| Angular     | `<teams-workspace>`          | `TeamsService` (signals)                           |
| Vanilla     | `mountTeams(el, props)`      | `createTeams(options)`                             |

```tsx
const { client, state } = useTeams({ workspaceId: 'acme', user, config });
state?.channels.map((c) => <button onClick={() => client!.select(c.id)}>{c.name}</button>);
```

See `demos/react/src/App.tsx` for both styles side by side.

## Bring your own server

The workspace needs no server at all (`mode: "local"`), or any server that speaks three small
contracts: y-websocket for sync, a JSON relay for call signaling, and optional file storage. The
reference implementation in `server/` (published as `openteams-server`) does all three; see
[docs/deploy.md](docs/deploy.md) for running it, TLS, tokens, and TURN for calls across NATs.

## Honest limits

No end-to-end encryption (your server can read chat), mesh calls scale to about a dozen people,
authorization is the server's job, and this is not Microsoft Teams: it implements channels, posts
with replies and reactions, files, search, presence and meetings, not 1:1 chats, notifications,
calendar or recording. Details in the core's `docs/teams-area.md`.

## Develop

```bash
bun run typecheck && bun run test   # vitest (jsdom) and the server's node:test suite
bun run build:packages && bun run pack:smoke
```

The repository's `tsc` checks the web component and all six bindings except the `.svelte` file,
which the Svelte toolchain compiles in a consuming app (`pack:smoke` compiles it too). Releases are
automated from Conventional Commits: see [CONTRIBUTING.md](CONTRIBUTING.md) and
[docs/releasing.md](docs/releasing.md).

Licence: Apache-2.0.
