# teams-viewer

An open-source, bring-your-own-server **team workspace** that sits beside the Open Office suite
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

## Run it

Needs Node 22+ and Bun, with the [`ooxml`](https://github.com/ChristopherVR/ooxml) repository checked
out next to this one (`../ooxml-core`) and built once (`bun run build` there).

```bash
bun install
bun run dev        # server on :8787, vanilla demo on :5173, React demo on :5174
```

Open `http://127.0.0.1:5173/?name=Ada` and, in another tab, `...?name=Bob`: two people in one
workspace. `?local=1` runs without the server (tabs of one browser share state over
`BroadcastChannel`). `TEAMS_USE_DIST=1` makes the demos use the built packages in `node_modules`
instead of the sibling checkout's source. Cameras and microphones need `localhost` or HTTPS.

## Use it

One element, full UI:

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
| Svelte 5    | `Teams.svelte`               | `teamsStore(options)` (a readable store)           |
| Angular     | `<teams-workspace>`          | `TeamsService` (signals)                           |
| Vanilla     | `mountTeams(el, props)`      | `createTeams(options)` from `teams-viewer`         |

```tsx
const { client, state } = useTeams({ workspaceId: 'acme', user, config });
state?.channels.map((c) => <button onClick={() => client!.select(c.id)}>{c.name}</button>);
```

See `demos/react/src/App.tsx` for both styles side by side.

## Bring your own server

The workspace needs no server at all (`mode: "local"`), or any server that speaks three small
contracts: y-websocket for sync, a JSON relay for call signaling, and optional file storage. The
reference implementation in `server/` does all three; see
[docs/deploy.md](docs/deploy.md) for running it, TLS, tokens, and TURN for calls across NATs.

## Honest limits

No end-to-end encryption (your server can read chat), mesh calls scale to about a dozen people,
authorization is the server's job, and this is not Microsoft Teams: it implements channels, posts
with replies and reactions, files, search, presence and meetings, not 1:1 chats, notifications,
calendar or recording. Details in the core's `docs/teams-area.md`.

## Develop

```bash
bun run typecheck && bun run test   # vitest (jsdom) and the server's node:test suite
```

Svelte and Angular bindings are source-only and are not compiled by the repository's `tsc`
(their frameworks' own toolchains do that in a consuming app); React, Vue, Solid and vanilla are.

Licence: Apache-2.0.
