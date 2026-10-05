# Collaboration

Collaboration is what OpenTeams is: every channel, message, reaction and presence signal is shared state that several people edit at once. This page explains how it is built, what the live demo shows and what it does not.

## How it works

- **The shared document is a Yjs CRDT.** Channels live in one `Y.Map`, messages in one `Y.Map` per channel and reactions in a flat map keyed by message, emoji and user, so concurrent edits and reactions merge instead of overwriting each other. The logic lives in `ooxml-core/teams`, built on the format-neutral `ooxml-core/collab` area.
- **Presence is Yjs awareness**: who is online, who is typing and an availability you set (available, busy, away). It is advisory and unauthenticated, like any Yjs awareness.
- **Calls are WebRTC.** Signaling travels through your server (or `BroadcastChannel` in local mode); media goes peer to peer in a full mesh, so calls suit a handful of people.
- **The transport is pluggable.** Two modes exist: `server` (your server speaks the y-websocket sync protocol, relays call signaling and optionally stores files) and `local` (tabs of one browser share everything over `BroadcastChannel`).

## Live demo: two people, one room, any framework

The two panes below are two separate demo apps in one local room. Ada runs the vanilla `<teams-app>` and Bob runs the React binding, yet they share one workspace: a room is just a name (`?room=`), and every binding speaks the same core protocol. They run in local mode, so they talk to each other through your browser only. Type a message as Ada and watch it appear for Bob.

On the home page [live demo](/#live-demo) each pane has its own framework picker, so you can pair any two of Vanilla, React, Vue, Angular, Svelte and Solid.

<iframe
	src="/teams-viewer/demo/?name=Ada&room=docs-collab"
	title="OpenTeams demo as Ada"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/teams-viewer/demo-react/?name=Bob&room=docs-collab"
	title="OpenTeams React demo as Bob"
	loading="lazy"
	allow="clipboard-read; clipboard-write"
	style="width: 100%; height: 560px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

The same works across real tabs, in any pairing: open [the vanilla demo as Ada](/demo/?name=Ada&room=docs-collab){target="_self"} and, for example, [the Vue demo as Bob](/demo-vue/?name=Bob&room=docs-collab){target="_self"}. Every demo takes the same `name` and `room` parameters; only frames and tabs of the same browser can meet in local mode, so for people on different machines you point every binding at your own server instead (see below).

## What is and is not supported

| Supported                                                        | Not supported                                                                      |
| ---------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| Concurrent chat, replies and reactions merged by a Yjs CRDT      | End-to-end encryption: your server can read and store everything                   |
| Presence and typing indicators                                   | Authenticated identity: names are claims, and the reference server has no accounts |
| Same-browser collaboration with no server (local mode, as above) | Collaboration between different machines without a server you run                  |
| Cross-machine collaboration through your own sync server         | A hosted service: none is provided, and the Pages demos have no server behind them |
| Calls with audio, video, screen share and raise hand (mesh, ~12) | Webinar-size calls, recording, notifications                                       |

## Across machines

Run the reference server (`npx openteams-server`) or any server that honours the three contracts (sync, signaling, optional files) and set `config.mode` to `'server'`. See [Bring your own server](/server) and [Deploying](/deploy). Co-editing Office files is not part of OpenTeams: files are shared as links, and a host can open them in the Word, Excel, PowerPoint or Visio viewer, each with its own collaboration story.
