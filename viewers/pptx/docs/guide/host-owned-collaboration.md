---
title: Host-owned collaboration
description: Supply an existing Yjs document, awareness and synchronization state without giving the viewer ownership of the transport.
---

# Host-owned collaboration

Use `collaboration.externalSession` when your application already owns authentication,
a Yjs provider, offline storage or a shared connection. React, Vue, Angular, Svelte and
Vanilla accept the same contract. The viewer handles its slide schema and presence;
the host handles connecting, reconnecting, authorization and resource lifetime.

```ts
import type { CollaborationConfig, ExternalCollaborationSession } from 'pptx-react-viewer';

// These resources and the store belong to your application, not to a viewer mount.
const session: ExternalCollaborationSession = {
	doc: hostDocument,
	awareness: hostAwareness,
	getSnapshot: () => connectionStore.getSnapshot(),
	subscribe: (listener) => connectionStore.subscribe(listener),
};

const collaboration: CollaborationConfig = {
	roomId: 'presentation-123',
	serverUrl: '', // Required by the shared config, but unused for an external session.
	userName: 'Alice',
	sessionIntent: 'join',
	externalSession: session,
};
```

Import the types from the package for your framework. Pass `collaboration` to its
usual viewer prop or option. The document must be a real `Y.Doc`; the awareness must
belong to that document. Keep the session object stable and notify subscribers when
`status` or `synced` changes. A subscription returns an unsubscribe function.

The host and viewer must resolve the same Yjs runtime. The viewer leaves Yjs external
to its bundle and declares it as an optional peer. Install `yjs` in the application
for either built-in or host-owned collaboration:

```sh
npm install yjs
```

Ordinary viewing and editing without collaboration do not require this peer. The
host also installs its chosen provider and awareness implementation. Deduplicate
Yjs if your dependency tree contains multiple versions: types from separate runtime
copies cannot safely be mixed.

## Using the shared collaboration contract

Word, Excel and Visio take one `collaboration` option shape from `ooxml-core/collab`
(`CollaborationOptions`: `roomId`, `serverUrl`, `provider`, `session`, `user`, `intent`). PowerPoint accepts
the same shape in its `collaboration` prop alongside its own `CollaborationConfig`, so a host
embedding several products can hand them all one session:

```ts
import { createCollabSession, transportProvider } from 'ooxml-core/collab';

const session = createCollabSession({ roomId, provider: transportProvider({ transport }), user });
// <PowerPointViewer collaboration={{ session, user: { name: 'Ada' }, intent: 'join' }} />
```

- `session` is borrowed: it becomes the `externalSession`, and the viewer never destroys it. The
  host creates, connects and destroys it. `roomId` defaults to the session's room.
- `user.name`, `color`, `avatar` and `role` map to `userName`, `userColor`, `userAvatar` and
  `role`; `intent` maps to `sessionIntent`. Explicit legacy fields win when both are set.
- Every binding runs the same `resolveCollaborationConfig` (from `ooxml-ui/pptx`) where it reads the
  prop, so they behave identically. A legacy `CollaborationConfig` is returned unchanged.

Without a `session`, the same options also cover the cases a host does not want to build a session for:

```ts
// Your own transport: the viewer opens and destroys the session around it.
collaboration = { roomId, provider: transportProvider({ transport }), user: { name: 'Ada' } };
// No server: tabs and windows of this browser share over a BroadcastChannel.
collaboration = { roomId, user: { name: 'Ada' } };
```

The transport is chosen in the shared order: `session` (borrowed, never destroyed) > `provider` > `serverUrl` > BroadcastChannel. For `provider` and the BroadcastChannel room the viewer calls `openCollaborationSession` (the `ooxml-core/collab` runtime is imported only then, sharing the one Yjs copy), presents the session through the same borrowed-session path a host `session` takes, and destroys it on stop, unmount or when the config changes. The session is writable once it syncs, or after the grace period for a lone peer on a mesh; `intent: 'join'` still waits for the room's deck, and `role: 'viewer'` stays read-only. The state reports `connecting` while the session opens and `error` for an invalid room name or a browser without BroadcastChannel.

A legacy `CollaborationConfig` keeps its transports: `serverUrl` selects y-websocket, and `transport: 'webrtc'` or a blank `serverUrl` selects serverless y-webrtc (which also syncs same-browser tabs). Only shared-contract input that names no `serverUrl` or `transport` uses the BroadcastChannel room. The viewer's presence and the session's identity heartbeat share the awareness `presence` field; the viewer keeps its slide and cursor payload across heartbeats, so remote cursors keep showing.

## Reporting the shared state

Every binding reports the same `CollaborationState` Word, Excel and Visio report (`active`, `roomId`,
`status`, `synced`, `people`, the local person flagged `self`; type from `ooxml-core/collab`, also
re-exported by each package). It fires when a session starts or stops (the off state, `status: 'off'`)
and when the status, the synced flag or the people change, for host-owned and viewer-owned sessions
alike, and never repeats an unchanged state:

| Binding | Notification                                                 |
| ------- | ------------------------------------------------------------ |
| React   | `onCollaborationChange` prop                                 |
| Vue     | `collaboration-change` event                                 |
| Angular | `collaborationChange` output                                 |
| Svelte  | `oncollaborationchange` callback prop                        |
| Vanilla | `onCollaborationChange` option and `getCollaborationState()` |

Vanilla keeps `onCollaborationStatus` and `getCollaborationStatus()` for the raw connection status.

## Readiness is the host's decision

`getSnapshot()` returns `{ status, synced }`. Status is `connecting`, `connected`,
`disconnected` or `error`. `synced` means the authoritative document is available and
the viewer may publish edits. It is independent of network status:

- Keep `synced: false` while loading or hydrating the shared document. No timeout
  enables writes on the host's behalf.
- Set `synced: true` after initial synchronization. Existing room slides are adopted
  before any local bootstrap deck can be published.
- Keep it true during an offline period if edits may queue in the host's Yjs provider.
- Set it false to make the editor read-only and suspend durable write-back during
  a new sync. Editing resumes after the authoritative room state is adopted.

Accepted inline text and gesture previews are written to the borrowed document
synchronously, so a later readiness change cannot discard a queued final keystroke.
The host provider can still batch network traffic. An active text edit closes when
readiness is revoked; resuming adopts the authoritative room instead of replaying a
stale local draft.

Use `sessionIntent: 'create'` for the participant allowed to seed an empty room, and
`'join'` for participants whose startup deck must not seed it. An explicit File > Open
is still an intentional replacement. Coordinate room creation in the host; this is
not server-side authorization or an election protocol.

## Ownership and persistence

The viewer never creates a second provider, disconnects the host provider, or destroys
the supplied document or awareness. On unmount it removes its listeners and restores
only the presence field it still owns. Unrelated awareness fields and later host writes
are preserved. Use one active slide-presence writer per awareness instance.

The host may persist Yjs updates directly. Alternatively, designate one participant
with `role: 'owner'` and `onWriteBack(bytes)` to receive debounced PPTX snapshots, using
the same source PPTX and save options as the editor. Pending and in-flight snapshots
are cancelled if the session leaves or the host revokes readiness. Writer election remains the host's
responsibility. Return a promise from `onWriteBack` to acknowledge storage: the viewer reports `saving`, `saved` or `failed` in `collaborationState.writeBack` (owners only), serialises writes (one in flight, later changes coalesce into one follow-up) and retries a rejection as `writeBackRetry` (`{ attempts, baseDelayMs?, maxDelayMs? }`) says; a plain return counts as acknowledged at once. A pending write-back is flushed on `pagehide` and when the page turns hidden. Browsers do not wait for asynchronous work when a page unloads, so a write that needs a network round trip can still be lost on close. A host whose `write` must survive a close should send the bytes with `navigator.sendBeacon(url, bytes)` or `fetch(url, { method: 'PUT', body: bytes, keepalive: true })`, which the browser completes after the page is gone (keepalive bodies are capped at about 64 KiB, so a large document may still not make it).
An active session with `role: 'viewer'` makes the editor read-only
and blocks publication, for both built-in and host-owned collaboration. The server
must enforce access control independently.

## Integration boundaries

This option changes session ownership, not the existing collaboration data model.
The viewer owns `pptx:slides` and its nested Yjs types. Give all participants the same
source PPTX for package resources, themes and media. Other application data can occupy
separate Yjs keys. Do not independently reconcile complete slide snapshots from the host.

Check the host's message and persistence limits against representative decks, including
the initial document update and `Y.encodeStateAsUpdate(doc)` after editing. A compressed
PPTX file's size is not its Yjs snapshot size: the shared schema also carries parsed XML,
shape/group metadata and asset payloads. Moving images out of the document alone may
not be sufficient. This session API does not add transport chunking or change host limits.

It does not add comments-backend integration, a new local-only undo implementation,
or synchronization for every deck-level field. Validate those workflows separately
before exposing them in a collaborative product. A native PowerPoint reopen is also
a separate fidelity check from two browser peers converging.

## Trying the demos

Start the demo relay (`bun run collab`) and a framework demo, then open:

```text
/?externalSession=1&room=example&sample=1&server=ws://localhost:1234
```

A second participant uses the same room without `sample=1`. The host strip exposes
readiness and editor mount controls while retaining the document and provider. The
demo disables BroadcastChannel so synchronization uses the WebSocket relay.

Add `&headless=1` to try custom chrome in any of the five demos. These examples use
the binding's public building blocks without mounting a hidden full viewer. Their
host-owned Save button calls the public serialization path; the zoom selector lets
you check remote cursors and selection outlines at different scales.

See [React hooks](/react/hooks), [Vue collaboration](/vue/collaboration),
[Angular collaboration](/angular/collaboration), [Svelte collaboration](/svelte/collaboration)
and [Vanilla custom shells](/vanilla/api#custom-shell-collaboration) for each composition.

For a complete application around this option (accounts, permission-checked relay, live access
changes, owner versions and a gated AI route), see
[Review inside your application](/guide/review-workflow).
