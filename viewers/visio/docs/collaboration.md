# Collaboration

File > Share starts a live session between windows of one browser. A host application can share across devices by setting the `collaboration` property, which takes the same options as every other Office editor.

::: warning Partial
There is no hosted service in this build. With File > Share nothing leaves your browser. Across devices you bring your own server (a y-websocket compatible URL), provider or Yjs session; pages and shapes merge item by item, everything else is shared whole (see below).
:::

## Embedding: the `collaboration` property

`<visio-viewer>` implements the shared embedding contract of `ooxml-core/collab` (`CollaborationOptions`, `CollaborationState`, `CollaborativeEditor`), the same one Word, Excel and PowerPoint take, so a host that embeds several products integrates once.

```ts
viewer.collaboration = {
	roomId: 'floor-plan', // 1 to 128 letters, digits, hyphens or underscores
	serverUrl: 'wss://example.com/collab', // optional: a y-websocket compatible server
	user: { name: 'Ada', color: '#2563eb' }, // optional; `role: 'viewer'` makes this window read-only
	intent: 'join', // optional: never seed an empty room with the drawing already open
};
viewer.addEventListener('collaboration-change', (event) => console.log(event.detail));
viewer.collaborationState; // { active, roomId, status, synced, people }
viewer.reconnectCollaboration();
viewer.stopCollaboration(); // or: viewer.collaboration = null
```

- **Transport**: `session` (a host-owned `CollabSession`, never destroyed by the viewer), else `provider`, else `webrtc`, else `serverUrl`, else a `BroadcastChannel` between windows of this browser.
- **Declarative**: while `collaboration` is set the open drawing is shared, and a drawing loaded later is shared too. Setting it to `null` leaves the room. A session the viewer opened itself is destroyed on leave; a host `session` is not.
- **Intent**: with `join` an empty room is not seeded with the local drawing, so a late joiner cannot overwrite a room that has not loaded yet. A drawing loaded or edited afterwards is still published. Default `create`.
- **Roles**: `role: 'viewer'` makes the viewer read-only while it is in the room (`canEdit` is restored when it leaves). The role is advisory; a server must enforce it.
- **Events**: `collaboration-change` carries a `CollaborationState` when sharing starts or stops and when the connection status, first sync or people change. A failed join is reported as `document-error`. File > Share uses the same controller, so starting from the panel reports state and the event too.
- **Bindings**: every binding forwards a `collaboration` prop (React, Vue, Solid, Svelte, vanilla `mountViewer`/`update`, Angular `[collaboration]`) and relays `collaboration-change` (`events['collaboration-change']`, Vue `@collaboration-change`, Angular `(collaborationChange)`). Unmounting leaves the room.
- **Types**: `VisioCollaborationOptions` (the shared options plus `onWriteBack` and `writeBackDebounceMs`) and `VisioCollaborationState` are exported from `ooxml-ui/visio`, and every binding's `collaboration` prop is typed with the same options.

## Server and peer-to-peer: auth token and WebRTC

```ts
// A y-websocket compatible server. The token is sent as the `token` query parameter:
// wss://example.com/collab/floor-plan?token=<authToken>
viewer.collaboration = {
	roomId: 'floor-plan',
	serverUrl: 'wss://example.com/collab',
	authToken: await getToken(),
};

// Peer to peer over y-webrtc. Pass the `WebrtcProvider` class from your own y-webrtc install;
// the viewer does not bundle it. The room password defaults to `authToken`.
import { WebrtcProvider } from 'y-webrtc';
viewer.collaboration = {
	roomId: 'floor-plan',
	authToken: 'room password',
	webrtc: { WebrtcProvider, signaling: ['wss://signal.example.com'] },
};
```

The viewer hands these options to `openCollaborationSession` untouched. The token only authenticates to the server you run (the server must check it), and a WebRTC password only keeps peers without it out of the room; neither makes roles anything more than advisory. Peer-to-peer rooms never confirm a first sync, so the session's grace period decides when this window may write.

## Live cursors and selections

Collaborators on the same page see each other's pointer (a cursor with the person's name in their colour) and selected shapes (a dashed outline in their colour) over the canvas. They move and scale with the zoom and pan of the drawing, and disappear when the person leaves, goes to another page, moves the pointer off the page, clears the selection, or goes stale.

The payload each window publishes is `{ page, x?, y?, selection? }`: the page id, the pointer in the page's drawing units (inches) and up to 64 selected shape ids. `sanitizeVisioPresence` (exported from `ooxml-core/visio/ui`) validates every inbound payload. The viewer applies it to sessions it opens itself. **A host-owned `session` should be created with it**, otherwise its peers carry no payload and no cursors or selections are drawn (nothing breaks):

```ts
import { createCollabSession } from 'ooxml-core/collab';
import { sanitizeVisioPresence, VISIO_INITIAL_PRESENCE } from 'ooxml-core/visio/ui';

const session = createCollabSession({
	roomId: 'floor-plan',
	provider,
	user: { name: 'Ada' },
	sanitizePayload: sanitizeVisioPresence,
	initialPresence: VISIO_INITIAL_PRESENCE,
});
viewer.collaboration = { session };
```

Limits: the cursor and selection are shown only for the page both people have open; selection outlines cover top-level and nested shapes by their saved frame, not by their rendered geometry.

## Owner write-back

A host that wants the shared drawing saved can set `onWriteBack` and optionally `writeBackDebounceMs` (default 5000):

```ts
viewer.collaboration = {
	roomId: 'floor-plan',
	user: { name: 'Ada', role: 'owner' },
	onWriteBack: (bytes) => upload(bytes), // VSDX bytes
	writeBackDebounceMs: 5000,
};
```

Only the window whose role is `owner` writes, so collaborators never race each other to save. The callback fires after the drawing has been quiet for the debounce, following local and remote changes, and once more when the owner leaves the room if a save is pending. It never fires for `collaborator` or `viewer`. This is not a durable-storage acknowledgement: the viewer cannot tell whether your callback stored the bytes, does not retry, and a window closed before the debounce ends loses a pending save (leaving through `stopCollaboration()` or unmounting flushes it). The same options shape as PowerPoint's `onWriteBack`.

## What exists today

- A live session built on `ooxml-core/collab` (Yjs) over a `BroadcastChannel`, so windows and tabs of the same browser can share a drawing.
- Pages, shapes and connectors are shared item by item (`ooxml-core/visio/collab`); every other part of the package (masters, styles, theme, data, media, document settings) is shared whole, last write wins per part. The collab runtime loads only when sharing starts.
- Concurrent edits to different shapes both survive; two edits of one shape resolve to one value on every window; a deleted shape or page wins over a concurrent edit of it; shapes given the same id by two windows at once are renumbered the same way everywhere.
- Undo and redo (Ctrl+Z, the ribbon) act on this window's own edits only and never revert a collaborator's. Remote changes are applied outside your undo history, and the history starts when you join.
- Collaborators appear through the shared `ooxml-ui` presence element.
- The session name is 1 to 128 letters, digits, hyphens or underscores. A window joining a session whose room already holds a drawing adopts it; otherwise it seeds the room (unless the host asked for `intent: 'join'`).
- Panel sessions keep their original BroadcastChannel name (`ooxml-collab:visio-<room>`), so windows on an older build still meet.

## Live demo: two windows, any framework

A sharing session is just a name. Both panes below are separate playgrounds, built from different framework bindings, that join the session `docs-collab` on load: window A is the vanilla playground and opens the sample, window B is the React playground and receives the drawing from the room. Edit shape text in one pane and watch the other follow; undo reverts the remote step.

<iframe
	src="/ooxml/visio/demo/?embed=1&sample=1&share=docs-collab"
	title="visio-viewer window A (vanilla)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px"
></iframe>

<iframe
	src="/ooxml/visio/demo-react/?embed=1&share=docs-collab"
	title="visio-viewer window B (React)"
	loading="lazy"
	allow="clipboard-read; clipboard-write; fullscreen"
	style="width: 100%; height: 640px; border: 1px solid var(--vp-c-divider); border-radius: 8px; margin-top: 1rem"
></iframe>

### Pair any two frameworks

Every playground accepts `?share=<session>`, which fills in File > Share and starts it. Only the session name matters, so any pairing works: the home page [live demo](/#live-demo) has a framework picker for each window, and in two tabs you can open, for example, `/demo-vue/?sample=1&share=my-room` and `/demo-angular/?share=my-room`. You can also do it by hand in any two windows: File, then Share, then the same session name.

### Why File > Share is same-browser only

The panel uses a `BroadcastChannel`, which connects windows, tabs and frames of one browser profile on one origin, whichever framework they were built with, and nothing else. Two different machines cannot meet this way; a host sets `collaboration` with a `serverUrl`, `provider` or `session` for that, and this build ships no server.

## What is and is not supported

| Supported                                                     | Not supported                                                                |
| ------------------------------------------------------------- | ---------------------------------------------------------------------------- |
| Same-browser sessions over BroadcastChannel, built on Yjs     | A hosted server: bring your own y-websocket server, provider or session      |
| Per-shape and per-page merging, deletes win, local-only undo  | Merging inside one shape: two edits of the same shape keep one whole shape   |
| Collaborators shown with the shared presence element          | Comments; roles are advisory (a server must enforce them)                    |
| File > Share and the `collaboration` property as entry points | Cursors across different pages (only the shared page is drawn)               |
| Live cursors and selections of collaborators on the same page | Presence from a host `session` made without `sanitizeVisioPresence`          |
| Owner write-back (`onWriteBack`, debounced, flushed on leave) | A storage acknowledgement, retry or durability guarantee for write-back      |
| `authToken` for a server and `webrtc` (your `WebrtcProvider`) | A bundled y-webrtc or signaling server; token validation (your server's job) |

### Limits of shape-level merging

- A shape is one shared entry (its XML). Two people editing different properties of the same shape at once keep only one version of that shape.
- Page parts are split by text, so shapes round-trip exactly as the editor wrote them. A page part that cannot be split (no usable `Shapes` list) falls back to sharing its shapes as part of the page shell, last write wins.
- Parts other than pages (masters, document settings, styles, theme, data, media) merge per part, last write wins. Masters or media added by two people at the same moment can overwrite each other. The window state and thumbnail stay local.
- When two windows add shapes with the same id at once, one is renumbered. Formulas in other shapes that name the renumbered shape (`Sheet.n!...`) are not rewritten.
- A glue (connect) removed by one person and re-made identically by another keeps the re-made one.
- Undo history is per window, starts at join, and is not kept across a reconnect. The rebuild applies the room's whole package, so a very large drawing is re-read on every remote change.

The evidence is recorded in the [capability ledger](/parity) under "Collaboration and comments".
