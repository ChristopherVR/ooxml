# Collaboration

File > Share starts a live session between windows of one browser. A host application can share across devices by setting the `collaboration` property, which takes the same options as every other Office editor.

::: warning Partial
There is no hosted service in this build. With File > Share nothing leaves your browser. Across devices you bring your own server (a y-websocket compatible URL), provider or Yjs session; sharing is whole-package, last write wins.
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

- **Transport**: `session` (a host-owned `CollabSession`, never destroyed by the viewer), else `provider`, else `serverUrl`, else a `BroadcastChannel` between windows of this browser.
- **Declarative**: while `collaboration` is set the open drawing is shared, and a drawing loaded later is shared too. Setting it to `null` leaves the room. A session the viewer opened itself is destroyed on leave; a host `session` is not.
- **Intent**: with `join` an empty room is not seeded with the local drawing, so a late joiner cannot overwrite a room that has not loaded yet. A drawing loaded or edited afterwards is still published. Default `create`.
- **Roles**: `role: 'viewer'` makes the viewer read-only while it is in the room (`canEdit` is restored when it leaves). The role is advisory; a server must enforce it.
- **Events**: `collaboration-change` carries a `CollaborationState` when sharing starts or stops and when the connection status, first sync or people change. A failed join is reported as `document-error`. File > Share uses the same controller, so starting from the panel reports state and the event too.
- **Bindings**: every binding forwards a `collaboration` prop (React, Vue, Solid, Svelte, vanilla `mountViewer`/`update`, Angular `[collaboration]`) and relays `collaboration-change` (`events['collaboration-change']`, Vue `@collaboration-change`, Angular `(collaborationChange)`). Unmounting leaves the room.
- **Types**: `VisioCollaborationOptions` and `VisioCollaborationState` are exported from `ooxml-ui/visio`.

## What exists today

- A live session built on `ooxml-core/collab` (Yjs) over a `BroadcastChannel`, so windows and tabs of the same browser can share a drawing.
- The whole VSDX package is shared, and the collab runtime loads only when sharing starts.
- Remote changes arrive as one undoable history step, and collaborators appear through the shared `ooxml-ui` presence element.
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

| Supported                                                      | Not supported                                                                 |
| -------------------------------------------------------------- | ----------------------------------------------------------------------------- |
| Same-browser sessions over BroadcastChannel, built on Yjs      | A hosted server: bring your own y-websocket server, provider or session       |
| Whole-package sharing with one undoable step per remote change | Field-level merging: concurrent changes are not merged, the last package wins |
| Collaborators shown with the shared presence element           | Comments and authentication; roles are advisory (a server must enforce them)  |
| File > Share and the `collaboration` property as entry points  | Live cursors or selections of collaborators on the canvas                     |

The evidence is recorded in the [capability ledger](/parity) under "Collaboration and comments".
