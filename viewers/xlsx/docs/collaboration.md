# Collaboration

Several people can edit one workbook at the same time. The editor maps the workbook to a Yjs document through `ooxml-core/xlsx/collab` and joins a room through the shared `ooxml-core/collab` session (the same session and providers the Word and Visio editors use). Every edit is sent as it is made, concurrent edits to different cells merge, and each person sees where the others are.

## Start sharing

From the editor: **File > Share** (or the Share button at the right end of the ribbon tab row), type a session name and choose **Start sharing**. Another window or tab of the same browser that has a workbook open and starts sharing with the same session name joins it. Without a server set by the page, nothing leaves the browser: the session runs over a `BroadcastChannel`.

From code, set the `collaboration` property (or prop, in every binding):

```ts
editor.newWorkbook(); // a window joining an existing room still needs a workbook open
editor.collaboration = {
	roomId: 'budget-2026',
	serverUrl: 'wss://collab.example.com/rooms', // y-websocket compatible; omit for same-browser only
	user: { name: 'Ada', color: '#2563eb' },
};
editor.addEventListener('collaboration-change', (event) => console.log(event.detail));
```

| Option      | Notes                                                                                                                               |
| ----------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `roomId`    | 1 to 128 letters, digits, hyphens or underscores.                                                                                   |
| `serverUrl` | Base URL of a y-websocket compatible server; the room id is appended.                                                               |
| `provider`  | A provider factory of your own (WebRTC signaling, or an existing Yjs provider through `externalProvider`). Wins over `serverUrl`.   |
| `session`   | An existing `CollabSession` created with `sanitizePayload: sanitizeXlsxPresence`. The editor never destroys a session it was given. |
| `user`      | `name` (defaults to `authorName`), `color` (defaults to a palette colour), `role` (`viewer` makes the editor read-only).            |

While `collaboration` is set, a workbook opened later joins the same room; when the room already has a workbook, the room's content replaces the opened one. Setting it to `null` leaves the room. The imperative forms are `startCollaboration(options)` (resolves once joined), `stopCollaboration()`, `reconnectCollaboration()`, `collaborationState` and `share()` (opens File > Share). Every binding forwards `collaboration`, reports `collaboration-change` (`onCollaborationChange`, Vue `collaboration-change`, Angular `collaborationChange`, Svelte `oncollaborationchange`) and exposes `share()` on its handle; unmounting the editor leaves the room.

## What is shared

- Cell values, formulas (with their cached values), number formats and cell formatting.
- Sheets: add, rename, move, delete, hide and tab colour.
- Row heights, column widths, hidden rows and columns, merged cells and defined names.
- Each person's selection, drawn in the others' grids as an outline in their colour with their name, and the people in the session as avatars in the title bar.

Undo (Ctrl+Z, the ribbon and the Quick Access Toolbar) undoes only your own edits while the workbook is shared; a collaborator's work is never reverted. When the connection drops, the editor says so and keeps your edits; they are sent when it returns, and a second notice confirms the reconnection.

## What is not shared

Charts, pictures and other drawings, comments, hyperlinks, conditional formats, data validation, tables and filters, pivots, freeze panes and other view settings, page setup, protection, document properties and calculation settings stay in each window and are saved from there. Row and column inserts and deletes travel as the cells they move, so a concurrent edit to a moved cell lands at its old address. A formula a collaborator typed with a sheet name that was renamed at the same time keeps the old name. The conflict rules are listed in the `ooxml` repository (`docs/collab-area.md`).

## What the host owns

The editor ships no server. Your application owns the transport (the server URL or provider), authentication and authorization (roles are advisory in the client), room persistence, and when and from which window the workbook is saved: the editor does not coordinate simultaneous saves. This is not Excel co-authoring and does not claim parity with it.
