---
title: Review inside your application
description: A recipe for collaborative presentation review in a host application that owns accounts, permissions, the session, persistence and the AI key.
---

# Review inside your application

Collaboration is usually most useful inside a workflow your users already have: a deal desk,
a compliance queue, a CRM record, a project board. Several people open the same presentation
from that system, make small changes by hand or with the AI assistant, and the system keeps
track of who may do what and which version is current.

This recipe splits that work along one line:

| Your application owns                                    | The viewer owns                                        |
| -------------------------------------------------------- | ------------------------------------------------------ |
| Sign-in, accounts and the user's display name and colour | Rendering and editing the deck                         |
| Who may open, edit or only watch each item               | The slide schema in the shared Yjs document            |
| The Yjs document, awareness, provider and relay          | Presence (cursors, selection, "follow a collaborator") |
| When the session is ready, and when it stops being ready | Holding edits back until you say the room is ready     |
| Where saved versions go, and who writes them             | Turning the shared deck into PPTX bytes for the owner  |
| The AI model, its key and who may use it                 | The assistant panel, its deck tools and the proposals  |

The example app in
[`demos/pptx/demo-host-review`](https://github.com/ChristopherVR/ooxml/tree/main/demos/pptx/demo-host-review)
implements every step below with the vanilla binding. Every binding accepts the same
`collaboration` and `ai` options, so the steps read the same in React, Vue, Angular and Svelte.

## 1. Map your permissions to a role

Keep your own permission model. Derive the viewer's `role` from it when you mount the editor:

| Your permission (example) | `role`           | Notes                                                |
| ------------------------- | ---------------- | ---------------------------------------------------- |
| Owner of the item         | `'owner'`        | Edits, and receives `onWriteBack` snapshots to save  |
| Reviewer                  | `'collaborator'` | Edits live with the others                           |
| Watcher                   | `'viewer'`       | Read-only; also pass `editable: false`               |
| No access                 | (do not mount)   | Your API should not even reveal that the item exists |

`role: 'viewer'` is enforced **in the browser only**. Enforce it on the server as well (step 3).

## 2. Own the session

Create the Yjs document, the awareness and the provider yourself, and pass them as
`collaboration.externalSession` (see [Host-owned collaboration](/guide/host-owned-collaboration) for the
full contract). The viewer never connects, reconnects or destroys them.

```ts
import { createPptxViewer } from 'pptx-vanilla-viewer';

const grant = await api.requestGrant(reviewId); // your API: ticket, role, intent, user
const doc = new Y.Doc();
const awareness = new Awareness(doc);
const provider = new WebsocketProvider(`wss://${location.host}/collab`, reviewId, doc, {
	awareness,
	params: { ticket: grant.ticket },
});

const viewer = createPptxViewer(container, {
	source: sourceBytes, // the same source PPTX for every participant
	editable: grant.role !== 'viewer',
	collaboration: {
		roomId: reviewId,
		serverUrl: '', // unused with an external session
		userName: grant.user.name, // from your directory, not typed by the user
		userColor: grant.user.color,
		role: grant.role,
		sessionIntent: grant.intent,
		externalSession: { doc, awareness, getSnapshot, subscribe },
	},
});
```

`getSnapshot()` returns `{ status, synced }`. Report `synced: true` only when the provider has
synchronised **and** the user still holds the access the editor was mounted with. Setting it to
`false` makes the editor read-only at once.

**Who seeds the room.** An empty room is filled by the first participant with
`sessionIntent: 'create'`; everybody else passes `'join'` and adopts what is there. Let your
server decide, not the clients: the example grants `'create'` to one writer at a time with a
short lease and only while the room is still empty, so two people opening the item at the same
moment never both publish their startup deck.

## 3. Admit and police sockets on the server

A WebSocket cannot carry your site's session cookie to a relay on another origin, and a long-lived
token in a URL leaks into logs. Issue a **short-lived, signed ticket** from an ordinary
authenticated API call instead, and let the relay trust only what the ticket says:

- **On upgrade**: verify the ticket's signature, expiry and room. Refuse anything else.
- **On every message**: a read-only connection may ask for the document (sync step 1) and share
  presence, but its sync step 2 and update messages are dropped. This is the server half of
  `role: 'viewer'`.
- **Awareness**: let each connection speak only for the awareness client ids it first claimed, so
  nobody can clear or impersonate another person's cursor.
- **Ticket refresh**: a ticket only admits a connection. Refresh it before it expires and update
  `provider.params.ticket`, so an automatic reconnect uses a valid one.

You do not have to write this relay. `openteams-server` (the reference server of the OpenTeams
viewer, a y-websocket compatible Yjs server) takes an `authorizeSync` hook that returns `'write'`,
`'read'` or a refusal per socket, already pins awareness ids, and exposes `closeSync`,
`syncPrincipals`, `isSyncRoomEmpty` and `handleUpgrade` for the steps below. The example's
`server/relay.ts` is about sixty lines around it.

## 4. Change access while people are working

When an owner changes someone's access, close that person's open sockets with an application
close code (the example uses `4403`). The client reacts by:

1. reporting `synced: false`, which freezes the editor;
2. destroying the viewer, then its own document and provider;
3. asking for a new grant and mounting a fresh session with the new role, or showing "access
   removed" when there is none.

Start from a **new** Y.Doc after a downgrade. Edits the user made just before the change were
dropped by the relay, and reusing the old document would keep showing them locally.

## 5. Save versions with one writer

Mount exactly one participant, normally the item's owner, with `role: 'owner'` and
`onWriteBack(bytes)`. The viewer debounces changes (`writeBackDebounceMs`, 5 seconds by default)
and hands you a PPTX snapshot of the shared deck to store as a version. Other participants never
write back, so there is no last-save-wins race.

`onWriteBack` is not a durable-storage acknowledgement: retries, failures and choosing a writer
when the owner is offline remain your decisions. If the room must survive a relay restart,
persist the Yjs updates on the server too (for example with your own store, or a provider
backend such as y-leveldb).

## 6. Add AI-assisted changes

Pass the `ai` option only to accounts that may use it, and point it at **your** route so the
model key never reaches the browser:

```ts
ai: {
	connection: {
		kind: 'endpoint',
		api: `/api/ai/chat?review=${reviewId}`,
		credentials: 'same-origin',
	},
	writePolicy: 'stage', // every change is a proposal the reviewer accepts or rejects
},
```

The route checks the caller's account and permission, then runs the model with the schema-only
deck tools from `ooxml-ui/pptx/ai` (see [AI Assistant](/guide/ai-assistant)). Tool calls run in the
reviewer's browser against the live deck. An accepted proposal is an ordinary edit by that
reviewer: it goes through the editor history (one Ctrl+Z reverts it) and reaches the room like
typing does. Your audit log can record that the assistant was used; the shared document does not
mark which edits came from it.

## Run the example

```sh
bun install                                   # at the repository root
cd demos/pptx/demo-host-review
bun run start                                 # http://localhost:4190
```

Open the page in two browsers (or one normal and one private window) and sign in as different
people: Ana owns the item, Ben and Chen review, Dee watches. As Ana, change Dee's access in the
side panel and watch Dee's editor switch between read-only, editing and no access without a
reload. To try the assistant, set `REVIEW_AI_MODEL` (and `REVIEW_AI_API_KEY`, or
`REVIEW_AI_BASE_URL` for a local OpenAI-compatible server) before starting.

## Limits

- The example keeps everything in memory: accounts, versions and the shared rooms are gone when
  the process stops. Its sign-in is an account picker.
- Awareness content (the display name and colour) comes from the client. The relay pins awareness
  ids to connections but does not check the payload against the ticket's account.
- Comments, a review approval state and per-change attribution are not part of the shared model
  in this example; validate those workflows separately before you rely on them.
- Browser peers converging is not the same as PowerPoint reopening the saved version without a
  repair; check representative decks in PowerPoint before you promise that.
