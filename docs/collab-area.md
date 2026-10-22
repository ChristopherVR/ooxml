# The `collab` area (`ooxml-core/collab`)

Format-neutral real-time collaboration on Yjs. It owns the logic that has nothing to do with a particular document format or UI: session and provider lifecycle, awareness and presence, a transport-neutral sync protocol, update codecs, ordering helpers, binary asset sync and the seam products plug their document model into. `yjs`, `y-protocols` and `lib0` are real dependencies of the package (not optional peers); a host that also installs `yjs` itself must dedupe it, because two Yjs copies break `instanceof` checks.

## What moved here, what stayed

The format-neutral parts were generalised from the PowerPoint Yjs collaboration modules and the Word `prosemirror-collab` presence code, which were in the separate viewer repositories at the time. Per-module provenance is in `PROVENANCE.md`.

Moved (generalised, DOM-free):

- Validation and sanitising of anything a peer sends: room id, names (linear-time HTML strip), colours, avatar URLs, indices, coordinates, mixed-content detection.
- Identity: roles, palette, deterministic colours, labels, initials, roster, per-client collision-resistant id generator.
- Presence model, `presence` awareness field, derivation of the remote list (stale and malformed entries dropped), throttled publisher, borrowed-awareness lease for host-owned sessions.
- Lifecycle: `pagehide`/`beforeunload`/bfcache teardown and the synchronous same-browser departure channel (the viewers' ghost-collaborator fixes).
- Policies: first-write sync gate, "does the room replace a loaded document", broadcast auto-follow.
- Asset sync (binary payloads in a separate Y.Map, version counter for in-place swaps), parametrised by an `AssetSpec`.
- Docx-side ordering rules without ProseMirror: version classification, idempotent retries, sequence tracking, bounded caches.
- New: update codec with defensive application, typed emitter, `Transport` and `SyncProvider` interfaces, the y-websocket-compatible transport provider, in-memory and WebSocket transports, y-websocket/y-webrtc adapter, session object, `DocumentAdapter`/`bindDocument`.

Not format-neutral, so they live with their product (all in this repository):

- PowerPoint: the slide/element to Y.Map schema and its DOM-free sync (`collaboration-sync.ts`, `-reconcile`, `-live-patch*`, the `-text-*` codec, merge and session modules, `-assets`, `-departure`, `-load-origin`) are in `src/core/pptx/editor/render/`, with the field schema in `src/core/pptx/core/types/collaboration-field-schema.ts`. The editor-side modules (the inline editor and its DOM, `-shell-state`, `-presence-projector`, `-writeback`, `collaboration-options.ts` and `collaboration-owned-session.ts`) are in `src/ui/src/pptx/render/`.
- Word: the `prosemirror-collab` client and authority (`CollaborationClient`, `CollaborationAuthority`, step batches), the identity and protocol modules and the Yjs variant (`yjs-collaboration.ts`) are in `src/core/docx/ui/collaboration*.ts`. They depend on ProseMirror steps. Word is on an authority/step protocol, PowerPoint on a CRDT; this area does not unify the two engines, it shares everything around them (see "Who uses what").
- Excel and Visio map their models to Yjs in their own areas (`ooxml-core/xlsx/collab`, `ooxml-core/visio/collab`, below).

## API

```ts
import {
	createCollabSession,
	transportProvider,
	createMemoryHub,
	createWebSocketTransport,
	roomUrl,
	bindDocument,
	LOCAL_ORIGIN,
	type DocumentAdapter,
} from 'ooxml-core/collab';
```

**Session**

- `createCollabSession<P>({ roomId, provider, user: { name, color?, avatar?, role? }, doc?, initialPresence?, sanitizePayload?, onReady?, teardown?, heartbeatMs?, syncGraceMs?, connectionTimeoutMs?, autoConnect? })` returns a `CollabSession<P>`: `doc`, `awareness`, `clientId`, `status`, `synced`, `canWrite()`, `peers()`, `updatePresence(patch)`, `connect()`, `disconnect()`, `destroy()` and typed `on('status' | 'synced' | 'peers' | 'ready' | 'error', ...)`.
- `P` is the product's "where am I" payload (Word: selection anchor/head; PowerPoint: slide index, cursor, selected shape). `sanitizePayload(raw)` validates what peers send and returns `null` to hide a peer.
- `canWrite()` is true once the gate opened (provider synced, or the grace period passed for a lone peer on a mesh transport) and the role is not `viewer`. A user `disconnect()` resets the gate; a transient provider drop does not, so ordinary offline Yjs edits keep working and merge on reconnect.
- `reconnect()` explicitly disconnects and reconnects while retaining the same document and offline updates. `resync()` returns false if disconnected or the provider lacks this optional capability. The stock transport provider requests remote differences and sends a local snapshot to repair missing updates in both directions without leaving the room. This can send more bytes than an ordinary incremental update. Sync completion is not a durable-save acknowledgement.
- Attaching a provider that is already synced immediately opens the session gate. A read-only document binding also adopts an already-synced room; its role continues to prohibit publication.

**Providers and transports**

- `SyncProvider` (`connect`, `disconnect`, `destroy`, `status`, `synced`, `on('status' | 'synced' | 'error')`) is what a session talks to; `ProviderFactory = ({ doc, awareness, roomId }) => SyncProvider`.
- `Transport` is a byte channel: `connect(handlers)`, `send(bytes)`, `disconnect()`; handlers `open`, `close`, `message`, `error` and optional `peer` (mesh: a peer joined). Reconnect policy belongs to the transport.
- `transportProvider({ transport })` is the stock provider. It speaks the y-websocket wire format (sync step 1/2/update, awareness, query-awareness), so the WebSocket transport works against a stock y-websocket server. Inbound messages are decoded defensively: malformed, truncated, unknown or oversized messages become `error` events and are dropped; updates are structurally validated before they touch the document.
- `createWebSocketTransport({ url, WebSocket?, baseDelayMs?, maxDelayMs?, maxRetries?, pageProtocol? })` with exponential backoff and a mixed-content fail-fast. `createMemoryHub({ async?, filter? }).createTransport(room)` is the in-memory mesh used by tests and local-first prototypes. `createBroadcastTransport({ roomId, prefix?, BroadcastChannel? })` is the same-browser mesh over `BroadcastChannel`: tabs and windows of one origin share a room with no server and nothing leaving the device (`canBroadcast()` tells whether it is available).
- WebRTC: write a `Transport` over your data channel, or wrap an existing y-webrtc/y-websocket provider with `adaptYjsProvider(provider)`. `observeExternalSession` and `borrowAwareness` cover host-owned sessions the viewer must not destroy.

**Pure helpers**: `codec` (`encodeSnapshot`, `encodeDiff`, `applyUpdateSafe`, `validateUpdate`, `restoreSnapshot`, `mergeUpdates`, base64), `ordering` (`classifyVersion`, `IdempotencyCache`, `SequenceTracker`, `BoundedMap`, `freezeDeep`), `presence`/`identity`/`validation`, `policy`, `lifecycle`, `assets` (`createAssetSync(spec)`).

**Product seam**: `DocumentAdapter<TModel>` (`isEmpty`, `read`, `write(doc, model, LOCAL_ORIGIN)`, `observe`) and `bindDocument(session, adapter, { getLocalModel, onRemoteModel })`, which seeds an empty room once writes are allowed, adopts a non-empty one, skips echoes of its own writes and gates `push(model)` / `handleLoad(origin, model)` (bootstrap loads lose to the room, user loads are published).

`packageAdapter({ name?, maxBytes? })` is a ready-made `DocumentAdapter<{ bytes, revision }>` that shares the whole saved package (VSDX, DOCX, PPTX, XLSX). It is the fallback for products whose edits produce a new package and have no mergeable mapping (Excel and Visio have their own adapters, below). It never merges: the last package written wins, and packages above `maxBytes` (32 MiB by default) are refused.

## Write-back scheduler (`createWriteBackScheduler`)

One DOM-free scheduler for the elected writer (a room's owner), shared by the PowerPoint and Visio editors (`write-back.ts`, `write-back-lifecycle.ts`):

```ts
const writer = createWriteBackScheduler({
	serialize: () => exportBytes(), // Uint8Array | null | Promise of either; null means nothing to save yet
	write: (bytes) => upload(bytes), // void, or a Promise that is the acknowledgement
	debounceMs: 5000, // trailing debounce
	retry: { attempts: 4, baseDelayMs: 1000, maxDelayMs: 30000 }, // optional; attempts include the first try
	onStatus: (state) => {}, // idle | pending | saving | saved | failed, lastSavedAt, lastError
});
writer.schedule(); // on every change
await writer.flush(); // write now, resolves when acknowledged, rejects when it failed after its retries
writer.cancel(); // drop pending work; dispose() stops for good
```

At most one write is in flight; changes made during a write coalesce into one follow-up write after the next quiet period. A returned promise is the acknowledgement; a plain `void` return counts as acknowledged at once. A rejection (or a throwing `serialize`) is retried with exponential backoff up to `retry.attempts` tries, then the status is `failed` with `lastError` and the change stays unsaved: the next `schedule()` or `flush()` tries again. `flush()` skips a retry backoff. `timers` and `now` are injectable for tests. `writeBackStateOf(state)` gives the JSON-friendly `CollaborationWriteBackState` that editors report as the optional `writeBack` field of `CollaborationState` (owners only).

`flushWriteBackOnPageClose(scheduler, { window?, document?, onError? })` is the hook point for page close: it flushes unsaved work on `pagehide` and when `visibilitychange` turns hidden, and returns the function that removes the listeners. The targets are structural (default: the globals), so core stays DOM-free; the PowerPoint shared layer and the Visio viewer register it. Browsers do not wait for asynchronous work when a page unloads, so a write that needs a network round trip can still be lost on close. A host whose `write` must survive a close should send the bytes with `navigator.sendBeacon(url, bytes)` or `fetch(url, { method: 'PUT', body: bytes, keepalive: true })`, which the browser completes after the page is gone (keepalive bodies are capped at about 64 KiB, so a large document may still not make it).

## Embedding contract (every editor)

A host embedding more than one product integrates collaboration once (issue #152). Every editor element takes the same `collaboration` property and reports the same state; the types and helpers are in `embedding.ts`:

- `CollaborationOptions<P>`: `roomId`, `serverUrl` (y-websocket compatible), `authToken` (the server's `token` query parameter, or the WebRTC room password), `provider` (a `ProviderFactory`), `webrtc` (`{ WebrtcProvider, signaling?, password? }`: the host passes the class from its own `y-webrtc` install, wrapped by `webrtcProvider`, so the package takes no WebRTC dependency), `session` (a host-owned `CollabSession`, never destroyed by the editor), `user` (`name`, `color`, `avatar`, `role`) and `intent` (`join` never seeds an empty room). One transport is used, in this order: `session`, `provider`, `webrtc`, `serverUrl`, else a BroadcastChannel between windows of this browser.
- `CollaborativeEditor`: the element surface. `collaboration` is declarative (while set the open document is shared, `null` leaves), `collaborationState` returns a `CollaborationState` (`active`, `roomId`, `status`, `synced`, `people`), `stopCollaboration()` and `reconnectCollaboration()`. Elements dispatch `collaboration-change` (`COLLABORATION_CHANGE_EVENT`) with the state as detail; every framework binding takes a `collaboration` prop and relays the event in its usual style (Word and Excel: `onCollaborationChange`, `@collaboration-change`, `(collaborationChange)`, `oncollaborationchange`; Visio: its `events` map and Angular `(collaborationChange)`). PowerPoint accepts the same options (`session`, `provider`, `user`, `intent`, `roomId`, `serverUrl`) in its `collaboration` prop through `resolveCollaborationConfig`, and reports the same `CollaborationState` from every binding (React `onCollaborationChange`, Vue `collaboration-change`, Angular `(collaborationChange)`, Svelte `oncollaborationchange`, Vanilla `onCollaborationChange` and `getCollaborationState()`). Without a `session`, a `provider`, or shared-contract input naming no `serverUrl` or `transport` (a BroadcastChannel room, prefix `ooxml-pptx:`), is opened by the viewer through `openCollaborationSession` (`openOwnedCollaboration` in `ooxml-ui/pptx`) and destroyed when collaboration stops; a legacy config with `serverUrl` or `transport: 'webrtc'` keeps y-websocket and y-webrtc.
- `openCollaborationSession(options, defaults)` is what an editor calls: it returns the host's session (`owned: false`) or opens one it owns, with the product's broadcast prefix, presence sanitiser and default user name. `collaborationStateOf(session)` and `onCollaborationStateChange(session, listener)` derive the reported state.
- `toExternalSession(session)` presents a `CollabSession` as the borrowed `ExternalSession` (doc, awareness, stable `getSnapshot`, `subscribe`) that viewers taking host-owned Yjs resources observe; `synced` is the session's write gate.

The document engine behind the contract still differs per product (Excel merges cell by cell, Word maps its document to Yjs, Visio shares a drawing property by property and keeps only masters, styles, theme and similar parts last-write-wins, PowerPoint keeps its slide schema), so what merges concurrently differs too; each product's docs state its limits. A host session must be created with the product's presence sanitiser (for example `sanitizeXlsxPresence`), so one session cannot be passed to two different products.

## Honest limits

- Validation checks structure and size, not semantics: a well-formed update from an authorised peer can still carry content the product must validate when it reads the model back.
- There is no authentication, authorisation or persistence here; `role: 'viewer'` is advisory on the client. A server must enforce permissions. `demos/pptx/demo-host-review` shows one way (signed room tickets, `openteams-server`'s `authorizeSync` hook dropping writes from read-only sockets, live access changes); its recipe is the PowerPoint guide "Review inside your application" (`viewers/pptx/docs/guide/review-workflow.md`).
- Awareness has no integrity protection; presence is sanitised for display, not trusted.
- Asset entries are not garbage-collected when their owner is deleted.
- No Word layout, document or PowerPoint model parity is claimed: this area only moves bytes and presence.

## Who uses what

The viewers are workspace members of this repository and import `ooxml-core/collab` directly; there is no separate adoption step.

| Product    | Document engine                                                              | Where                                             |
| ---------- | ---------------------------------------------------------------------------- | ------------------------------------------------- |
| PowerPoint | Yjs slide schema, `createAssetSync`, shared presence and lifecycle helpers   | `src/core/pptx/editor/render/`, `src/ui/src/pptx` |
| Word       | `prosemirror-collab` steps, shared identity, ordering and validation helpers | `src/core/docx/ui/collaboration*.ts`              |
| Excel      | `ooxml-core/xlsx/collab` (field by field)                                    | `src/core/xlsx/collab/`, `src/ui/src/xlsx`        |
| Visio      | `ooxml-core/visio/collab` (property by property)                             | `src/core/visio/collab/`, `src/ui/src/visio`      |

The viewers keep only UI: cursor and selection rendering, the Share dialog, status chips and remote-selection decorations.

## xlsx adapter (`ooxml-core/xlsx/collab`)

The first product mapping that lives in its format area. It maps the spreadsheet model to Yjs field by field, so concurrent edits merge instead of replacing the whole package.

**Schema.** Four top-level maps (top-level types never conflict when two peers initialise a room at once; nested maps created concurrently would lose one side):

- `xlsx:meta`: schema version, the room's default format key, `date1904`.
- `xlsx:styles`: format key to resolved `CellStyle`. Style ids are indexes into one workbook's `styles` array, so two peers adding a format at once would both claim the next index. The shared table is content-addressed instead: the key is a 53-bit hash of the format's canonical JSON (`styleKey`), entries are immutable, and each peer interns what it reads into its own array. One copy per distinct format keeps it small; a cell carries only the key, and no key when it has the room's default format.
- `xlsx:names`: `sheetKey!name` (lower case; empty sheet key for workbook scope) to `{ n, f, k?, h?, c? }`.
- `xlsx:sheets`: sheet key to a `Y.Map` with `props` (name, sheetId, fractional `order`, state, tab colour, default row height and column width), `cells` (`A1` to `{ v, t, f?, s?, a?, r?, d?, l? }`: value, type `n`/`s`/`b`/`e`, formula, format key, array range, rich text, dynamic-array and legacy-formula flags), `rows` (1-based row number to size/custom/hidden/format/outline), `cols` (`C:E` span to the same) and `merges` (`A1:B2` to `true`).

A sheet key is minted by the peer that adds the sheet (`<clientId>-<n>`), never the name or index, so renames and moves keep it and concurrent adds never collide. Cell entries are plain JSON, so a cell is replaced atomically: concurrent edits to different cells both survive, edits to the same cell resolve to one value by Yjs ordering on every peer. Spill results are not shared (every peer recomputes them); formula cells carry their cached value when they are written, and every peer recalculates after applying a change.

**Deterministic conflict rules.** Names and sheetIds that collide after concurrent adds are made unique on read, in key order (`Sheet2`, `Sheet2 (2)`). Concurrent moves of different sheets both apply (orders are per sheet; a local move rewrites the fewest orders). Concurrent splits of one column span resolve, column by column, to the narrowest span (the most specific edit). Overlapping concurrent merges keep the first in row-major order. A deleted sheet wins over concurrent edits on it.

**API.**

- `bindWorkbookSession(session, editSession, { onError? })` is the one call a UI needs. Each local edit-session change becomes one transaction tagged `XLSX_EDIT_ORIGIN` that writes only what differs (the changed ranges for cell and format edits, the whole workbook for structural ones). Remote transactions are applied to the session through `editSession.applyExternal`, which recalculates and announces them to `onChange` listeners with `external: true` but never records them in the session's history. `undo()`/`redo()` run a `Y.UndoManager` that tracks only `XLSX_EDIT_ORIGIN`, one step per edit, so they never revert a collaborator's work; route Ctrl+Z through the binding while bound. It follows `bindDocument`'s lifecycle: seed an empty room once writes are allowed (origin `XLSX_SYNC_ORIGIN`, not undoable), adopt a non-empty one, and publish edits made while writes were blocked once they are allowed again. `setSelection(sheet, range)` and `remoteSelections()` carry presence as `{ sheet: <sheet key>, range: 'B2:C4' }`; pass `sanitizeXlsxPresence` as the session's `sanitizePayload`.
- `xlsxDocumentAdapter({ base? })` is a whole-model `DocumentAdapter<Workbook>` for `bindDocument`; `write` reconciles the full workbook (differences only), `read` materialises it onto a clone of `base` (so unshared parts survive) or a new workbook. `readSharedWorkbook` and `writeSharedWorkbook` are the same operations without the adapter.

**Not shared in this slice** (kept locally and saved as before, never sent): charts, drawings and images, comments, hyperlinks, conditional formats, data validations, tables, auto filters, pivots, sheet views (freeze panes, zoom, selection, gridlines), page setup and print options, sheet and workbook protection, outline settings, document properties, calculation settings, theme and named cell styles. Cell `valueMetadata`/`cellMetadata` (pointers into the source package) are not shared either.

**Honest limits.** Row and column inserts and deletes travel as the resulting cell moves (cells are keyed by address), so a concurrent edit to a moved cell lands at its old address. Formulas referencing a renamed or deleted sheet are rewritten by the peer that renamed or deleted it; a formula a collaborator typed concurrently with the old name keeps it. Cached values of dependent formulas are only refreshed in the room by the next structural write; peers always recalculate locally. `<xlsx-editor>` binds through it (the `collaboration` property and File > Share in `src/ui/src/xlsx/collaboration.ts`), routes undo through the binding while shared and paints remote selections laid out by `layoutRemoteSelections` (`xlsx/layout`).

## visio adapter (`ooxml-core/visio/collab`)

`<visio-viewer>` shares a drawing property by property instead of resending the package. The editor works on whole VSDX packages, so the mapping splits a package, merges the pieces in Yjs and rebuilds a package; the XML of each piece is kept as the text the editor wrote (no re-serialisation), which is why entries are XML strings and not decoded cell objects. Whitespace between elements is not kept.

Schema version 4 (top-level maps, entries are plain JSON replaced atomically per key):

| Map              | Key                                                                                       | Value                                                                                                                    |
| ---------------- | ----------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| `visio:meta`     | `v`, `pagesPath`, `pagesOpen`, `pagesClose`, `mastersPath`, `mastersOpen`, `mastersClose` | schema version; the `Pages` and `Masters` parts around their lists                                                       |
| `visio:pages`    | page key                                                                                  | `{ pos, id, back?, xml }` (the `Page` element)                                                                           |
| `visio:bodies`   | page key                                                                                  | `{ shell, rels?, b? }` (page part without shapes and connects, its relationships)                                        |
| `visio:shapes`   | `pageKey/shapeKey`                                                                        | `{ pos, parent?, open, b? }` (the opening tag of one `Shape`; nested shapes name a parent)                               |
| `visio:cells`    | `pageKey/shapeKey/<piece path>`                                                           | `{ pos, xml, r? }` for a leaf (`r`: id and shape key pairs the formula names), `{ pos, open }` for a section or row      |
| `visio:connects` | `pageKey/<from>.<cell>><to>.<cell>`                                                       | the `Connect` fields, shapes by key                                                                                      |
| `visio:dead`     | page key or `pageKey/shapeKey`                                                            | `true` (tombstone)                                                                                                       |
| `visio:masters`  | master key                                                                                | `{ pos, id, xml, part, rels?, b? }` (one master: list element, its part without shapes and connects as base64, its rels) |
| `visio:renum`    | page key, `m~<master key>` or `<scope>/<shape key>`                                       | `{ i?: [from, to], m?: [from, to] }` (a persisted renumbering of an id or `Master` reference)                            |
| `visio:media`    | `<sha-256>\|<part path>`                                                                  | `{ path, hash, ct, d, b? }` (one media part and its declared content type)                                               |
| `visio:blobs`    | sha-256 hex                                                                               | base64 of the media bytes, stored once per hash                                                                          |
| `visio:parts`    | part path                                                                                 | base64 of every other part                                                                                               |

The shapes of a master's content (`MasterContents`) are stored like page shapes, in `visio:shapes`, `visio:cells` and `visio:connects`, with the master scope `m~<master key>` in place of the page key; the master entry keeps only its attributes, its relationships and the rest of the part. Edits inside one master therefore merge per cell. Piece paths inside a shape: `c/<N>` (a top-level Cell), `s/<N[:IX]>` (a Section's opening tag), `s/<sec>/c/<N>` and `s/<sec>/r/<IX|N>` (a Cell and a Row of a section), `s/<sec>/r/<row>/c/<N>` (a Cell of a row), `x/<Name>#<k>` (Text, Data, ForeignData and any other child, whole). The rebuilt shape lists its pieces by `pos` (ties by path), so every peer serialises it identically.

**Schema version rule.** `meta.v` must equal `SCHEMA_VERSION` (4). A peer that finds any other version in a room, older or newer, refuses with a `SchemaVersionError`, reports it once through `onError`, and stays inert for that room: it never writes to it and never rebuilds the drawing from it. There is no migration: a version 1 room holds one entry per whole shape, which this schema cannot merge. Version 3 changed what a master entry holds and added `visio:renum`; version 4 added `r` to cell entries and made the renumbering records claim the id a shape keeps, so a version 2 or 3 room is refused too. A version 1 build that meets a newer room does not check the version and sees pages without shapes, so upgrade every peer of a room together.

`bindDrawingSession(session, editor, { seed, onError, schedule })` takes a DOM-free `DrawingEditor` (export bytes, apply a remote package, subscribe to local changes and idleness). A local edit writes only entries that differ from what this peer last wrote or applied (its baseline, compared in a serialisation-independent form), so a collaborator's change not yet applied here is never overwritten by a stale copy. Anyone else's transaction, and this binding's own undo, rebuild the package from the room and hand it to the editor outside its history. Undo is a `Y.UndoManager` tracking only the local edit origin. An empty room is seeded once the session can write (not with `seed: false`, i.e. `intent: 'join'`); a non-empty room is adopted and replaces the drawing opened before joining. Seeding derives keys from the drawing's ids (`s<id>`) so two peers opening the same file agree; later items get minted keys (`<client>.<n>`).

Conflict rules:

- Different shapes, pages or properties of one shape (cells, rows, sections, text): both survive. The same cell or piece: one value, the same on every peer (Yjs key order).
- Delete wins for pages and shapes: removing one writes a tombstone that a concurrent edit cannot undo (the shape's cells and nested shapes go with it). Undo of the delete removes the tombstone. A cell, master or media entry that one peer deletes and another edits at the same time keeps the edit (no tombstone, so a cell can be removed and added again).
- Order is a fractional `pos` per container; a move rewrites only the moved item, ties break by key.
- Ids: shapes (nested ones too, and the shapes inside masters), masters and pages that two peers numbered alike are made unique on read, by a rule that depends on the room's content only: an item whose id is persisted in `visio:renum` keeps it and is visited first; the others are visited in key order, the first keeps its id and each later one takes the next number above the maximum (`BackPage` follows for pages). A peer that can write then persists the result (a fix-up, below), after which the ids no longer depend on what is added later.
- Fix-up: after a rebuild, a peer that can write persists every renumbering it computed, in one transaction with `VISIO_FIXUP_ORIGIN`. Both the item that moved (`i: [from, to]`) and the item that kept the contested id (a claim, `i: [id, id]`) are written, and the `Master` reference of shapes written with a renumbered master. A claim is what stops an add that was concurrent with the collision, numbered alike and sorting first by key from taking the id: once persisted, a renumbering is never contradicted by content that arrives later. The write is a pure function of the room, so two peers writing it at the same time write equal values and converge. It is not a local edit: it is not in either user's undo history (the id lives in `visio:renum`, not in the entry the author wrote, so undoing the author's add still removes the whole item). Fix-ups only write `visio:renum`; no cell entry is ever rewritten.
- Formulas naming a shape: a cell entry whose `F` attribute contains `Sheet.<id>` records, in `r`, the key of the shape each id meant to its author (taken from the author's own package when the edit is written). The reader writes the final id of that shape into the formula it hands the editor, so a formula follows its target whatever happens to the numbers, in shapes of any batch and whether written before, during or after a collision. The room keeps the author's entry untouched, so the author's undo (or a collaborator's overwrite of the cell) reverts it normally; the editor's copy is the resolved text, which is also the peer's baseline, so it is never mistaken for a local edit. Not resolved: a reference whose target has no key in the author's package (a dangling or cross-page `Sheet.<n>`), text content containing `Sheet.<n>`, and `Master` inside master parts.
- Whole-part, last write wins: styles, theme, document settings, data, content types, the document's relationships and every other part not named above. Kept local: `windows.xml` and the thumbnail. Not shared: undo history, selection (presence is separate).

Rebuild cost: remote transactions that arrive before the scheduled flush (a zero-delay timer, injectable with `schedule`) share one rebuild; a rebuild whose resolved drawing has the same signature as the one last applied (and an unchanged editor version) is skipped; the serialised part of an unchanged page is reused from a cache. The package is written as a stored ZIP by a small writer (`zip-store.ts`) that keeps the CRC-32 of every part whose bytes did not change, so unchanged parts are copied, not re-checksummed. The binding tells the editor which pages changed (`applyRemote(bytes, { pages })`) when the page list, masters, media and shared parts are those it last applied and the editor was not edited since (the signature of everything but page content, `baseSignature`, and the version of the editor are unchanged); otherwise it passes nothing and the editor reads the whole package. The viewer controller then reads only those pages (`parseVsdx(bytes, { pages })`, in the same isolated worker) and splices them into the open document with `mergeVisioPages`, which reuses every other page as the same object. The document-level data (colours, fonts, themes, styles, masters) that shape inheritance needs is kept between such reads (`ParseVsdxOptions.documentCache`): the worker that reads pages stays for the next page read, and reuses the data when every part it was read from (judged by the CRC-32 and size their ZIP entries declare) and the options are the same, restoring the limit counters the read would have consumed and replaying its diagnostics; otherwise it reads them again. The worker is released by a whole read, a failure, a timeout, a cancel or `destroy()`. The controller reads the whole package when the merge cannot match a whole read (a page that is missing, more shapes than the shape limit, diagnostics it cannot place or a diagnostics limit that might apply) or when the page read fails.

Limits: whitespace between elements is dropped; a shape without an `ID` is not shared; a page part without a splittable `Shapes` list is shared as an opaque shell; two peers that persist a renumbering while a third add is still unseen can write different fix-ups for a moment; the room settles on the next read (a duplicate persisted id is resolved by key order, so an id persisted concurrently with a lower-keyed add can still move once); a formula is rewritten only for ids the author's package could name, and is never rewritten when the target was deleted; changes to the page list, masters, styles, theme, document settings and media always read the whole package, as does the first drawing a joiner adopts and the first remote change after a local edit; blobs are never deleted, so bytes of removed media stay in the room (not garbage-collected, like the asset entries above), which also means a media delete racing an add of the same bytes cannot lose them. No Office parity claim: the result is checked by re-parsing with this package's own reader.
