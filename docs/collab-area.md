# The `collab` area (`ooxml-core/collab`)

Format-neutral real-time collaboration on Yjs. It owns the logic that has nothing to do with a particular document format or UI: session and provider lifecycle, awareness and presence, a transport-neutral sync protocol, update codecs, ordering helpers, binary asset sync and the seam products plug their document model into. `yjs`, `y-protocols` and `lib0` are real dependencies of the package (not optional peers); a host that also installs `yjs` itself must dedupe it, because two Yjs copies break `instanceof` checks.

## What moved here, what stayed

Sources: `pptx-viewer/packages/shared/src/render/collaboration-*.ts` (Yjs) and `docx-viewer/packages/web-component/src/{collaboration,presence,editor-presence}.ts` (`prosemirror-collab`). Per-module provenance is in `PROVENANCE.md`.

Moved (generalised, DOM-free):

- Validation and sanitising of anything a peer sends: room id, names (linear-time HTML strip), colours, avatar URLs, indices, coordinates, mixed-content detection.
- Identity: roles, palette, deterministic colours, labels, initials, roster, per-client collision-resistant id generator.
- Presence model, `presence` awareness field, derivation of the remote list (stale and malformed entries dropped), throttled publisher, borrowed-awareness lease for host-owned sessions.
- Lifecycle: `pagehide`/`beforeunload`/bfcache teardown and the synchronous same-browser departure channel (the viewers' ghost-collaborator fixes).
- Policies: first-write sync gate, "does the room replace a loaded document", broadcast auto-follow.
- Asset sync (binary payloads in a separate Y.Map, version counter for in-place swaps), parametrised by an `AssetSpec`.
- Docx-side ordering rules without ProseMirror: version classification, idempotent retries, sequence tracking, bounded caches.
- New: update codec with defensive application, typed emitter, `Transport` and `SyncProvider` interfaces, the y-websocket-compatible transport provider, in-memory and WebSocket transports, y-websocket/y-webrtc adapter, session object, `DocumentAdapter`/`bindDocument`.

Stayed in the viewers (not format-neutral or not DOM-free):

- pptx-viewer: the slide/element to Y.Map schema (`collaboration-sync.ts`, `-reconcile`, `-live-patch*`, `-text-*` codec, merge and session modules, `-writeback`), the inline editor and its DOM (`-inline-*`), `-shell-state`, `-presence-projector` and the legacy flat `mapAwarenessCursors`. These become the pptx adapter in the `pptx` area when PowerPoint migrates.
- docx-viewer: the `prosemirror-collab` client and authority (`CollaborationClient`, `CollaborationAuthority`, step batches), the presence plugin and decorations, and the paragraph-id repair transaction. They depend on ProseMirror steps. Word is on an authority/step protocol, PowerPoint on a CRDT; this area does not unify the two engines, it shares everything around them (see "Adoption").

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

`packageAdapter({ name?, maxBytes? })` is a ready-made `DocumentAdapter<{ bytes, revision }>` that shares the whole saved package (VSDX, DOCX, PPTX, XLSX). It suits products whose edits produce a new package rather than a mergeable model. It never merges: the last package written wins, and packages above `maxBytes` (32 MiB by default) are refused.

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

The document engine behind the contract still differs per product (Excel merges cell by cell, Word maps its document to Yjs, Visio shares the whole package with the last write winning, PowerPoint keeps its slide schema), so what merges concurrently differs too; each product's docs state its limits. A host session must be created with the product's presence sanitiser (for example `sanitizeXlsxPresence`), so one session cannot be passed to two different products.

## Honest limits

- Validation checks structure and size, not semantics: a well-formed update from an authorised peer can still carry content the product must validate when it reads the model back.
- There is no authentication, authorisation or persistence here; `role: 'viewer'` is advisory on the client. A server must enforce permissions. `demos/pptx/demo-host-review` shows one way (signed room tickets, `openteams-server`'s `authorizeSync` hook dropping writes from read-only sockets, live access changes); its recipe is the PowerPoint guide "Review inside your application" (`viewers/pptx/docs/guide/review-workflow.md`).
- Awareness has no integrity protection; presence is sanitised for display, not trusted.
- Asset entries are not garbage-collected when their owner is deleted.
- No Word layout, document or PowerPoint model parity is claimed: this area only moves bytes and presence.

## Adoption (next wave)

**pptx-viewer** (Yjs already): replace `collaboration-presence`, `-presence-publisher`, `-departure`, `-teardown`, `-sync-gate`, `-load-origin`, `-broadcast-follow`, `-assets` and `-external-session` with imports from `/collab` (assets: `createAssetSync({ mapName: 'pptx:assets', fields: { mediaData: '_mdRef', posterFrameData: '_pfdRef', oleEmbeddedData: '_oedRef', previewImageData: '_pidRef', modelData: '_moRef' } })` keeps documents wire-compatible; pass `LEGACY_PPTX_LEAVE_MESSAGE` / `LEGACY_PPTX_DEPARTURE_CHANNEL` to keep the old embedder contract). Replace the per-binding `WebsocketProvider` wiring with `createCollabSession` plus `transportProvider(createWebSocketTransport(...))`, or `adaptYjsProvider(...)` while y-websocket/y-webrtc stay. Wrap the slide schema/reconcile code as a `DocumentAdapter<PptxSlide[]>` and use `bindDocument` instead of the readiness code. Presence keeps its wire shape through `deriveCanvasPresence`.

**docx-viewer** (`prosemirror-collab`): keep the step client/authority; import `isValidId`, `validateDisplayName`, `CURSOR_PALETTE`/identity helpers, `classifyVersion`, `IdempotencyCache`, `SequenceTracker`, `BoundedMap`, `freezeDeep`, `createIdGenerator(clientId, 'dve')` from `/collab` and delete the local copies. Cursor sharing can move to a `createCollabSession<{ anchor: number; head: number }>` presence channel (awareness instead of the hand-rolled `PresenceMessage`) once the host transport is a `Transport`; document content stays on step batches until Word chooses to model its document as Yjs. Do not change viewer repos before the package is published and `OOXML_CORE_REF` is bumped.

Both viewers keep only UI: cursor and selection rendering, the Share dialog, status chips and remote-selection decorations.

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

Schema version 2 (top-level maps, entries are plain JSON replaced atomically per key):

| Map              | Key                                                                                       | Value                                                                                      |
| ---------------- | ----------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| `visio:meta`     | `v`, `pagesPath`, `pagesOpen`, `pagesClose`, `mastersPath`, `mastersOpen`, `mastersClose` | schema version; the `Pages` and `Masters` parts around their lists                         |
| `visio:pages`    | page key                                                                                  | `{ pos, id, back?, xml }` (the `Page` element)                                             |
| `visio:bodies`   | page key                                                                                  | `{ shell, rels?, b? }` (page part without shapes and connects, its relationships)          |
| `visio:shapes`   | `pageKey/shapeKey`                                                                        | `{ pos, parent?, open, b? }` (the opening tag of one `Shape`; nested shapes name a parent) |
| `visio:cells`    | `pageKey/shapeKey/<piece path>`                                                           | `{ pos, xml }` for a leaf, `{ pos, open }` for a section or row                            |
| `visio:connects` | `pageKey/<from>.<cell>><to>.<cell>`                                                       | the `Connect` fields, shapes by key                                                        |
| `visio:dead`     | page key or `pageKey/shapeKey`                                                            | `true` (tombstone)                                                                         |
| `visio:masters`  | master key                                                                                | `{ pos, id, xml, part, rels?, b? }` (one master: list element, part base64, its rels)      |
| `visio:media`    | `<sha-256>\|<part path>`                                                                  | `{ path, hash, ct, d, b? }` (one media part and its declared content type)                 |
| `visio:blobs`    | sha-256 hex                                                                               | base64 of the media bytes, stored once per hash                                            |
| `visio:parts`    | part path                                                                                 | base64 of every other part                                                                 |

Piece paths inside a shape: `c/<N>` (a top-level Cell), `s/<N[:IX]>` (a Section's opening tag), `s/<sec>/c/<N>` and `s/<sec>/r/<IX|N>` (a Cell and a Row of a section), `s/<sec>/r/<row>/c/<N>` (a Cell of a row), `x/<Name>#<k>` (Text, Data, ForeignData and any other child, whole). The rebuilt shape lists its pieces by `pos` (ties by path), so every peer serialises it identically.

**Schema version rule.** `meta.v` must equal `SCHEMA_VERSION` (2). A peer that finds any other version in a room, older or newer, refuses with a `SchemaVersionError`, reports it once through `onError`, and stays inert for that room: it never writes to it and never rebuilds the drawing from it. There is no migration: a version 1 room holds one entry per whole shape, which this schema cannot merge. A version 1 build that meets a version 2 room does not check the version and sees pages without shapes, so upgrade every peer of a room together.

`bindDrawingSession(session, editor, { seed, onError, schedule })` takes a DOM-free `DrawingEditor` (export bytes, apply a remote package, subscribe to local changes and idleness). A local edit writes only entries that differ from what this peer last wrote or applied (its baseline, compared in a serialisation-independent form), so a collaborator's change not yet applied here is never overwritten by a stale copy. Anyone else's transaction, and this binding's own undo, rebuild the package from the room and hand it to the editor outside its history. Undo is a `Y.UndoManager` tracking only the local edit origin. An empty room is seeded once the session can write (not with `seed: false`, i.e. `intent: 'join'`); a non-empty room is adopted and replaces the drawing opened before joining. Seeding derives keys from the drawing's ids (`s<id>`) so two peers opening the same file agree; later items get minted keys (`<client>.<n>`).

Conflict rules:

- Different shapes, pages or properties of one shape (cells, rows, sections, text): both survive. The same cell or piece: one value, the same on every peer (Yjs key order).
- Delete wins for pages and shapes: removing one writes a tombstone that a concurrent edit cannot undo (the shape's cells and nested shapes go with it). Undo of the delete removes the tombstone. A cell, master or media entry that one peer deletes and another edits at the same time keeps the edit (no tombstone, so a cell can be removed and added again).
- Order is a fractional `pos` per container; a move rewrites only the moved item, ties break by key.
- Ids: shapes (nested ones too) and masters that two peers numbered alike are renumbered on read in key order (the lowest key keeps its id, others take the next numbers above the maximum). Pages collide the same way (`BackPage` follows).
- Formulas naming a renumbered shape: `Sheet.<old>` in the `F` attribute of cells is rewritten to the new id in the renumbered shape, in shapes nested in it, and in every shape written in the same batch (`b`, one write transaction of the author, which numbered them together). Not rewritten: references in shapes of other batches, because a reference written after its author adopted the new id cannot be told from a reference to the shape that kept the old id; text content containing `Sheet.<n>`; and the renumbered id itself can still change if a later concurrent add raises the page maximum. Renumbered masters rewrite `Master` on the shapes of their batch the same way; `Master` inside master parts is not rewritten.
- Masters: one entry each, so two peers adding different masters both keep theirs; the masters list, its relationships and the content type overrides are rebuilt from the entries. Two edits of the same master resolve to one (the whole master is the unit).
- Media: one entry per content hash and path. Identical bytes at one path are one entry (stored once); different bytes at one path keep the lowest key there and the others move to `<name>-<hash prefix><ext>`, with the relationship targets of the page bodies and masters of the same batch following; identical bytes at different paths stay at both. Content type declarations (`Default` or `Override`) are carried per media entry and restored on rebuild, because the shared content types part is last write wins.
- Whole-part, last write wins: styles, theme, document settings, data, content types, the document's relationships and every other part not named above. Kept local: `windows.xml` and the thumbnail. Not shared: undo history, selection (presence is separate).

Rebuild cost: remote transactions that arrive before the scheduled flush (a zero-delay timer, injectable with `schedule`) share one rebuild; a rebuild whose resolved drawing has the same signature as the one last applied (and an unchanged editor version) is skipped; the serialised part of an unchanged page is reused from a cache. The rebuilt package is still zipped whole and the editor re-reads it whole: there is no incremental parsing in the controller.

Limits: a master is one entry (concurrent edits inside the same master keep one); whitespace between elements is dropped; a shape without an `ID` is not shared; a page part without a splittable `Shapes` list is shared as an opaque shell; the ids of renumbered shapes are recomputed on every read; blobs are never deleted, so bytes of removed media stay in the room (not garbage-collected, like the asset entries above), which also means a media delete racing an add of the same bytes cannot lose them. No Office parity claim: the result is checked by re-parsing with this package's own reader.
