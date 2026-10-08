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

## Honest limits

- Validation checks structure and size, not semantics: a well-formed update from an authorised peer can still carry content the product must validate when it reads the model back.
- There is no authentication, authorisation or persistence here; `role: 'viewer'` is advisory on the client. A server must enforce permissions.
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

**Honest limits.** Row and column inserts and deletes travel as the resulting cell moves (cells are keyed by address), so a concurrent edit to a moved cell lands at its old address. Formulas referencing a renamed or deleted sheet are rewritten by the peer that renamed or deleted it; a formula a collaborator typed concurrently with the old name keeps it. Cached values of dependent formulas are only refreshed in the room by the next structural write; peers always recalculate locally. The `<xlsx-editor>` element and its bindings are not wired to the binding yet.
