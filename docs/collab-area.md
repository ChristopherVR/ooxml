# The `collab` area (`@christophervr/ooxml-core/collab`)

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
} from '@christophervr/ooxml-core/collab';
```

**Session**

- `createCollabSession<P>({ roomId, provider, user: { name, color?, avatar?, role? }, doc?, initialPresence?, sanitizePayload?, onReady?, teardown?, heartbeatMs?, syncGraceMs?, connectionTimeoutMs?, autoConnect? })` returns a `CollabSession<P>`: `doc`, `awareness`, `clientId`, `status`, `synced`, `canWrite()`, `peers()`, `updatePresence(patch)`, `connect()`, `disconnect()`, `destroy()` and typed `on('status' | 'synced' | 'peers' | 'ready' | 'error', ...)`.
- `P` is the product's "where am I" payload (Word: selection anchor/head; PowerPoint: slide index, cursor, selected shape). `sanitizePayload(raw)` validates what peers send and returns `null` to hide a peer.
- `canWrite()` is true once the gate opened (provider synced, or the grace period passed for a lone peer on a mesh transport) and the role is not `viewer`. A user `disconnect()` resets the gate; a transient provider drop does not, so ordinary offline Yjs edits keep working and merge on reconnect.

**Providers and transports**

- `SyncProvider` (`connect`, `disconnect`, `destroy`, `status`, `synced`, `on('status' | 'synced' | 'error')`) is what a session talks to; `ProviderFactory = ({ doc, awareness, roomId }) => SyncProvider`.
- `Transport` is a byte channel: `connect(handlers)`, `send(bytes)`, `disconnect()`; handlers `open`, `close`, `message`, `error` and optional `peer` (mesh: a peer joined). Reconnect policy belongs to the transport.
- `transportProvider({ transport })` is the stock provider. It speaks the y-websocket wire format (sync step 1/2/update, awareness, query-awareness), so the WebSocket transport works against a stock y-websocket server. Inbound messages are decoded defensively: malformed, truncated, unknown or oversized messages become `error` events and are dropped; updates are structurally validated before they touch the document.
- `createWebSocketTransport({ url, WebSocket?, baseDelayMs?, maxDelayMs?, maxRetries?, pageProtocol? })` with exponential backoff and a mixed-content fail-fast. `createMemoryHub({ async?, filter? }).createTransport(room)` is the in-memory mesh used by tests and local-first prototypes.
- WebRTC: write a `Transport` over your data channel, or wrap an existing y-webrtc/y-websocket provider with `adaptYjsProvider(provider)`. `observeExternalSession` and `borrowAwareness` cover host-owned sessions the viewer must not destroy.

**Pure helpers**: `codec` (`encodeSnapshot`, `encodeDiff`, `applyUpdateSafe`, `validateUpdate`, `restoreSnapshot`, `mergeUpdates`, base64), `ordering` (`classifyVersion`, `IdempotencyCache`, `SequenceTracker`, `BoundedMap`, `freezeDeep`), `presence`/`identity`/`validation`, `policy`, `lifecycle`, `assets` (`createAssetSync(spec)`).

**Product seam**: `DocumentAdapter<TModel>` (`isEmpty`, `read`, `write(doc, model, LOCAL_ORIGIN)`, `observe`) and `bindDocument(session, adapter, { getLocalModel, onRemoteModel })`, which seeds an empty room once writes are allowed, adopts a non-empty one, skips echoes of its own writes and gates `push(model)` / `handleLoad(origin, model)` (bootstrap loads lose to the room, user loads are published).

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
