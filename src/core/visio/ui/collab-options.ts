// The `collaboration` options of a Visio viewer: the shared embedding contract of
// `ooxml-core/collab` with the Visio presence payload and the owner's write-back hook. Types only,
// so the viewer bindings (React, Vue, Angular, Svelte, Solid, vanilla) take them from the contract.
import type { CollabSession, CollaborationOptions, WriteBackRetry } from '../../collab/index';
import type { VisioPresence } from './collab-presence';

/** Write-back settings of {@link VisioCollaborationOptions}. */
export interface VisioWriteBackOptions {
	/**
	 * Called with the saved VSDX bytes, debounced, after local and remote changes, only in the
	 * window whose `user.role` is `owner`. Return a promise to acknowledge durable storage: the
	 * viewer then reports `saving`, `saved` or `failed` in `collaborationState.writeBack`, and
	 * retries a rejection as `writeBackRetry` says. A plain return counts as acknowledged at once.
	 * A pending write-back is flushed when the viewer leaves the room and when the page is hidden
	 * or closing; browsers do not wait for asynchronous work then, so a write that needs a network
	 * round trip should use `navigator.sendBeacon` or `fetch(url, { keepalive: true })`.
	 */
	onWriteBack?: (bytes: Uint8Array) => void | Promise<void>;
	/** Quiet time after the last change before `onWriteBack` fires. Default 5000. */
	writeBackDebounceMs?: number;
	/** Retries of a rejected `onWriteBack` with exponential backoff. Default: none. */
	writeBackRetry?: WriteBackRetry;
}

/**
 * The options of `<visio-viewer>.collaboration`. A host-owned `session` should be created with
 * `sanitizeVisioPresence` as its `sanitizePayload` (and `VISIO_INITIAL_PRESENCE` as
 * `initialPresence`); a session made without them is accepted, but its peers carry no cursors or
 * selections.
 */
export type VisioCollaborationOptions = Omit<CollaborationOptions<VisioPresence>, 'session'> &
	VisioWriteBackOptions & {
		session?: CollabSession<VisioPresence> | CollabSession;
	};
