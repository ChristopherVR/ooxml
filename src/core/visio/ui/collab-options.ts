// The `collaboration` options of a Visio viewer: the shared embedding contract of
// `ooxml-core/collab` with the Visio presence payload and the owner's write-back hook. Types only,
// so the viewer bindings (React, Vue, Angular, Svelte, Solid, vanilla) take them from the contract.
import type { CollabSession, CollaborationOptions } from '../../collab/index';
import type { VisioPresence } from './collab-presence';

/** Write-back settings of {@link VisioCollaborationOptions}. */
export interface VisioWriteBackOptions {
	/**
	 * Called with the saved VSDX bytes, debounced, after local and remote changes, only in the
	 * window whose `user.role` is `owner`. Returning is not a durable-storage acknowledgement: the
	 * viewer cannot tell whether the host stored the bytes, and a pending write-back is flushed
	 * once when the viewer leaves the room, not retried.
	 */
	onWriteBack?: (bytes: Uint8Array) => void;
	/** Quiet time after the last change before `onWriteBack` fires. Default 5000. */
	writeBackDebounceMs?: number;
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
