/**
 * collaboration-teardown.ts: make a departing peer actually leave the room.
 *
 * Every binding tears its collaboration session down on component unmount, but
 * that path never runs when the *document* goes away: closing the tab,
 * navigating away, or (the case that motivated this module) an embedding page
 * removing the viewer's `<iframe>` outright. y-webrtc installs its own
 * `beforeunload` handler, but `beforeunload` is only fired by the
 * "prompt to unload" algorithm during a navigation; destroying a document by
 * detaching its frame fires `pagehide`/`unload` and never `beforeunload`. The
 * peer therefore keeps its awareness entry until the 30s awareness timeout,
 * which reads as a ghost collaborator in everyone else's presence list.
 *
 * This helper registers the missing listeners once per session:
 *  - `pagehide` (the reliable signal, and the only one mobile Safari honours),
 *  - `beforeunload` (belt and braces for the plain navigation case),
 *  - a `postMessage` from an embedding page (see {@link COLLAB_LEAVE_MESSAGE}),
 *    so a host that is about to destroy the frame can ask it to leave first
 *    instead of relying on the unload path alone.
 *
 * A `pagehide` with `persisted === true` means the document went into the
 * back/forward cache and may come back. We only leave in that case when the
 * caller supplied a `rejoin`, and then rejoin from `pageshow`.
 *
 * The leave message is obeyed only when the embedding window (or the frame
 * itself) sends it from an allowed origin: the frame's own origin unless the
 * host names others in `CollaborationConfig.leaveMessageOrigin`. The listeners
 * are core's `registerTeardown` (`ooxml-core/collab`) with the pptx message name.
 *
 * The window is injected (structurally typed) so bindings can unit-test the
 * behaviour without a DOM.
 */

import {
	LEGACY_PPTX_LEAVE_MESSAGE,
	registerTeardown,
	type TeardownEventLike,
	type TeardownListener,
	type TeardownOptions,
	type TeardownWindowLike,
} from 'ooxml-core/collab';

export type { TeardownEventLike, TeardownListener, TeardownWindowLike };

/**
 * `postMessage` payload an embedding page can send into a viewer frame to make
 * it leave its collaboration room before the frame is destroyed:
 * `frame.contentWindow?.postMessage({ type: COLLAB_LEAVE_MESSAGE }, viewerOrigin)`.
 */
export const COLLAB_LEAVE_MESSAGE = LEGACY_PPTX_LEAVE_MESSAGE;

export type CollaborationTeardownOptions = Omit<TeardownOptions, 'leaveMessage'>;

/**
 * Register document-teardown listeners that make the local peer leave its
 * collaboration room. Returns a disposer that removes them again; call it from
 * the binding's normal unmount cleanup (which performs its own teardown).
 *
 * Safe to call in a non-browser environment: it becomes a no-op disposer.
 */
export function registerCollaborationTeardown(options: CollaborationTeardownOptions): () => void {
	return registerTeardown({ ...options, leaveMessage: COLLAB_LEAVE_MESSAGE });
}

/** The slice of a Yjs `Awareness` needed to withdraw the local presence. */
export interface AwarenessTeardownLike {
	setLocalState?: (state: null) => void;
}

/**
 * Withdraw the local awareness state so peers drop us immediately instead of
 * waiting out the 30s awareness timeout. Yjs broadcasts the resulting `removed`
 * update through the still-open transport, so call this BEFORE destroying the
 * provider.
 */
export function clearLocalAwareness(awareness: AwarenessTeardownLike | null | undefined): void {
	awareness?.setLocalState?.(null);
}
