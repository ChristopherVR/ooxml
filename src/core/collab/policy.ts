// Small session policies shared by every viewer: the first-write gate, who wins when a loaded
// document meets a non-empty room, and broadcast auto-follow. Extracted from pptx-viewer
// `collaboration-sync-gate.ts`, `collaboration-load-origin.ts` and `collaboration-broadcast-follow.ts`;
// see PROVENANCE.md.
import type { CollaborationRole } from './identity.js';

/**
 * Grace period (ms) before the first local write when the provider has not signalled initial sync.
 * WebSocket providers emit `synced` reliably; WebRTC only syncs once a peer is present, so a lone
 * fresh-room peer seeds the document after this delay. Gating the first write prevents a late
 * joiner's bootstrap document from merging into a room whose real content has not arrived.
 */
export const INITIAL_SYNC_GRACE_MS = 3_000;

export interface SyncGate {
	isOpen: () => boolean;
	/** (Re)start the grace timer that opens the gate without a sync event. */
	arm: () => void;
	/** Open now (the provider confirmed its initial sync). Idempotent. */
	open: () => void;
	/** Close and cancel the grace timer (teardown). */
	reset: () => void;
}

/** First-write gate. `onOpen` fires once per session (until `reset`). */
export function createSyncGate(
	onOpen: () => void,
	graceMs: number = INITIAL_SYNC_GRACE_MS,
): SyncGate {
	let opened = false;
	let timer: ReturnType<typeof setTimeout> | null = null;
	const clear = (): void => {
		if (timer !== null) clearTimeout(timer);
		timer = null;
	};
	const open = (): void => {
		clear();
		if (opened) return;
		opened = true;
		onOpen();
	};
	return {
		isOpen: () => opened,
		arm: () => {
			clear();
			if (!opened) timer = setTimeout(open, graceMs);
		},
		open,
		reset: () => {
			clear();
			opened = false;
		},
	};
}

/**
 * Why a content load ran: `bootstrap` (the document the host mounted the viewer with, or a session
 * restore; nobody chose it during the session) or `user` (File > Open, a drop, a host API call).
 */
export type LoadOrigin = 'bootstrap' | 'user';

/**
 * Whether the room's content should replace the document that just loaded. Only a bootstrap load
 * into a non-empty room loses: opening a file is an act of authorship and is published instead.
 * An unknown origin defaults to `user`, which keeps the document rather than dropping it.
 */
export function shouldRoomReplaceLoad(
	origin: LoadOrigin | undefined,
	roomIsNonEmpty: boolean,
): boolean {
	return (origin ?? 'user') === 'bootstrap' && roomIsNonEmpty;
}

/**
 * Broadcast sessions have one presenter (`owner`). Auto-follow is one-way: only a local `viewer` is
 * pulled along to the owner's location; collaborators and the owner navigate freely.
 */
export function shouldAutoFollowBroadcaster(input: {
	localRole: CollaborationRole | undefined;
	broadcasterRole: CollaborationRole | undefined;
}): boolean {
	return input.localRole === 'viewer' && input.broadcasterRole === 'owner';
}
