// A host-owned collaboration session: the page creates the Yjs document, awareness and provider,
// keeps the room ticket fresh and decides readiness. The editor receives only `external` and never
// connects, reconnects or destroys anything itself.
import type { ExternalCollaborationSession } from 'pptx-vanilla-viewer';
import { Awareness } from 'y-protocols/awareness';
import { WebsocketProvider } from 'y-websocket';
import * as Y from 'yjs';

import { requestGrant } from './api';
import type { Grant } from './api';

/** The relay's close code when the host changed this account's access (see server/relay.ts). */
const ACCESS_CHANGED = 4403;
const REFRESH_BEFORE_MS = 60_000;

type Status = 'connecting' | 'connected' | 'disconnected' | 'error';
export interface Snapshot {
	status: Status;
	synced: boolean;
}

export interface ReviewSession {
	readonly grant: Grant;
	readonly external: ExternalCollaborationSession;
	destroy(): void;
}

/**
 * Open the review's room with a fresh grant. `onAccessChanged` runs once when the host changes the
 * account's access; the session is already read-only by then, and the page replaces it (and the
 * editor) with one opened under the new grant.
 */
export async function openReviewSession(
	reviewId: string,
	onAccessChanged: () => void,
): Promise<ReviewSession> {
	const grant = await requestGrant(reviewId);
	const doc = new Y.Doc();
	const awareness = new Awareness(doc);
	const relay = `${location.protocol === 'https:' ? 'wss' : 'ws'}://${location.host}/sync`;
	const provider = new WebsocketProvider(relay, reviewId, doc, {
		awareness,
		disableBc: true,
		params: { ticket: grant.ticket },
	});
	const listeners = new Set<() => void>();
	let snapshot: Snapshot = { status: 'connecting', synced: false };
	let status: Status = 'connecting';
	let accessChanged = false;
	let destroyed = false;
	let expiresAt = grant.expiresAt;
	let timer: ReturnType<typeof setTimeout> | undefined;

	const notify = (): void => {
		// Readiness is the host's decision: synced with the room and still holding the same access.
		const next = { status, synced: provider.synced && !accessChanged };
		if (next.status === snapshot.status && next.synced === snapshot.synced) return;
		snapshot = next;
		for (const listener of listeners) listener();
	};
	const changeAccess = (): void => {
		if (accessChanged) return;
		accessChanged = true;
		provider.disconnect();
		notify();
		onAccessChanged();
	};
	// A ticket only admits a connection, so it is refreshed for the next reconnect, not the current one.
	const refresh = async (): Promise<void> => {
		try {
			const next = await requestGrant(reviewId, true);
			if (destroyed) return;
			if (next.permission !== grant.permission) return changeAccess();
			provider.params.ticket = next.ticket;
			expiresAt = next.expiresAt;
			schedule();
		} catch {
			// Signed out or access removed: the page finds out on its next poll.
			if (!destroyed) timer = setTimeout(() => void refresh(), 30_000);
		}
	};
	const schedule = (): void => {
		clearTimeout(timer);
		timer = setTimeout(
			() => void refresh(),
			Math.max(0, expiresAt - Date.now() - REFRESH_BEFORE_MS),
		);
	};
	const onStatus = ({ status: next }: { status: string }): void => {
		status =
			next === 'connected' ? 'connected' : next === 'connecting' ? 'connecting' : 'disconnected';
		notify();
	};
	const onClose = (event: CloseEvent | null): void => {
		if (event?.code === ACCESS_CHANGED) return changeAccess();
		if (Date.now() > expiresAt - REFRESH_BEFORE_MS) void refresh();
	};
	provider.on('status', onStatus);
	provider.on('sync', notify);
	provider.on('connection-close', onClose);
	schedule();

	return {
		grant,
		external: {
			doc,
			awareness,
			getSnapshot: () => snapshot,
			subscribe(listener) {
				listeners.add(listener);
				return () => listeners.delete(listener);
			},
		},
		destroy(): void {
			if (destroyed) return;
			destroyed = true;
			clearTimeout(timer);
			listeners.clear();
			provider.off('status', onStatus);
			provider.off('sync', notify);
			provider.off('connection-close', onClose);
			provider.destroy();
			awareness.destroy();
			doc.destroy();
		},
	};
}
