// Structural views of the y-protocols `Awareness`, a throttled presence publisher, a borrowed
// awareness lease for host-owned sessions, and the helpers that make a departing peer disappear.
// Extracted from pptx-viewer `collaboration-presence-publisher.ts`, `collaboration-external-session.ts`
// and `collaboration-departure.ts` / `collaboration-teardown.ts`; see PROVENANCE.md.
import { type PresenceIdentity, BROADCAST_THROTTLE_MS, PRESENCE_FIELD } from './presence.js';

/** The slice of y-protocols `Awareness` the publisher needs; the real class satisfies it. */
export interface AwarenessLike {
	clientID?: number;
	setLocalStateField: (field: string, value: unknown) => void;
	getStates: () => Map<number, Record<string, unknown>>;
	on: (event: string, cb: (...args: never[]) => void) => void;
	off?: (event: string, cb: (...args: never[]) => void) => void;
}

export interface PresencePublisher<S extends object> {
	/** Merge a partial state patch and publish (throttled; a no-op patch is dropped). */
	update: (patch: Partial<S>) => void;
	/** Re-publish immediately (heartbeat). Never dropped, or peers would time us out. */
	flush: () => void;
	/** Cancel any pending throttled publish. */
	dispose: () => void;
}

export interface PresencePublisherOptions {
	throttleMs?: number;
	now?: () => number;
}

/**
 * Throttled writer of the local `presence` awareness field: identity (fixed) merged with the
 * product payload `S` (cursor, selection, page...). Leading edge plus a trailing flush; patches that
 * change nothing never reach the wire (they would re-stamp `lastUpdated` and re-render every peer).
 * The initial presence is announced synchronously.
 */
export function createPresencePublisher<S extends object>(
	awareness: AwarenessLike,
	identity: PresenceIdentity,
	initial: S,
	options: PresencePublisherOptions = {},
): PresencePublisher<S> {
	const throttleMs = options.throttleMs ?? BROADCAST_THROTTLE_MS;
	const clock = options.now ?? Date.now;
	const local: S = { ...initial };
	let lastBroadcast = 0;
	let pending: ReturnType<typeof setTimeout> | null = null;

	const publish = (): void => {
		lastBroadcast = clock();
		awareness.setLocalStateField(PRESENCE_FIELD, {
			...identity,
			...local,
			lastUpdated: new Date(clock()).toISOString(),
		});
	};
	const cancel = (): void => {
		if (pending !== null) clearTimeout(pending);
		pending = null;
	};
	const update = (patch: Partial<S>): void => {
		const keys = Object.keys(patch) as (keyof S)[];
		if (!keys.some((key) => patch[key] !== local[key])) return;
		Object.assign(local, patch);
		const elapsed = clock() - lastBroadcast;
		if (elapsed >= throttleMs) {
			cancel();
			publish();
		} else if (pending === null) {
			pending = setTimeout(() => {
				pending = null;
				publish();
			}, throttleMs - elapsed);
		}
	};
	publish();
	lastBroadcast = 0; // the announce does not consume the throttle budget
	return {
		update,
		flush: () => {
			cancel();
			publish();
		},
		dispose: cancel,
	};
}

/** The public y-protocols Awareness surface a borrowed session needs. */
export interface HostAwareness extends AwarenessLike {
	clientID: number;
	getLocalState: () => Record<string, unknown> | null;
	setLocalState: (state: Record<string, unknown> | null) => void;
	off: (event: string, cb: (...args: never[]) => void) => void;
}

/** Listener/publisher view of a host awareness, deliberately without lifecycle methods. */
export interface BorrowedAwareness extends AwarenessLike {
	clientID: number;
	off: (event: string, cb: (...args: never[]) => void) => void;
}

/**
 * Give a presence publisher a borrowed view of a host-owned awareness. Only the `presence` field
 * last written through the lease is restored on dispose; unrelated host fields and later host
 * presence writes are preserved, and a departed host is never brought back by viewer cleanup.
 */
export function borrowAwareness(host: HostAwareness): {
	awareness: BorrowedAwareness;
	dispose: () => void;
} {
	let disposed = false;
	let ownsPresence = false;
	let hadPresence = false;
	let previousPresence: unknown;
	let lastPresence: unknown;
	const awareness: BorrowedAwareness = {
		clientID: host.clientID,
		getStates: () => host.getStates(),
		on: (event, callback) => host.on(event, callback),
		off: (event, callback) => host.off(event, callback),
		setLocalStateField: (field, value) => {
			if (disposed) return;
			const state = host.getLocalState();
			if (state === null) return;
			if (field === PRESENCE_FIELD) {
				if (!ownsPresence || state[PRESENCE_FIELD] !== lastPresence) {
					hadPresence = Object.hasOwn(state, PRESENCE_FIELD);
					previousPresence = state[PRESENCE_FIELD];
				}
				lastPresence = value;
				ownsPresence = true;
			}
			host.setLocalStateField(field, value);
		},
	};
	return {
		awareness,
		dispose: () => {
			if (disposed) return;
			disposed = true;
			const current = host.getLocalState();
			if (!ownsPresence || current === null || current[PRESENCE_FIELD] !== lastPresence) return;
			const restored = { ...current };
			if (hadPresence) restored[PRESENCE_FIELD] = previousPresence;
			else delete restored[PRESENCE_FIELD];
			host.setLocalState(restored);
		},
	};
}

/** Withdraw the local awareness state so peers drop us at once (call before destroying the provider). */
export function clearLocalAwareness(
	awareness: { setLocalState?: (state: null) => void } | null | undefined,
): void {
	awareness?.setLocalState?.(null);
}

/** The slice of an `Awareness` needed to drop a departed peer locally. */
export interface AwarenessStatesLike {
	clientID?: number;
	states?: Map<number, Record<string, unknown>>;
	emit?: (name: string, args: unknown[]) => void;
}

/**
 * Drop `clients` from `awareness` locally and notify its observers, mirroring y-protocols'
 * `removeAwarenessStates` for remote clients. Returns the ids actually removed.
 */
export function removeAwarenessStatesLocally(
	awareness: AwarenessStatesLike | null | undefined,
	clients: readonly number[],
): number[] {
	const states = awareness?.states;
	if (!awareness || !states) return [];
	const removed = clients.filter((clientId) => states.delete(clientId));
	if (removed.length > 0) {
		const change = { added: [], updated: [], removed };
		awareness.emit?.('change', [change, 'peer-departed']);
		awareness.emit?.('update', [change, 'peer-departed']);
	}
	return removed;
}
