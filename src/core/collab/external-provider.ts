// Adapters for providers that are created and owned outside the collab area: the stock y-websocket
// `WebsocketProvider` / y-webrtc `WebrtcProvider` (wrapped as a {@link SyncProvider}) and a
// host-owned session whose lifecycle the viewer must only observe, never destroy.
// Extracted from pptx-viewer `collaboration-external-session.ts` (subscription semantics) and the
// per-binding provider wiring; see PROVENANCE.md.
import type * as Y from 'yjs';
import { Emitter } from './emitter.js';
import type { HostAwareness } from './awareness.js';
import type { ConnectionStatus, ProviderEvents, SyncProvider } from './provider.js';

/** The event surface shared by y-websocket and y-webrtc providers. */
export interface YjsProviderLike {
	on: (event: string, listener: (payload: never) => void) => void;
	off?: (event: string, listener: (payload: never) => void) => void;
	connect?: () => void;
	disconnect?: () => void;
	destroy?: () => void;
	/** y-websocket: socket currently open. */
	wsconnected?: boolean;
	/** y-websocket: `synced`; y-webrtc: `synced`. */
	synced?: boolean;
	/** y-webrtc: `connected`. */
	connected?: boolean;
	/** y-websocket: `wsconnecting`. */
	wsconnecting?: boolean;
}

/**
 * Wrap a y-websocket or y-webrtc provider as a {@link SyncProvider}. Status maps from the
 * provider's `status` events (`{status}` for y-websocket, `{connected}` for y-webrtc) and
 * `connection-error`; `synced` from `sync` (y-websocket) or `synced` (y-webrtc). The wrapper owns
 * the provider: `destroy` destroys it.
 */
export function adaptYjsProvider(provider: YjsProviderLike): SyncProvider {
	const events = new Emitter<ProviderEvents>();
	let status: ConnectionStatus =
		provider.wsconnected || provider.connected ? 'connected' : 'disconnected';
	let synced = Boolean(provider.synced);
	const subscriptions: [string, (payload: never) => void][] = [];
	const listen = <P>(name: string, handler: (payload: P) => void): void => {
		const wrapped = handler as (payload: never) => void;
		provider.on(name, wrapped);
		subscriptions.push([name, wrapped]);
	};
	const setStatus = (next: ConnectionStatus): void => {
		if (next === status) return;
		status = next;
		events.emit('status', next);
	};
	const setSynced = (next: boolean): void => {
		if (next === synced) return;
		synced = next;
		events.emit('synced', next);
	};
	listen<{ status?: string; connected?: boolean }>('status', (payload) => {
		if (typeof payload.connected === 'boolean')
			return setStatus(payload.connected ? 'connected' : 'disconnected');
		if (payload.status === 'connected' || payload.status === 'connecting')
			setStatus(payload.status);
		else if (payload.status === 'disconnected') {
			setSynced(false);
			setStatus('disconnected');
		}
	});
	listen<boolean | { synced?: boolean }>('sync', (value) => setSynced(Boolean(value)));
	listen<{ synced?: boolean }>('synced', (value) => setSynced(Boolean(value.synced)));
	listen<unknown>('connection-error', (cause) => {
		setStatus('error');
		events.emit('error', cause instanceof Error ? cause : new Error('Provider connection error'));
	});
	return {
		get status() {
			return status;
		},
		get synced() {
			return synced;
		},
		connect: () => {
			setStatus('connecting');
			provider.connect?.();
		},
		disconnect: () => provider.disconnect?.(),
		destroy: () => {
			for (const [name, handler] of subscriptions) provider.off?.(name, handler);
			subscriptions.length = 0;
			events.clear();
			provider.destroy?.();
		},
		on: (event, listener) => events.on(event, listener),
	};
}

/** Connection state a host reports for a borrowed session. */
export interface ExternalSnapshot {
	status: ConnectionStatus;
	/**
	 * The host has loaded the authoritative document and permits local edits. Never inferred from
	 * `status` and never opened by a timeout; keep true while offline to allow ordinary offline edits.
	 */
	synced: boolean;
}

/** Host-owned Yjs resources. The host creates, connects and destroys them. */
export interface ExternalSession {
	readonly doc: Y.Doc;
	/** Must belong to `doc`. One viewer publishes presence at a time. */
	readonly awareness: HostAwareness;
	getSnapshot: () => ExternalSnapshot;
	/** Subscribe to status or synced changes; returns an unsubscribe function. */
	subscribe: (listener: () => void) => () => void;
}

/**
 * Observe a host-owned session. Subscribes before reading so a sync that completes during
 * attachment is not lost; the initial snapshot is delivered synchronously, then only changes.
 */
export function observeExternalSession(
	session: ExternalSession,
	onSnapshot: (snapshot: ExternalSnapshot) => void,
): () => void {
	let active = true;
	let subscribing = true;
	let previous: ExternalSnapshot | undefined;
	let unsubscribe: (() => void) | undefined;
	const notify = (): void => {
		if (!active || subscribing) return;
		const snapshot = session.getSnapshot();
		if (previous?.status === snapshot.status && previous.synced === snapshot.synced) return;
		previous = { status: snapshot.status, synced: snapshot.synced };
		onSnapshot(previous);
	};
	const dispose = (): void => {
		if (!active) return;
		active = false;
		unsubscribe?.();
	};
	try {
		unsubscribe = session.subscribe(notify);
		subscribing = false;
		notify();
	} catch (error) {
		dispose();
		throw error;
	}
	return dispose;
}
