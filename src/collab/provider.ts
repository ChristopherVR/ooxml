// The pluggable seams of the collab area: a byte `Transport` (WebSocket, WebRTC, BroadcastChannel,
// in-memory, anything that moves Uint8Array) and a `SyncProvider` (what a session talks to). The
// stock provider, {@link createTransportProvider}, speaks the y-websocket wire protocol over a
// Transport. Hosts that already run y-websocket/y-webrtc wrap their provider with
// `adaptYjsProvider` (external-provider.ts) instead.
import type * as Y from 'yjs';
import type { Awareness } from 'y-protocols/awareness';

/** Connection lifecycle shared by every provider. */
export type ConnectionStatus = 'disconnected' | 'connecting' | 'connected' | 'error';

/** Maximum time (ms) to wait for an initial connection before giving up. */
export const CONNECTION_TIMEOUT_MS = 30_000;

/** Callbacks a transport invokes; supplied by the provider on every `connect`. */
export interface TransportHandlers {
	/** The channel is usable (also after every automatic reconnect). */
	open: () => void;
	/** The channel closed; the transport may reconnect and call `open` again. */
	close: (reason?: string) => void;
	/** A binary message arrived. */
	message: (data: Uint8Array) => void;
	/** A non-fatal failure (refused socket, mixed content, send error). */
	error: (error: Error) => void;
	/**
	 * Mesh transports only: another peer joined the room. The provider answers with a sync handshake
	 * and its awareness so the newcomer learns our state (star transports get that from the server).
	 */
	peer?: () => void;
}

/** A bidirectional, ordered-enough byte channel. Reconnect policy belongs to the transport. */
export interface Transport {
	connect: (handlers: TransportHandlers) => void;
	/** Send to the room. Silently dropped while closed; the provider re-syncs on the next `open`. */
	send: (data: Uint8Array) => void;
	/** Close and stop any reconnect attempts. */
	disconnect: () => void;
}

export interface ProviderEvents {
	status: ConnectionStatus;
	/** True once the initial state exchange completed for the current connection. */
	synced: boolean;
	/** A rejected inbound message, a transport failure. The provider keeps running. */
	error: Error;
}

/** Everything a session needs from a sync backend. */
export interface SyncProvider {
	readonly status: ConnectionStatus;
	readonly synced: boolean;
	connect: () => void;
	disconnect: () => void;
	/** Release every listener and timer; the provider cannot be reused. */
	destroy: () => void;
	on: <K extends keyof ProviderEvents>(
		event: K,
		listener: (payload: ProviderEvents[K]) => void,
	) => () => void;
}

export interface ProviderContext {
	doc: Y.Doc;
	awareness: Awareness;
	roomId: string;
}

/** Factory a session calls once with the document and awareness it owns. */
export type ProviderFactory = (context: ProviderContext) => SyncProvider;
