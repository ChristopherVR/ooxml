// The stock provider: Yjs document sync plus awareness over any {@link Transport}, using the
// y-websocket wire format (message type 0 = sync, 1 = awareness; sync sub-types step1/step2/update),
// so a stock y-websocket server works with a WebSocket transport. Inbound payloads are decoded
// defensively: a malformed or oversized message is reported as an `error` event and dropped, never
// thrown, and updates are validated before they touch the document (see codec.ts).
// New code written for the collab area; protocol constants follow y-protocols.
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as awarenessProtocol from 'y-protocols/awareness';
import * as Y from 'yjs';
import { applyUpdateSafe, MAX_UPDATE_BYTES } from './codec.js';
import { Emitter } from './emitter.js';
import type {
	ConnectionStatus,
	ProviderContext,
	ProviderEvents,
	SyncProvider,
	Transport,
} from './provider.js';

export const MESSAGE_SYNC = 0;
export const MESSAGE_AWARENESS = 1;
export const MESSAGE_QUERY_AWARENESS = 3;
const SYNC_STEP1 = 0;
const SYNC_STEP2 = 1;
const SYNC_UPDATE = 2;

export interface TransportProviderOptions {
	transport: Transport;
	/** Reject inbound messages larger than this. Default 32 MiB. */
	maxMessageBytes?: number;
}

/** Create the provider factory for a session: `provider: transportProvider({ transport })`. */
export function transportProvider(
	options: TransportProviderOptions,
): (context: ProviderContext) => SyncProvider {
	return (context) => createTransportProvider(context, options);
}

export function createTransportProvider(
	{ doc, awareness }: ProviderContext,
	{ transport, maxMessageBytes = MAX_UPDATE_BYTES }: TransportProviderOptions,
): SyncProvider {
	const events = new Emitter<ProviderEvents>();
	let status: ConnectionStatus = 'disconnected';
	let synced = false;
	let open = false;
	let destroyed = false;
	const origin = {}; // identifies our own writes so they are not echoed back

	const setStatus = (next: ConnectionStatus): void => {
		if (status === next) return;
		status = next;
		events.emit('status', next);
	};
	const setSynced = (next: boolean): void => {
		if (synced === next) return;
		synced = next;
		events.emit('synced', next);
	};
	const fail = (reason: string): void => events.emit('error', new Error(reason));

	const send = (encoder: encoding.Encoder): void => {
		if (open) transport.send(encoding.toUint8Array(encoder));
	};
	const sendStep1 = (): void => {
		const encoder = encoding.createEncoder();
		encoding.writeVarUint(encoder, MESSAGE_SYNC);
		encoding.writeVarUint(encoder, SYNC_STEP1);
		encoding.writeVarUint8Array(encoder, Y.encodeStateVector(doc));
		send(encoder);
	};
	const sendAwareness = (clients: number[]): void => {
		if (clients.length === 0) return;
		const encoder = encoding.createEncoder();
		encoding.writeVarUint(encoder, MESSAGE_AWARENESS);
		encoding.writeVarUint8Array(
			encoder,
			awarenessProtocol.encodeAwarenessUpdate(awareness, clients),
		);
		send(encoder);
	};
	const handshake = (): void => {
		sendStep1();
		// Re-setting the state bumps the awareness clock: peers that already recorded our departure
		// (clock + 1) would otherwise ignore the announcement. The update event broadcasts it.
		const local = awareness.getLocalState();
		if (local !== null) awareness.setLocalState(local);
	};

	const onDocUpdate = (update: Uint8Array, updateOrigin: unknown): void => {
		if (updateOrigin === origin || !open) return;
		const encoder = encoding.createEncoder();
		encoding.writeVarUint(encoder, MESSAGE_SYNC);
		encoding.writeVarUint(encoder, SYNC_UPDATE);
		encoding.writeVarUint8Array(encoder, update);
		send(encoder);
	};
	const onAwarenessUpdate = (
		change: { added: number[]; updated: number[]; removed: number[] },
		updateOrigin: unknown,
	): void => {
		if (updateOrigin === origin) return;
		sendAwareness([...change.added, ...change.updated, ...change.removed]);
	};

	const readSync = (decoder: decoding.Decoder): void => {
		const kind = decoding.readVarUint(decoder);
		const payload = decoding.readVarUint8Array(decoder);
		if (kind === SYNC_STEP1) {
			const encoder = encoding.createEncoder();
			encoding.writeVarUint(encoder, MESSAGE_SYNC);
			encoding.writeVarUint(encoder, SYNC_STEP2);
			let diff: Uint8Array;
			try {
				diff = Y.encodeStateAsUpdate(doc, payload);
			} catch {
				return fail('malformed state vector');
			}
			encoding.writeVarUint8Array(encoder, diff);
			return send(encoder);
		}
		if (kind !== SYNC_STEP2 && kind !== SYNC_UPDATE) return fail(`unknown sync message ${kind}`);
		const result = applyUpdateSafe(doc, payload, origin, { maxBytes: maxMessageBytes });
		if (!result.ok) return fail(result.reason);
		if (kind === SYNC_STEP2) setSynced(true);
	};

	const onMessage = (data: Uint8Array): void => {
		if (destroyed) return;
		if (!(data instanceof Uint8Array) || data.byteLength === 0 || data.byteLength > maxMessageBytes)
			return fail('message is empty, not binary, or too large');
		try {
			const decoder = decoding.createDecoder(data);
			const type = decoding.readVarUint(decoder);
			if (type === MESSAGE_SYNC) readSync(decoder);
			else if (type === MESSAGE_AWARENESS)
				awarenessProtocol.applyAwarenessUpdate(
					awareness,
					decoding.readVarUint8Array(decoder),
					origin,
				);
			else if (type === MESSAGE_QUERY_AWARENESS) sendAwareness([...awareness.getStates().keys()]);
			// Other types (auth, custom) are ignored for forward compatibility.
		} catch (cause) {
			fail(`malformed message: ${cause instanceof Error ? cause.message : String(cause)}`);
		}
	};

	const dropRemotePeers = (): void => {
		const remote = [...awareness.getStates().keys()].filter((id) => id !== doc.clientID);
		if (remote.length > 0) awarenessProtocol.removeAwarenessStates(awareness, remote, origin);
	};

	const connect = (): void => {
		if (destroyed) return;
		setStatus('connecting');
		transport.connect({
			open: () => {
				if (destroyed) return;
				open = true;
				setStatus('connected');
				handshake();
			},
			close: () => {
				open = false;
				setSynced(false);
				dropRemotePeers();
				if (!destroyed) setStatus('disconnected');
			},
			message: onMessage,
			error: (error) => {
				events.emit('error', error);
				if (!open) setStatus('error');
			},
			peer: () => {
				if (open && !destroyed) handshake();
			},
		});
	};

	doc.on('update', onDocUpdate);
	awareness.on('update', onAwarenessUpdate);

	return {
		get status() {
			return status;
		},
		get synced() {
			return synced;
		},
		connect,
		disconnect: () => {
			if (destroyed) return;
			// Tell peers we are gone before the channel closes.
			const local = awareness.getLocalState();
			if (open && local !== null) awareness.setLocalState(null); // broadcasts the removal
			transport.disconnect();
			open = false;
			if (local !== null) awareness.setLocalState(local); // restored for the next connect
			setSynced(false);
			dropRemotePeers();
			setStatus('disconnected');
		},
		destroy: () => {
			if (destroyed) return;
			if (open) awareness.setLocalState(null);
			destroyed = true;
			open = false;
			doc.off('update', onDocUpdate);
			awareness.off('update', onAwarenessUpdate);
			transport.disconnect();
			events.clear();
		},
		on: (event, listener) => events.on(event, listener),
	};
}
