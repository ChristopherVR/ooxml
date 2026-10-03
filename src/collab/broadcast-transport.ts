// Same-browser mesh transport over BroadcastChannel: every tab or window of one origin that joins
// the same room hears every other, with no server and nothing leaving the device. Useful for
// local-first viewers (documents stay local) and demos. New code written for the collab area,
// following the mesh contract of memory-transport.ts.
import type { Transport, TransportHandlers } from './provider.js';
import { validateRoomId } from './validation.js';

/** Structural BroadcastChannel, so the area needs no DOM types. */
export interface BroadcastChannelLike {
	postMessage: (message: unknown) => void;
	close: () => void;
	onmessage: ((event: { data: unknown }) => void) | null;
}
export type BroadcastChannelConstructor = new (name: string) => BroadcastChannelLike;

export interface BroadcastTransportOptions {
	roomId: string;
	/** Channel name prefix; one channel per room. Default `ooxml-collab:`. */
	prefix?: string;
	/** Defaults to `globalThis.BroadcastChannel`. */
	BroadcastChannel?: BroadcastChannelConstructor;
}

type Envelope = { kind: 'hello' | 'data'; from: string; data?: Uint8Array };

/** Realm-independent: structured clone may deliver another realm's Uint8Array. */
const isBytes = (value: unknown): value is Uint8Array =>
	ArrayBuffer.isView(value) && Object.prototype.toString.call(value) === '[object Uint8Array]';

const isEnvelope = (value: unknown): value is Envelope => {
	const envelope = value as Partial<Envelope> | null;
	return (
		!!envelope &&
		typeof envelope.from === 'string' &&
		(envelope.kind === 'hello' || (envelope.kind === 'data' && isBytes(envelope.data)))
	);
};

/** Whether this environment can create a BroadcastChannel transport. */
export function canBroadcast(scope: object = globalThis): boolean {
	return typeof (scope as { BroadcastChannel?: unknown }).BroadcastChannel === 'function';
}

export function createBroadcastTransport(options: BroadcastTransportOptions): Transport {
	const roomId = validateRoomId(options.roomId);
	const Channel =
		options.BroadcastChannel ??
		(globalThis as { BroadcastChannel?: BroadcastChannelConstructor }).BroadcastChannel;
	const self = Math.random().toString(36).slice(2) + Date.now().toString(36);
	let channel: BroadcastChannelLike | null = null;
	let handlers: TransportHandlers | null = null;

	const close = (): void => {
		if (!channel) return;
		channel.onmessage = null;
		channel.close();
		channel = null;
	};
	return {
		connect(next) {
			close();
			handlers = next;
			if (!Channel) {
				next.error(new Error('BroadcastChannel is not available in this environment.'));
				next.close('unavailable');
				return;
			}
			const open = new Channel(`${options.prefix ?? 'ooxml-collab:'}${roomId}`);
			channel = open;
			open.onmessage = ({ data }) => {
				if (channel !== open || !isEnvelope(data) || data.from === self) return;
				// A newcomer announced itself: the provider answers with a sync handshake and its
				// awareness. The newcomer's own opening handshake reaches us through `send`.
				if (data.kind === 'hello') handlers?.peer?.();
				else handlers?.message(new Uint8Array(data.data!));
			};
			next.open();
			open.postMessage({ kind: 'hello', from: self } satisfies Envelope);
		},
		send(data) {
			try {
				channel?.postMessage({ kind: 'data', from: self, data } satisfies Envelope);
			} catch (cause) {
				handlers?.error(cause instanceof Error ? cause : new Error(String(cause)));
			}
		},
		disconnect() {
			const was = handlers;
			const open = channel !== null;
			close();
			handlers = null;
			if (open) was?.close('disconnected');
		},
	};
}
