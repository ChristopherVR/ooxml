// WebRTC signaling: the JSON messages peers exchange to set up a call, a defensive parser for
// them, and the pluggable `SignalingChannel` seam with an in-memory hub, a same-browser
// BroadcastChannel mesh and a WebSocket client for a bring-your-own relay server. The relay never
// needs to understand a message: it forwards JSON to the other sockets of the room (see
// docs/teams-area.md for the three-rule server contract). DOM-free. New code.
import { isValidId } from '../collab/validation';
import type {
	BroadcastChannelConstructor,
	BroadcastChannelLike,
} from '../collab/broadcast-transport';
import type { WebSocketConstructor } from '../collab/websocket-transport';

export interface SessionDescription {
	type: 'offer' | 'answer';
	sdp: string;
}

export interface IceCandidateInit {
	candidate: string;
	sdpMid?: string | null;
	sdpMLineIndex?: number | null;
	usernameFragment?: string | null;
}

export interface MediaState {
	audio: boolean;
	video: boolean;
	screen: boolean;
	hand: boolean;
}

export type Signal =
	| { type: 'hello'; from: string; name: string; to?: string; reply?: boolean; state: MediaState }
	| { type: 'bye'; from: string }
	| { type: 'description'; from: string; to: string; description: SessionDescription }
	| { type: 'candidate'; from: string; to: string; candidate: IceCandidateInit }
	| { type: 'state'; from: string; state: MediaState };

export const MAX_SDP_CHARS = 200_000;
export const MAX_SIGNAL_CHARS = 262_144;

const bool = (v: unknown): boolean => v === true;
const parseState = (raw: unknown): MediaState => {
	const r = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>;
	return { audio: bool(r.audio), video: bool(r.video), screen: bool(r.screen), hand: bool(r.hand) };
};

/** Validate one inbound signal. Anything malformed, oversized or mistyped returns null. */
export function parseSignal(input: unknown): Signal | null {
	let raw: unknown = input;
	if (typeof input === 'string') {
		if (input.length > MAX_SIGNAL_CHARS) return null;
		try {
			raw = JSON.parse(input);
		} catch {
			return null;
		}
	}
	if (!raw || typeof raw !== 'object') return null;
	const r = raw as Record<string, unknown>;
	if (!isValidId(r.from, 128)) return null;
	const from = r.from;
	const to = isValidId(r.to, 128) ? r.to : undefined;
	switch (r.type) {
		case 'hello': {
			const name = typeof r.name === 'string' ? r.name.slice(0, 80) : '';
			if (!name) return null;
			const out: Signal = { type: 'hello', from, name, state: parseState(r.state) };
			if (to) out.to = to;
			if (r.reply === true) out.reply = true;
			return out;
		}
		case 'bye':
			return { type: 'bye', from };
		case 'state':
			return { type: 'state', from, state: parseState(r.state) };
		case 'description': {
			const d = r.description as Record<string, unknown> | undefined;
			if (!to || !d || (d.type !== 'offer' && d.type !== 'answer')) return null;
			if (typeof d.sdp !== 'string' || d.sdp.length > MAX_SDP_CHARS) return null;
			return { type: 'description', from, to, description: { type: d.type, sdp: d.sdp } };
		}
		case 'candidate': {
			const c = r.candidate as Record<string, unknown> | undefined;
			if (!to || !c || typeof c.candidate !== 'string' || c.candidate.length > 2048) return null;
			const candidate: IceCandidateInit = { candidate: c.candidate };
			if (typeof c.sdpMid === 'string' || c.sdpMid === null) candidate.sdpMid = c.sdpMid;
			if (typeof c.sdpMLineIndex === 'number' || c.sdpMLineIndex === null)
				candidate.sdpMLineIndex = c.sdpMLineIndex;
			if (typeof c.usernameFragment === 'string') candidate.usernameFragment = c.usernameFragment;
			return { type: 'candidate', from, to, candidate };
		}
		default:
			return null;
	}
}

export interface SignalingHandlers {
	open: () => void;
	close: (reason?: string) => void;
	message: (signal: Signal) => void;
	error: (error: Error) => void;
}

/** A room-scoped broadcast channel for signals. Delivery to `to` filtering is the caller's job. */
export interface SignalingChannel {
	connect: (handlers: SignalingHandlers) => void;
	send: (signal: Signal) => void;
	disconnect: () => void;
}

export interface SignalingHub {
	createChannel: (roomId: string) => SignalingChannel;
}

/** In-process signaling for tests and single-page demos. */
export function createMemorySignalingHub(): SignalingHub {
	const rooms = new Map<string, Set<SignalingHandlers>>();
	return {
		createChannel(roomId) {
			let mine: SignalingHandlers | null = null;
			const room = (): Set<SignalingHandlers> => {
				let set = rooms.get(roomId);
				if (!set) rooms.set(roomId, (set = new Set()));
				return set;
			};
			return {
				connect(handlers) {
					mine = handlers;
					room().add(handlers);
					queueMicrotask(() => handlers.open());
				},
				send(signal) {
					const wire = parseSignal(JSON.stringify(signal));
					if (!wire) return;
					for (const other of room()) if (other !== mine) queueMicrotask(() => other.message(wire));
				},
				disconnect() {
					if (mine) room().delete(mine);
					mine = null;
				},
			};
		},
	};
}

export interface BroadcastSignalingOptions {
	roomId: string;
	prefix?: string;
	BroadcastChannel?: BroadcastChannelConstructor;
}

/** Signaling across tabs of one origin: no server, nothing leaves the device. */
export function createBroadcastSignaling(options: BroadcastSignalingOptions): SignalingChannel {
	let channel: BroadcastChannelLike | null = null;
	return {
		connect(handlers) {
			const Impl =
				options.BroadcastChannel ??
				(globalThis as { BroadcastChannel?: BroadcastChannelConstructor }).BroadcastChannel;
			if (!Impl) return handlers.error(new Error('BroadcastChannel is not available'));
			channel = new Impl(`${options.prefix ?? 'teams-signal'}:${options.roomId}`);
			channel.onmessage = (event) => {
				const signal = parseSignal(event.data);
				if (signal) handlers.message(signal);
			};
			handlers.open();
		},
		send(signal) {
			channel?.postMessage(JSON.stringify(signal));
		},
		disconnect() {
			channel?.close();
			channel = null;
		},
	};
}

export interface WebSocketSignalingOptions {
	/** Full URL including the room, for example `wss://host/signal/room-1`. */
	url: string;
	WebSocket?: WebSocketConstructor;
	baseDelayMs?: number;
	maxDelayMs?: number;
}

/** Signaling over a bring-your-own WebSocket relay, reconnecting with exponential backoff. */
export function createWebSocketSignaling(options: WebSocketSignalingOptions): SignalingChannel {
	const { baseDelayMs = 250, maxDelayMs = 10_000 } = options;
	let socket: { send: (d: string) => void; close: () => void; readyState: number } | null = null;
	let timer: ReturnType<typeof setTimeout> | null = null;
	let failures = 0;
	let stopped = true;
	let active: SignalingHandlers | null = null;

	const open = (): void => {
		if (stopped || !active) return;
		const handlers = active;
		const Impl =
			options.WebSocket ?? (globalThis as { WebSocket?: WebSocketConstructor }).WebSocket;
		if (!Impl) return handlers.error(new Error('No WebSocket implementation available'));
		const ws = new Impl(options.url);
		socket = ws as unknown as typeof socket;
		ws.onopen = () => {
			failures = 0;
			handlers.open();
		};
		ws.onmessage = (event) => {
			const signal = parseSignal(event.data);
			if (signal) handlers.message(signal);
		};
		ws.onerror = () => handlers.error(new Error('Signaling socket error'));
		ws.onclose = (event) => {
			socket = null;
			handlers.close(event.reason);
			if (stopped) return;
			timer = setTimeout(open, Math.min(maxDelayMs, baseDelayMs * 2 ** failures++));
		};
	};

	return {
		connect(handlers) {
			active = handlers;
			stopped = false;
			open();
		},
		send(signal) {
			if (socket && socket.readyState === 1) socket.send(JSON.stringify(signal));
		},
		disconnect() {
			stopped = true;
			if (timer) clearTimeout(timer);
			timer = null;
			socket?.close();
			socket = null;
		},
	};
}
