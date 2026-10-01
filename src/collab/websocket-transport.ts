// WebSocket transport (y-websocket compatible server protocol) with exponential-backoff reconnect.
// The WebSocket constructor is injectable, so it works in browsers, Bun, Node (`ws`) and tests, and
// the mixed-content check fails fast instead of waiting out a socket the browser will never open.
// New code written for the collab area; the reconnect policy mirrors y-websocket's.
import type { Transport, TransportHandlers } from './provider.js';
import { isMixedContentBlocked } from './validation.js';

/** The structural slice of `WebSocket` this transport uses. */
export interface WebSocketLike {
	binaryType: string;
	readyState: number;
	send: (data: Uint8Array) => void;
	close: () => void;
	onopen: ((event: unknown) => void) | null;
	onclose: ((event: { reason?: string }) => void) | null;
	onerror: ((event: unknown) => void) | null;
	onmessage: ((event: { data: unknown }) => void) | null;
}
export type WebSocketConstructor = new (url: string) => WebSocketLike;

export interface WebSocketTransportOptions {
	/** Full URL including the room, for example `wss://host/room-id` (see {@link roomUrl}). */
	url: string;
	WebSocket?: WebSocketConstructor;
	/** First reconnect delay in ms; doubles up to `maxDelayMs`. Default 250. */
	baseDelayMs?: number;
	maxDelayMs?: number;
	/** Give up after this many consecutive failed attempts. Default unlimited. */
	maxRetries?: number;
	/** Page protocol for the mixed-content check; defaults to the ambient location. */
	pageProtocol?: string | undefined;
}

const OPEN = 1;

/** Join a server base URL and a validated room id: `roomUrl('wss://h/collab', 'r1')`. */
export function roomUrl(server: string, roomId: string): string {
	return `${server.replace(/\/+$/u, '')}/${encodeURIComponent(roomId)}`;
}

export function createWebSocketTransport(options: WebSocketTransportOptions): Transport {
	const { baseDelayMs = 250, maxDelayMs = 10_000, maxRetries = Infinity } = options;
	let socket: WebSocketLike | null = null;
	let handlers: TransportHandlers | null = null;
	let timer: ReturnType<typeof setTimeout> | null = null;
	let failures = 0;
	let stopped = true;

	const open = (): void => {
		if (stopped || !handlers) return;
		const active = handlers;
		const Impl =
			options.WebSocket ?? (globalThis as { WebSocket?: WebSocketConstructor }).WebSocket;
		if (!Impl) return active.error(new Error('No WebSocket implementation available'));
		if (isMixedContentBlocked(options.url, options.pageProtocol)) {
			stopped = true;
			return active.error(new Error('Insecure ws:// connection blocked from an https page'));
		}
		const ws = new Impl(options.url);
		socket = ws;
		ws.binaryType = 'arraybuffer';
		ws.onopen = () => {
			failures = 0;
			active.open();
		};
		ws.onmessage = (event) => {
			const data = event.data;
			if (data instanceof ArrayBuffer) active.message(new Uint8Array(data));
			else if (data instanceof Uint8Array) active.message(data);
			else active.error(new Error('Received a non-binary WebSocket message'));
		};
		ws.onerror = () => active.error(new Error('WebSocket error'));
		ws.onclose = (event) => {
			if (socket !== ws) return;
			socket = null;
			active.close(event.reason);
			if (stopped) return;
			if (++failures > maxRetries) {
				stopped = true;
				return active.error(new Error('WebSocket reconnect attempts exhausted'));
			}
			timer = setTimeout(open, Math.min(maxDelayMs, baseDelayMs * 2 ** (failures - 1)));
		};
	};

	return {
		connect(next) {
			handlers = next;
			stopped = false;
			failures = 0;
			if (timer !== null) clearTimeout(timer);
			timer = null;
			if (socket) socket.close();
			open();
		},
		send(data) {
			if (socket && socket.readyState === OPEN) socket.send(data);
		},
		disconnect() {
			stopped = true;
			if (timer !== null) clearTimeout(timer);
			timer = null;
			const ws = socket;
			socket = null;
			if (ws) {
				ws.onclose = null;
				ws.close();
				handlers?.close('disconnected');
			}
		},
	};
}
