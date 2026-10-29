// WebSocket transport (y-websocket compatible server protocol) with exponential-backoff reconnect.
// The WebSocket constructor is injectable, so it works in browsers, Bun, Node (`ws`) and tests, and
// the mixed-content check fails fast instead of waiting out a socket the browser will never open.
// New code written for the collab area; the reconnect policy mirrors y-websocket's.
import type { Transport, TransportHandlers } from './provider';
import { isMixedContentBlocked } from './validation';

/** The structural slice of `WebSocket` this transport uses. */
export interface WebSocketLike {
	binaryType: string;
	readyState: number;
	send: (data: Uint8Array) => void;
	close: () => void;
	onopen: ((event: unknown) => void) | null;
	onclose: ((event: { reason?: string; code?: number }) => void) | null;
	onerror: ((event: unknown) => void) | null;
	onmessage: ((event: { data: unknown }) => void) | null;
}
export type WebSocketConstructor = new (url: string, protocols?: string[]) => WebSocketLike;

export interface WebSocketTransportOptions {
	/** Full URL including the room, for example `wss://host/room-id` (see {@link roomUrl}). */
	url: string;
	WebSocket?: WebSocketConstructor;
	/** First reconnect delay in ms; doubles up to `maxDelayMs`. Default 250. */
	baseDelayMs?: number;
	maxDelayMs?: number;
	/** Give up after this many consecutive failed attempts. Default unlimited. */
	maxRetries?: number;
	/**
	 * Stop reconnecting after this many consecutive closes of `1008 update rejected: ...` (the server
	 * refuses the same local state each time). Default 3.
	 */
	maxRejectedRetries?: number;
	/** WebSocket subprotocols to offer (for example a credential that must stay out of the URL). */
	protocols?: string[] | undefined;
	/** Page protocol for the mixed-content check; defaults to the ambient location. */
	pageProtocol?: string | undefined;
}

const OPEN = 1;
/** WebSocket close code the reference server uses for policy refusals. */
const POLICY_VIOLATION = 1008;
const POLICY_BASE_DELAY_MS = 5000;
const POLICY_MAX_DELAY_MS = 60_000;
/** A connection that stayed up this long was healthy: its later refusal starts a fresh count. */
const POLICY_RESET_MS = 30_000;

/** Join a server base URL and a validated room id: `roomUrl('wss://h/collab', 'r1')`. */
export function roomUrl(server: string, roomId: string): string {
	return `${trimTrailingSlashes(server)}/${encodeURIComponent(roomId)}`;
}

export function createWebSocketTransport(options: WebSocketTransportOptions): Transport {
	const {
		baseDelayMs = 250,
		maxDelayMs = 10_000,
		maxRetries = Infinity,
		maxRejectedRetries = 3,
	} = options;
	let socket: WebSocketLike | null = null;
	let handlers: TransportHandlers | null = null;
	let timer: ReturnType<typeof setTimeout> | null = null;
	let failures = 0;
	let policyCloses = 0;
	let openedAt = 0;
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
		const ws = options.protocols?.length
			? new Impl(options.url, options.protocols)
			: new Impl(options.url);
		socket = ws;
		ws.binaryType = 'arraybuffer';
		ws.onopen = () => {
			// A server that accepts the handshake and then refuses us (1008) must not reset the backoff.
			if (policyCloses === 0) failures = 0;
			openedAt = Date.now();
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
			const policy = event.code === POLICY_VIOLATION;
			if (!policy || Date.now() - openedAt > POLICY_RESET_MS) policyCloses = 0;
			if (policy) policyCloses++;
			if (++failures > maxRetries) {
				stopped = true;
				return active.error(new Error('WebSocket reconnect attempts exhausted'));
			}
			let delay = Math.min(maxDelayMs, baseDelayMs * 2 ** (failures - 1));
			if (policy) {
				// The server closed on purpose (rejected update, rate limit, revoked session): say why,
				// and wait longer each time. A rejected update would be sent again on every reconnect.
				const reason = event.reason || 'connection refused by the server';
				const rejected = reason.startsWith('update rejected');
				if (rejected && policyCloses >= maxRejectedRetries) {
					stopped = true;
					return active.error(
						new Error(`The server rejected your changes and sync stopped: ${reason}`),
					);
				}
				active.error(new Error(`The server closed the connection: ${reason}`));
				delay = Math.max(
					delay,
					Math.min(POLICY_MAX_DELAY_MS, POLICY_BASE_DELAY_MS * 2 ** (policyCloses - 1)),
				);
			}
			timer = setTimeout(open, delay);
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

function trimTrailingSlashes(value: string): string {
	let end = value.length;
	while (end > 0 && value.charCodeAt(end - 1) === 47) end--;
	return value.slice(0, end);
}
