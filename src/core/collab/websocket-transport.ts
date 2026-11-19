// WebSocket transport (y-websocket compatible server protocol) with exponential-backoff reconnect.
// The WebSocket constructor is injectable, so it works in browsers, Bun, Node (`ws`) and tests, and
// the mixed-content check fails fast instead of waiting out a socket the browser will never open.
// New code written for the collab area; the reconnect policy mirrors y-websocket's.
import type { Transport, TransportHandlers } from './provider';
import { UpdateRefusedError } from './refusal';
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

/** Where one connection attempt goes. */
export interface WebSocketTarget {
	url: string;
	protocols?: string[] | undefined;
}

export interface WebSocketTransportOptions {
	/**
	 * Full URL including the room, for example `wss://host/room-id` (see {@link roomUrl}). Or a
	 * function called before every attempt, the first and each reconnect, that resolves the URL and
	 * subprotocols (it then replaces `protocols`), so a short-lived credential is fetched fresh. A
	 * rejection is reported through `error` and retried with the same backoff as a refused socket.
	 */
	url: string | (() => Promise<WebSocketTarget>);
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
/** Close code for a message the server will not take: too big, too complex, or a full document. */
const MESSAGE_TOO_BIG = 1009;
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
	/** Counts connects and disconnects, so a target resolved for an earlier one is dropped. */
	let generation = 0;

	const retry = (active: TransportHandlers, delay: number): void => {
		if (++failures > maxRetries) {
			stopped = true;
			return active.error(new Error('WebSocket reconnect attempts exhausted'));
		}
		timer = setTimeout(
			open,
			Math.max(delay, Math.min(maxDelayMs, baseDelayMs * 2 ** (failures - 1))),
		);
	};

	const open = (): void => {
		if (stopped || !handlers) return;
		const active = handlers;
		const { url } = options;
		if (typeof url === 'string') return dial(active, { url, protocols: options.protocols });
		const current = generation;
		url().then(
			(target) => {
				if (current === generation && !stopped) dial(active, target);
			},
			(error: unknown) => {
				if (current !== generation || stopped) return;
				active.error(error instanceof Error ? error : new Error(String(error)));
				retry(active, 0);
			},
		);
	};

	const dial = (active: TransportHandlers, target: WebSocketTarget): void => {
		const Impl =
			options.WebSocket ?? (globalThis as { WebSocket?: WebSocketConstructor }).WebSocket;
		if (!Impl) return active.error(new Error('No WebSocket implementation available'));
		if (isMixedContentBlocked(target.url, options.pageProtocol)) {
			stopped = true;
			return active.error(new Error('Insecure ws:// connection blocked from an https page'));
		}
		const ws = target.protocols?.length
			? new Impl(target.url, target.protocols)
			: new Impl(target.url);
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
			if (stopped) return active.close(event.reason);
			/** Give up: the reason comes before the close, so listeners know why the connection ends. */
			const stop = (error: Error): void => {
				stopped = true;
				active.error(error);
				active.close(event.reason);
			};
			if (event.code === MESSAGE_TOO_BIG) {
				// The same change would be sent again on every reconnect.
				const reason = event.reason || 'message too big';
				const code = /document too large/i.test(reason) ? 'server-full' : 'server-rejected';
				return stop(new UpdateRefusedError(code, reason));
			}
			const policy = event.code === POLICY_VIOLATION;
			if (!policy || Date.now() - openedAt > POLICY_RESET_MS) policyCloses = 0;
			if (policy) policyCloses++;
			// `retry` counts this failure; checking here first puts the reason before the close.
			if (failures >= maxRetries) return stop(new Error('WebSocket reconnect attempts exhausted'));
			// The server closed on purpose (rejected update, rate limit, revoked session): say why, and
			// wait longer each time. A rejected update would be sent again on every reconnect.
			const reason = event.reason || 'connection refused by the server';
			if (policy && reason.startsWith('update rejected') && policyCloses >= maxRejectedRetries)
				return stop(new UpdateRefusedError('server-rejected', reason));
			active.close(event.reason);
			let delay = 0;
			if (policy) {
				active.error(new Error(`The server closed the connection: ${reason}`));
				delay = Math.min(POLICY_MAX_DELAY_MS, POLICY_BASE_DELAY_MS * 2 ** (policyCloses - 1));
			}
			retry(active, delay);
		};
	};

	return {
		connect(next) {
			handlers = next;
			generation++;
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
			generation++;
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
