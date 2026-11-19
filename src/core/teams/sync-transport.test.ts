import { describe, expect, it, vi } from 'vitest';
import type { WebSocketLike } from '../collab/websocket-transport';
import { parseServerConfig } from './server-config';
import { createSyncTransport } from './sync-transport';

function socketUrls(): {
	urls: string[];
	protocols: (string[] | undefined)[];
	sockets: WebSocketLike[];
	WebSocket: new (url: string, protocols?: string[]) => WebSocketLike;
} {
	const urls: string[] = [];
	const protocols: (string[] | undefined)[] = [];
	const sockets: WebSocketLike[] = [];
	class FakeSocket {
		binaryType = 'arraybuffer';
		readyState = 0;
		onopen: (() => void) | null = null;
		onclose: ((event: { code?: number }) => void) | null = null;
		onerror: (() => void) | null = null;
		onmessage: ((event: { data: unknown }) => void) | null = null;
		constructor(url: string, offered?: string[]) {
			urls.push(url);
			protocols.push(offered);
			sockets.push(this);
		}
		send(): void {}
		close(): void {}
	}
	return {
		urls,
		protocols,
		sockets,
		WebSocket: FakeSocket as unknown as new (url: string, protocols?: string[]) => WebSocketLike,
	};
}

const handlers = { open() {}, close() {}, message() {}, error() {} };

describe('createSyncTransport', () => {
	it('sends the signed session in the subprotocol with every room, private chats included', () => {
		const { config } = parseServerConfig({
			mode: 'server',
			syncUrl: 'wss://h/sync',
			token: 'tok-1',
			sessionToken: 'v1.a.b',
		});
		const { urls, protocols, WebSocket } = socketUrls();
		const transport = createSyncTransport(config, 'chat-1', { WebSocket }, { members: 'a,b' });
		transport.connect(handlers);
		transport.disconnect();
		const url = new URL(urls[0]!);
		expect(url.searchParams.has('session')).toBe(false);
		expect(url.searchParams.has('token')).toBe(false);
		expect(url.searchParams.get('members')).toBe('a,b');
		expect(protocols[0]).toEqual([
			'openteams',
			'openteams.token.tok-1',
			'openteams.session.v1.a.b',
		]);
	});

	it('falls back to the query string when a credential cannot be a subprotocol', () => {
		const { config } = parseServerConfig({
			mode: 'server',
			syncUrl: 'wss://h/sync',
			token: 'a b/c',
			sessionToken: 'v1.a.b',
		});
		const { urls, protocols, WebSocket } = socketUrls();
		const transport = createSyncTransport(config, 'room', { WebSocket });
		transport.connect(handlers);
		transport.disconnect();
		const url = new URL(urls[0]!);
		expect(url.searchParams.get('session')).toBe('v1.a.b');
		expect(url.searchParams.get('token')).toBe('a b/c');
		expect(protocols[0]).toBeUndefined();
	});

	it('sends no session without one', () => {
		const { config } = parseServerConfig({ mode: 'server', syncUrl: 'wss://h/sync' });
		const { urls, protocols, WebSocket } = socketUrls();
		const transport = createSyncTransport(config, 'room', { WebSocket });
		transport.connect(handlers);
		transport.disconnect();
		expect(new URL(urls[0]!).searchParams.has('session')).toBe(false);
		expect(protocols[0]).toBeUndefined();
	});

	it('fetches a host-issued token before the first connection and again before each reconnect', async () => {
		vi.useFakeTimers();
		try {
			const { config } = parseServerConfig({
				mode: 'server',
				syncUrl: 'wss://h/sync',
				token: 'old',
			});
			const { protocols, sockets, WebSocket } = socketUrls();
			const issued = ['tok-1', 'tok-2'];
			const token = vi.fn(async () => issued.shift() ?? 'none');
			const errors: string[] = [];
			const transport = createSyncTransport(config, 'room', { WebSocket }, undefined, token);
			transport.connect({ ...handlers, error: (e) => errors.push(e.message) });
			await vi.advanceTimersByTimeAsync(0);
			expect(protocols[0]).toEqual(['openteams', 'openteams.token.tok-1']);
			sockets[0]!.onclose!({});
			await vi.advanceTimersByTimeAsync(250);
			expect(token).toHaveBeenCalledTimes(2);
			expect(protocols[1]).toEqual(['openteams', 'openteams.token.tok-2']);
			transport.disconnect();
		} finally {
			vi.useRealTimers();
		}
	});

	it('reports a failed credential fetch and retries it without opening a socket', async () => {
		vi.useFakeTimers();
		try {
			const { config } = parseServerConfig({ mode: 'server', syncUrl: 'wss://h/sync' });
			const { urls, WebSocket } = socketUrls();
			const token = vi.fn(async () => {
				throw new Error('no credential');
			});
			const errors: string[] = [];
			const transport = createSyncTransport(config, 'room', { WebSocket }, undefined, token);
			transport.connect({ ...handlers, error: (e) => errors.push(e.message) });
			await vi.advanceTimersByTimeAsync(0);
			expect(errors).toEqual(['no credential']);
			await vi.advanceTimersByTimeAsync(250);
			expect(token).toHaveBeenCalledTimes(2);
			transport.disconnect();
			await vi.advanceTimersByTimeAsync(60_000);
			expect(token).toHaveBeenCalledTimes(2);
			expect(urls).toHaveLength(0);
		} finally {
			vi.useRealTimers();
		}
	});
});
