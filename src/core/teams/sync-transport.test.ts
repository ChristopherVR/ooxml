import { describe, expect, it } from 'vitest';
import type { WebSocketLike } from '../collab/websocket-transport';
import { parseServerConfig } from './server-config';
import { createSyncTransport } from './sync-transport';

function socketUrls(): {
	urls: string[];
	protocols: (string[] | undefined)[];
	WebSocket: new (url: string, protocols?: string[]) => WebSocketLike;
} {
	const urls: string[] = [];
	const protocols: (string[] | undefined)[] = [];
	class FakeSocket {
		binaryType = 'arraybuffer';
		readyState = 0;
		onopen: (() => void) | null = null;
		onclose: (() => void) | null = null;
		onerror: (() => void) | null = null;
		onmessage: ((event: { data: unknown }) => void) | null = null;
		constructor(url: string, offered?: string[]) {
			urls.push(url);
			protocols.push(offered);
		}
		send(): void {}
		close(): void {}
	}
	return {
		urls,
		protocols,
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
});
