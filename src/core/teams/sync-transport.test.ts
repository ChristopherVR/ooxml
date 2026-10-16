import { describe, expect, it } from 'vitest';
import type { WebSocketLike } from '../collab/websocket-transport';
import { parseServerConfig } from './server-config';
import { createSyncTransport } from './sync-transport';

function socketUrls(): { urls: string[]; WebSocket: new (url: string) => WebSocketLike } {
	const urls: string[] = [];
	class FakeSocket {
		binaryType = 'arraybuffer';
		readyState = 0;
		onopen: (() => void) | null = null;
		onclose: (() => void) | null = null;
		onerror: (() => void) | null = null;
		onmessage: ((event: { data: unknown }) => void) | null = null;
		constructor(url: string) {
			urls.push(url);
		}
		send(): void {}
		close(): void {}
	}
	return { urls, WebSocket: FakeSocket as unknown as new (url: string) => WebSocketLike };
}

const handlers = { open() {}, close() {}, message() {}, error() {} };

describe('createSyncTransport', () => {
	it('sends the signed session with every room, private chats included', () => {
		const { config } = parseServerConfig({
			mode: 'server',
			syncUrl: 'wss://h/sync',
			sessionToken: 'v1.a.b',
		});
		const { urls, WebSocket } = socketUrls();
		const transport = createSyncTransport(config, 'chat-1', { WebSocket }, { members: 'a,b' });
		transport.connect(handlers);
		transport.disconnect();
		const url = new URL(urls[0]!);
		expect(url.searchParams.get('session')).toBe('v1.a.b');
		expect(url.searchParams.get('members')).toBe('a,b');
	});

	it('sends no session without one', () => {
		const { config } = parseServerConfig({ mode: 'server', syncUrl: 'wss://h/sync' });
		const { urls, WebSocket } = socketUrls();
		const transport = createSyncTransport(config, 'room', { WebSocket });
		transport.connect(handlers);
		transport.disconnect();
		expect(new URL(urls[0]!).searchParams.has('session')).toBe(false);
	});
});
