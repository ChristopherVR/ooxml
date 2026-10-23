import { describe, expect, it } from 'vitest';

import type { CollabTransportModules } from './collaboration-provider';
import { createCollabProvider } from './collaboration-provider';

function modules(): {
	calls: { ws: unknown[][]; webrtc: unknown[][] };
	modules: CollabTransportModules;
} {
	const calls = { ws: [] as unknown[][], webrtc: [] as unknown[][] };
	return {
		calls,
		modules: {
			loadWebsocket: async () => ({
				WebsocketProvider: class {
					awareness = { clientID: 1 };
					wsconnected = true;
					synced = false;
					constructor(...args: unknown[]) {
						calls.ws.push(args);
					}
					on() {}
					disconnect() {}
					destroy() {}
				},
			}),
			loadWebrtc: async () => ({
				WebrtcProvider: class {
					awareness = { clientID: 2 };
					constructor(...args: unknown[]) {
						calls.webrtc.push(args);
					}
					on() {}
					destroy() {}
				},
			}),
		},
	};
}

const doc = {};

describe('createCollabProvider', () => {
	it('creates a y-websocket provider and reports its live connection state', async () => {
		const { calls, modules: loaders } = modules();
		const handle = await createCollabProvider(
			'websocket',
			{ roomId: 'room', serverUrl: 'wss://x', userName: 'Ada', authToken: 'tok' },
			doc,
			loaders,
		);
		expect(calls.ws).toHaveLength(1);
		// (serverUrl, roomId, doc, { params: { token } })
		expect(calls.ws[0]?.[0]).toBe('wss://x');
		expect(calls.ws[0]?.[1]).toBe('room');
		expect(calls.ws[0]?.[3]).toStrictEqual({ params: { token: 'tok' } });
		expect(handle.connectedNow).toBeTruthy();
		expect(handle.syncedNow).toBeFalsy();
	});

	it('creates a serverless y-webrtc provider with signaling + password options', async () => {
		const { calls, modules: loaders } = modules();
		const handle = await createCollabProvider(
			'webrtc',
			{
				roomId: 'p2p',
				serverUrl: '',
				userName: 'Ada',
				authToken: 'secret',
				signaling: ['wss://sig.example'],
			},
			doc,
			loaders,
		);
		expect(calls.webrtc).toHaveLength(1);
		// (roomId, doc, { signaling, password })
		expect(calls.webrtc[0]?.[0]).toBe('p2p');
		expect(calls.webrtc[0]?.[2]).toStrictEqual({
			signaling: ['wss://sig.example'],
			password: 'secret',
		});
		// Same-browser tabs meet immediately, so report connected at once.
		expect(handle.connectedNow).toBeTruthy();
	});

	it('omits signaling and password when none is supplied (y-webrtc uses its defaults)', async () => {
		const { calls, modules: loaders } = modules();
		await createCollabProvider(
			'webrtc',
			{ roomId: 'p2p', serverUrl: '', userName: 'Ada' },
			doc,
			loaders,
		);
		expect(calls.webrtc[0]?.[2]).toStrictEqual({});
	});
});
