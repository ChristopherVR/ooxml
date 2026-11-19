import { describe, expect, it } from 'vitest';

import type { CollabTransportModules } from './collaboration-provider';
import { createCollabProvider } from './collaboration-provider';

interface Calls {
	ws: unknown[][];
	webrtc: unknown[][];
	instances: { messageHandlers: ((...args: unknown[]) => void)[] }[];
	applied: number;
	disconnects: number;
}

function modules(): { calls: Calls; modules: CollabTransportModules } {
	const calls: Calls = { ws: [], webrtc: [], instances: [], applied: 0, disconnects: 0 };
	return {
		calls,
		modules: {
			loadWebsocket: async () => ({
				WebsocketProvider: class {
					awareness = { clientID: 1 };
					wsconnected = true;
					synced = false;
					/** y-websocket's per-instance handler table; entry 0 applies sync messages. */
					messageHandlers: ((...args: unknown[]) => void)[] = [() => calls.applied++];
					constructor(...args: unknown[]) {
						calls.ws.push(args);
						calls.instances.push(this);
					}
					on() {}
					disconnect() {
						calls.disconnects++;
					}
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

	it('refuses an update over budget before y-websocket applies it, and stops syncing', async () => {
		const { calls, modules: loaders } = modules();
		const handle = await createCollabProvider(
			'websocket',
			{ roomId: 'room', serverUrl: 'wss://x', userName: 'Ada' },
			doc,
			loaders,
		);
		const refused: string[] = [];
		handle.onRefused?.((error) => refused.push(error.code));
		// A sync update (type 2) of 100,000 two-byte structs: twice what its size may hold.
		const varUint = (value: number): number[] => {
			const out: number[] = [];
			for (; value > 127; value = Math.floor(value / 128)) out.push((value & 127) | 128);
			return [...out, value];
		};
		const structs = Array.from({ length: 100_000 }, () => [0, 1]).flat();
		const update = [...varUint(1), ...varUint(100_000), 7, 0, ...structs, 0];
		const bytes = new Uint8Array([2, ...varUint(update.length), ...update]);
		// What the guard reads from: lib0's decoder shape, a buffer and a position.
		calls.instances[0]!.messageHandlers[0]!({}, { arr: bytes, pos: 0 });
		expect(calls.applied).toBe(0);
		expect(refused).toEqual(['too-large']);
		expect(calls.disconnects).toBe(1);
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
