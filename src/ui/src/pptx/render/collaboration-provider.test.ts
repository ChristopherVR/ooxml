import { createShareSecret } from 'ooxml-core/collab';
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

	it('opens a serverless y-webrtc room named by the secret, not the room name', async () => {
		const { calls, modules: loaders } = modules();
		const secret = createShareSecret();
		const handle = await createCollabProvider(
			'webrtc',
			{
				roomId: 'quarterly-board-deck',
				serverUrl: '',
				userName: 'Ada',
				secret,
				signaling: ['wss://sig.example'],
			},
			doc,
			loaders,
		);
		expect(calls.webrtc).toHaveLength(1);
		// (room, doc, { signaling, password })
		const room = calls.webrtc[0]?.[0];
		expect(room).toMatch(/^[0-9a-f]{64}$/u);
		expect(room).not.toContain('quarterly');
		expect(room).not.toContain(secret);
		expect(calls.webrtc[0]?.[2]).toStrictEqual({
			signaling: ['wss://sig.example'],
			password: secret,
		});
		// Same-browser tabs meet immediately, so report connected at once.
		expect(handle.connectedNow).toBeTruthy();
	});

	it('gives two sessions with the same room name but different secrets different rooms', async () => {
		const { calls, modules: loaders } = modules();
		for (const secret of [createShareSecret(), createShareSecret()]) {
			await createCollabProvider(
				'webrtc',
				{ roomId: 'team', serverUrl: '', userName: 'Ada', secret },
				doc,
				loaders,
			);
		}
		expect(calls.webrtc[0]?.[0]).not.toBe(calls.webrtc[1]?.[0]);
		expect(calls.webrtc[0]?.[2]).toStrictEqual({ password: expect.any(String) });
	});

	it('refuses to open a serverless room without a valid secret', async () => {
		const { calls, modules: loaders } = modules();
		for (const secret of [undefined, 'p2p', 'x'.repeat(22)]) {
			await expect(
				createCollabProvider(
					'webrtc',
					{
						roomId: 'p2p',
						serverUrl: '',
						userName: 'Ada',
						authToken: 'token',
						...(secret ? { secret } : {}),
					},
					doc,
					loaders,
				),
			).rejects.toThrow(/secret/u);
		}
		expect(calls.webrtc).toHaveLength(0);
	});
});
