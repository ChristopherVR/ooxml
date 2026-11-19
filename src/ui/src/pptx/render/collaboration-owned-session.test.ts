import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { CollaborationConfig } from '../types';
import { collaborationFailureOf } from './collaboration-failure';
import { resolveCollaborationConfig } from './collaboration-options';
import { needsOwnedSession, openOwnedCollaboration } from './collaboration-owned-session';

const opened: { destroy: () => void }[] = [];
afterEach(() => {
	for (const item of opened.splice(0)) item.destroy();
});

function hostSession(hub: ReturnType<typeof createMemoryHub>, roomId = 'room-1') {
	const session = createCollabSession({
		roomId,
		provider: transportProvider({ transport: hub.createTransport(roomId) }),
		user: { name: 'Host' },
		heartbeatMs: 0,
	});
	opened.push(session);
	return session;
}

describe('needsOwnedSession', () => {
	const base: CollaborationConfig = { roomId: 'r', serverUrl: '', userName: 'A' };
	it('keeps websocket and webrtc configs on the built-in provider path', () => {
		expect(needsOwnedSession({ ...base, serverUrl: 'wss://x' })).toBe(false);
		expect(needsOwnedSession({ ...base, transport: 'webrtc' })).toBe(false);
	});
	it('opens a session for a provider (over serverUrl) or for no server at all', () => {
		const provider = (() => undefined) as never;
		expect(needsOwnedSession({ ...base, provider })).toBe(true);
		expect(needsOwnedSession({ ...base, serverUrl: 'wss://x', provider })).toBe(true);
		const shared = resolveCollaborationConfig({ roomId: 'r', user: { name: 'A' } });
		expect(needsOwnedSession(shared!)).toBe(true);
	});
	it('keeps a legacy blank serverUrl on the serverless WebRTC transport', () => {
		expect(needsOwnedSession(base)).toBe(false);
		expect(needsOwnedSession(resolveCollaborationConfig({ ...base, intent: 'join' })!)).toBe(false);
		expect(
			needsOwnedSession(
				resolveCollaborationConfig({ roomId: 'r', transport: 'webrtc', user: { name: 'A' } })!,
			),
		).toBe(false);
	});
	it('always borrows a host session, even with a provider', () => {
		const session = hostSession(createMemoryHub());
		const config = resolveCollaborationConfig({ session, provider: (() => undefined) as never });
		expect(needsOwnedSession(config!)).toBe(false);
	});
});

describe('openOwnedCollaboration', () => {
	it('connects through a provider and shares the document with a host session', async () => {
		const hub = createMemoryHub();
		const host = hostSession(hub);
		const config = resolveCollaborationConfig({
			roomId: 'room-1',
			provider: transportProvider({ transport: hub.createTransport('room-1') }),
			user: { name: 'Ada', color: '#112233' },
			intent: 'join',
		})!;
		const owned = openOwnedCollaboration(config);
		const prepared = (await owned.ready)!;
		expect(prepared.externalSession).toBeDefined();
		expect(prepared).toMatchObject({
			roomId: 'room-1',
			userName: 'Ada',
			userColor: '#112233',
			sessionIntent: 'join',
		});
		host.doc.getMap('shared').set('k', 'v');
		expect(prepared.externalSession!.doc.getMap('shared').get('k')).toBe('v');
		owned.release();
		expect(prepared.externalSession!.awareness.getLocalState()).toBeNull();
	});

	it('reports an update the owned session refused', async () => {
		const hub = createMemoryHub();
		const owned = openOwnedCollaboration(
			resolveCollaborationConfig({
				roomId: 'room-1',
				provider: transportProvider({ transport: hub.createTransport('room-1') }),
				user: { name: 'Ada' },
			})!,
		);
		const refused: string[] = [];
		owned.onRefused((code) => refused.push(code));
		const prepared = (await owned.ready)!;
		const evil = hub.createTransport('room-1');
		evil.connect({ open: () => {}, close: () => {}, message: () => {}, error: () => {} });
		// A whole sync update whose struct has an unknown content type.
		evil.send(new Uint8Array([0, 2, 5, 1, 1, 1, 0, 0x1f]));
		expect(refused).toEqual(['malformed']);
		expect(prepared.externalSession!.getSnapshot().status).toBe('error');
		owned.release();
	});

	it('release destroys the owned session only, never a borrowed host session', async () => {
		const hub = createMemoryHub();
		const host = hostSession(hub);
		const owned = openOwnedCollaboration({
			roomId: 'room-1',
			serverUrl: '',
			userName: '',
			provider: transportProvider({ transport: hub.createTransport('room-1') }),
		});
		const prepared = (await owned.ready)!;
		expect(prepared.userName).toBe('PowerPoint user');
		const awareness = prepared.externalSession!.awareness;
		expect(awareness.getLocalState()).not.toBeNull();
		owned.release();
		expect(awareness.getLocalState()).toBeNull();
		expect(host.awareness.getLocalState()).not.toBeNull();
	});

	it('releasing while the import is pending opens nothing', async () => {
		const hub = createMemoryHub();
		const owned = openOwnedCollaboration({
			roomId: 'room-1',
			serverUrl: '',
			userName: 'A',
			provider: transportProvider({ transport: hub.createTransport('room-1') }),
		});
		owned.release();
		expect(await owned.ready).toBeNull();
	});

	it('opens the host y-webrtc class from the shared webrtc option with the token as password', async () => {
		const made: { roomId: string; options: Record<string, unknown> }[] = [];
		class FakeWebrtc {
			connected = false;
			constructor(roomId: string, _doc: unknown, options: Record<string, unknown>) {
				made.push({ roomId, options });
			}
			on() {}
			off() {}
			destroy() {}
		}
		const config = resolveCollaborationConfig({
			roomId: 'room-1',
			authToken: 'secret',
			webrtc: { WebrtcProvider: FakeWebrtc, signaling: ['wss://signal'] },
			user: { name: 'Ada' },
		});
		expect(needsOwnedSession(config!)).toBe(true);
		const owned = openOwnedCollaboration(config!);
		const ready = await owned.ready;
		expect(ready?.externalSession).toBeDefined();
		expect(made).toHaveLength(1);
		expect(made[0]!.roomId).toBe('room-1');
		expect(made[0]!.options).toMatchObject({ signaling: ['wss://signal'], password: 'secret' });
		owned.release();
	});

	it('rejects an invalid room', async () => {
		const owned = openOwnedCollaboration({ roomId: '', serverUrl: '', userName: 'A' });
		await expect(owned.ready).rejects.toThrow(/room/);
		expect(collaborationFailureOf(await owned.ready.catch((error: unknown) => error))).toBe(
			'invalid-room',
		);
	});

	it('names why an owned room cannot open: no BroadcastChannel, no Web Crypto, a throwing provider', async () => {
		const reason = (config: CollaborationConfig) =>
			openOwnedCollaboration(config).ready.then(
				() => 'opened',
				(error: unknown) => collaborationFailureOf(error),
			);
		vi.stubGlobal('BroadcastChannel', undefined);
		try {
			expect(await reason({ roomId: 'room-1', serverUrl: '', userName: 'A' })).toBe('no-broadcast');
		} finally {
			vi.unstubAllGlobals();
		}
		class ThrowingWebrtc {
			constructor() {
				throw new Error('boom');
			}
		}
		const webrtc = resolveCollaborationConfig({
			roomId: 'room-1',
			authToken: 'secret',
			webrtc: { WebrtcProvider: ThrowingWebrtc as never },
		})!;
		expect(await reason(webrtc)).toBe('unknown');
		vi.stubGlobal('crypto', {});
		try {
			expect(await reason(webrtc)).toBe('insecure-page');
		} finally {
			vi.unstubAllGlobals();
		}
	});

	it('keeps the PowerPoint presence payload when the session heartbeats identity', async () => {
		const hub = createMemoryHub();
		const owned = openOwnedCollaboration({
			roomId: 'room-1',
			serverUrl: '',
			userName: 'A',
			provider: transportProvider({ transport: hub.createTransport('room-1') }),
		});
		const { awareness } = (await owned.ready)!.externalSession!;
		awareness.setLocalStateField('presence', { userName: 'A', activeSlideIndex: 2, cursorX: 5 });
		awareness.setLocalStateField('presence', { userName: 'A', lastUpdated: 'now' });
		expect(awareness.getLocalState()?.presence).toMatchObject({
			activeSlideIndex: 2,
			cursorX: 5,
			lastUpdated: 'now',
		});
		owned.release();
	});
});
