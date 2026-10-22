import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	COLLABORATION_OFF_STATE,
	INVALID_ROOM_MESSAGE,
	NO_BROADCAST_MESSAGE,
	collaborationStateOf,
	onCollaborationStateChange,
	openCollaborationSession,
	toExternalSession,
} from './embedding';
import { observeExternalSession } from './external-provider';
import { createMemoryHub } from './memory-transport';
import type { CollabSession } from './session';
import { transportProvider } from './transport-provider';

const open: CollabSession[] = [];
const track = <T extends CollabSession>(session: T): T => {
	open.push(session);
	return session;
};
afterEach(() => {
	for (const session of open.splice(0)) session.destroy();
	vi.unstubAllGlobals();
});

function hubJoin(hub: ReturnType<typeof createMemoryHub>, name: string) {
	return openCollaborationSession(
		{
			roomId: 'room-1',
			provider: transportProvider({ transport: hub.createTransport('room-1') }),
			user: { name, color: '#123456', role: 'collaborator' },
		},
		{ defaultName: 'Author', broadcastPrefix: 'test:' },
	);
}

describe('openCollaborationSession', () => {
	it('returns a host session as not owned', () => {
		const host = track(hubJoin(createMemoryHub(), 'Host').session);
		const opened = openCollaborationSession(
			{ session: host },
			{ defaultName: 'x', broadcastPrefix: 'test:' },
		);
		expect(opened).toEqual({ session: host, owned: false });
	});

	it('opens an owned session through the provider with the given identity', () => {
		const { session, owned } = hubJoin(createMemoryHub(), 'Ada');
		track(session);
		expect(owned).toBe(true);
		expect(session.roomId).toBe('room-1');
		expect(session.identity).toMatchObject({ userName: 'Ada', userColor: '#123456' });
	});

	it('falls back to the default name', () => {
		const hub = createMemoryHub();
		const { session } = openCollaborationSession(
			{ roomId: 'r', provider: transportProvider({ transport: hub.createTransport('r') }) },
			{ defaultName: 'Author', broadcastPrefix: 'test:' },
		);
		track(session);
		expect(session.identity.userName).toBe('Author');
	});

	it('rejects invalid rooms and a missing BroadcastChannel', () => {
		const defaults = { defaultName: 'x', broadcastPrefix: 'test:' };
		expect(() => openCollaborationSession({ roomId: 'not a room!' }, defaults)).toThrow(
			INVALID_ROOM_MESSAGE,
		);
		vi.stubGlobal('BroadcastChannel', undefined);
		expect(() => openCollaborationSession({ roomId: 'room' }, defaults)).toThrow(
			NO_BROADCAST_MESSAGE,
		);
	});
});

describe('transports', () => {
	it('sends the auth token to the server', () => {
		const urls: string[] = [];
		class FakeSocket {
			binaryType = '';
			readyState = 0;
			onopen = null;
			onclose = null;
			onerror = null;
			onmessage = null;
			constructor(url: string) {
				urls.push(url);
			}
			send() {}
			close() {}
		}
		vi.stubGlobal('WebSocket', FakeSocket);
		const { session } = openCollaborationSession(
			{ roomId: 'room-1', serverUrl: 'wss://host/collab/', authToken: 'a b' },
			{ defaultName: 'x', broadcastPrefix: 'test:' },
		);
		track(session);
		expect(urls).toEqual(['wss://host/collab/room-1?token=a%20b']);
	});

	it('joins over the host WebrtcProvider with the token as password', () => {
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
		const { session } = openCollaborationSession(
			{
				roomId: 'room-1',
				serverUrl: 'wss://ignored',
				authToken: 'secret',
				webrtc: { WebrtcProvider: FakeWebrtc, signaling: ['wss://signal'] },
			},
			{ defaultName: 'x', broadcastPrefix: 'test:' },
		);
		track(session);
		expect(made).toHaveLength(1);
		expect(made[0]!.roomId).toBe('room-1');
		expect(made[0]!.options).toMatchObject({ signaling: ['wss://signal'], password: 'secret' });
		expect(made[0]!.options.awareness).toBe(session.awareness);
	});
});

describe('collaboration state', () => {
	it('starts off and lists this window first, then peers', () => {
		expect(COLLABORATION_OFF_STATE).toMatchObject({ active: false, status: 'off', people: [] });
		const hub = createMemoryHub();
		const ada = track(hubJoin(hub, 'Ada').session);
		const ben = track(hubJoin(hub, 'Ben').session);
		const changes = vi.fn();
		const stop = onCollaborationStateChange(ada, changes);
		const state = collaborationStateOf(ada);
		expect(state.active).toBe(true);
		expect(state.roomId).toBe('room-1');
		expect(state.people[0]).toMatchObject({ name: 'Ada', self: true });
		expect(state.people.map((p) => p.name)).toContain('Ben');
		ben.disconnect();
		ada.disconnect();
		expect(changes).toHaveBeenCalled();
		stop();
		changes.mockClear();
		ada.connect();
		expect(changes).not.toHaveBeenCalled();
	});
});

describe('toExternalSession', () => {
	it('exposes the doc and awareness and reports the write gate as synced', () => {
		const hub = createMemoryHub();
		const ada = track(hubJoin(hub, 'Ada').session);
		const external = toExternalSession(ada);
		expect(external.doc).toBe(ada.doc);
		expect(external.awareness).toBe(ada.awareness);
		const first = external.getSnapshot();
		expect(external.getSnapshot()).toBe(first);
		const seen: boolean[] = [];
		const stop = observeExternalSession(external, (snapshot) => seen.push(snapshot.synced));
		expect(seen.at(-1)).toBe(ada.synced || ada.gate.isOpen());
		stop();
	});

	it('never destroys the session', () => {
		const ada = track(hubJoin(createMemoryHub(), 'Ada').session);
		const destroy = vi.spyOn(ada, 'destroy');
		const stop = toExternalSession(ada).subscribe(() => {});
		stop();
		expect(destroy).not.toHaveBeenCalled();
	});
});
