import * as Y from 'yjs';
import { describe, expect, it, vi } from 'vitest';
import { createAssetSync } from './assets.js';
import {
	createPresencePublisher,
	borrowAwareness,
	removeAwarenessStatesLocally,
} from './awareness.js';
import {
	applyUpdateSafe,
	encodeSnapshot,
	fromBase64,
	mergeUpdates,
	restoreSnapshot,
	toBase64,
	validateUpdate,
} from './codec.js';
import {
	assignUserColor,
	buildRoster,
	createIdGenerator,
	formatCursorLabel,
	getUserInitials,
} from './identity.js';
import {
	createDepartureChannel,
	registerTeardown,
	type DepartureChannelLike,
	type TeardownListener,
} from './lifecycle.js';
import {
	BoundedMap,
	classifyVersion,
	freezeDeep,
	IdempotencyCache,
	SequenceTracker,
} from './ordering.js';
import { createSyncGate, shouldAutoFollowBroadcaster, shouldRoomReplaceLoad } from './policy.js';
import { deriveCanvasPresence, presenceToCursors } from './presence.js';
import {
	isMixedContentBlocked,
	isValidRoomId,
	sanitizeAvatarUrl,
	sanitizeColor,
	sanitizeUserName,
	stripHtmlTags,
	validateDisplayName,
} from './validation.js';
import {
	adaptYjsProvider,
	observeExternalSession,
	type ExternalSession,
} from './external-provider.js';
import { createWebSocketTransport, roomUrl, type WebSocketLike } from './websocket-transport.js';

describe('validation', () => {
	it('validates room ids, colours, avatars and names', () => {
		expect(isValidRoomId('a_b-1')).toBe(true);
		for (const bad of ['', 'a b', 'x'.repeat(129), '../x', 5])
			expect(isValidRoomId(bad)).toBe(false);
		expect(sanitizeColor('#12ab34')).toBe('#12ab34');
		expect(sanitizeColor('red', '#000000')).toBe('#000000');
		expect(sanitizeAvatarUrl('javascript:alert(1)')).toBeUndefined();
		expect(sanitizeAvatarUrl('https://x/y.png')).toBe('https://x/y.png');
		expect(stripHtmlTags('<scr<script>ipt>alert(1)')).toBe('iptalert(1)'); // no live tag survives
		expect(sanitizeUserName('  <b>Ada</b>  ')).toBe('Ada');
		expect(sanitizeUserName('<>')).toBe('Anonymous');
		expect(() => validateDisplayName('bad\u0007name')).toThrow();
		expect(validateDisplayName(' Ada ')).toBe('Ada');
	});
	it('detects mixed content except on loopback', () => {
		expect(isMixedContentBlocked('ws://example.com', 'https:')).toBe(true);
		expect(isMixedContentBlocked('ws://localhost:1234', 'https:')).toBe(false);
		expect(isMixedContentBlocked('wss://example.com', 'https:')).toBe(false);
		expect(isMixedContentBlocked('ws://example.com', 'http:')).toBe(false);
	});
});

describe('identity', () => {
	it('assigns stable colours, labels and initials', () => {
		expect(assignUserColor('ada')).toBe(assignUserColor('ada'));
		expect(formatCursorLabel('x'.repeat(30))).toHaveLength(20);
		expect(getUserInitials('Ada Lovelace King')).toBe('AK');
		expect(getUserInitials('ada')).toBe('AD');
		const roster = buildRoster({
			localUserName: 'Me',
			remoteUsers: [{ clientId: 7, userName: 'Bob Ray', userColor: '#000000' }],
		});
		expect(roster.map((u) => u.id)).toEqual(['local', '7']);
	});
	it('generates unique ids per client', () => {
		const a = createIdGenerator('client-a', 'dve');
		const b = createIdGenerator('client-b', 'dve');
		expect(a('paragraph')).not.toBe(b('paragraph'));
		expect(a('paragraph')).not.toBe(a('paragraph'));
		expect(() => createIdGenerator('')).toThrow();
	});
});

describe('presence derivation', () => {
	it('drops local, malformed and stale entries and clamps coordinates', () => {
		const now = Date.parse('2026-01-01T00:00:00Z');
		const fresh = new Date(now - 1000).toISOString();
		const states = new Map<number, Record<string, unknown>>([
			[1, { presence: { userName: 'me', lastUpdated: fresh } }],
			[
				2,
				{
					presence: {
						userName: '<i>Bob</i>',
						userColor: 'nope',
						cursorX: 99999,
						cursorY: -5,
						activeSlideIndex: 2.7,
						lastUpdated: fresh,
					},
				},
			],
			[3, { presence: { userName: 'old', lastUpdated: new Date(now - 60_000).toISOString() } }],
			[4, { presence: 'junk' }],
			[5, {}],
		]);
		const list = deriveCanvasPresence(states, 1, 960, 540, { now });
		expect(list).toHaveLength(1);
		expect(list[0]).toMatchObject({
			clientId: 2,
			userName: 'Bob',
			cursorX: 980,
			cursorY: -5,
			activeSlideIndex: 2,
		});
		expect(presenceToCursors(list, 0)).toEqual([]);
		expect(presenceToCursors(list, 2)).toHaveLength(1);
	});
});

describe('presence publisher', () => {
	function fakeAwareness() {
		const writes: Record<string, unknown>[] = [];
		return {
			writes,
			awareness: {
				setLocalStateField: (_f: string, v: unknown) => writes.push(v as Record<string, unknown>),
				getStates: () => new Map(),
				on: () => {},
			},
		};
	}
	it('announces, throttles, drops no-op patches and always flushes', () => {
		vi.useFakeTimers();
		const { writes, awareness } = fakeAwareness();
		const publisher = createPresencePublisher(
			awareness,
			{ userName: 'A', userColor: '#000000' },
			{ x: 0 },
		);
		expect(writes).toHaveLength(1);
		publisher.update({ x: 0 });
		expect(writes).toHaveLength(1);
		publisher.update({ x: 1 });
		expect(writes).toHaveLength(2);
		publisher.update({ x: 2 });
		publisher.update({ x: 3 });
		expect(writes).toHaveLength(2);
		vi.advanceTimersByTime(60);
		expect(writes).toHaveLength(3);
		expect(writes[2]).toMatchObject({ x: 3, userName: 'A' });
		publisher.flush();
		expect(writes).toHaveLength(4);
		publisher.dispose();
		vi.useRealTimers();
	});
	it('restores only its own presence on a borrowed awareness', () => {
		let state: Record<string, unknown> | null = { host: 1, presence: 'host-presence' };
		const host = {
			clientID: 1,
			getLocalState: () => state,
			setLocalState: (s: Record<string, unknown> | null) => (state = s),
			setLocalStateField: (f: string, v: unknown) => (state = { ...state, [f]: v }),
			getStates: () => new Map(),
			on: () => {},
			off: () => {},
		};
		const lease = borrowAwareness(host);
		lease.awareness.setLocalStateField('presence', 'viewer');
		expect(state).toMatchObject({ presence: 'viewer' });
		lease.dispose();
		expect(state).toEqual({ host: 1, presence: 'host-presence' });
	});
	it('removes departed peers and notifies', () => {
		const emit = vi.fn();
		const awareness = {
			states: new Map<number, Record<string, unknown>>([
				[2, {}],
				[3, {}],
			]),
			emit,
		};
		expect(removeAwarenessStatesLocally(awareness, [2, 9])).toEqual([2]);
		expect(emit).toHaveBeenCalledTimes(2);
	});
});

describe('ordering', () => {
	it('classifies versions', () => {
		expect(classifyVersion(3, 3)).toBe('apply');
		expect(classifyVersion(2, 3)).toBe('stale');
		expect(classifyVersion(4, 3)).toBe('out-of-order');
		expect(classifyVersion(-1, 3)).toBe('invalid');
		expect(classifyVersion(1.5, 3)).toBe('invalid');
	});
	it('detects duplicates and id reuse', () => {
		const cache = new IdempotencyCache<string>();
		expect(cache.check('c', '1', 'f').status).toBe('fresh');
		cache.remember('c', '1', 'f', 'result');
		expect(cache.check('c', '1', 'f')).toEqual({ status: 'duplicate', value: 'result' });
		expect(cache.check('c', '1', 'other').status).toBe('reused');
	});
	it('tracks per-peer sequences', () => {
		const tracker = new SequenceTracker();
		expect(tracker.classify('p', 1, 'a')).toBe('apply');
		tracker.commit('p', 1, 'a');
		expect(tracker.classify('p', 1, 'a')).toBe('duplicate');
		expect(tracker.classify('p', 1, 'b')).toBe('invalid');
		expect(tracker.classify('p', 2, 'c')).toBe('apply');
		tracker.commit('p', 2, 'c');
		expect(tracker.classify('p', 1, 'a')).toBe('stale');
		expect(tracker.classify('p', 0, 'a')).toBe('invalid');
	});
	it('bounds memory and freezes envelopes', () => {
		const map = new BoundedMap<number, number>(2);
		map.set(1, 1);
		map.set(2, 2);
		map.set(3, 3);
		expect(map.has(1)).toBe(false);
		expect(map.size).toBe(2);
		const frozen = freezeDeep({ a: { b: [1] } });
		expect(Object.isFrozen(frozen.a.b)).toBe(true);
	});
});

describe('policy', () => {
	it('opens the gate once, from sync or grace', () => {
		vi.useFakeTimers();
		const onOpen = vi.fn();
		const gate = createSyncGate(onOpen, 100);
		gate.arm();
		vi.advanceTimersByTime(100);
		gate.open();
		expect(onOpen).toHaveBeenCalledTimes(1);
		gate.reset();
		expect(gate.isOpen()).toBe(false);
		gate.arm();
		gate.open();
		expect(onOpen).toHaveBeenCalledTimes(2);
		vi.useRealTimers();
	});
	it('applies the load and follow rules', () => {
		expect(shouldRoomReplaceLoad('bootstrap', true)).toBe(true);
		expect(shouldRoomReplaceLoad('bootstrap', false)).toBe(false);
		expect(shouldRoomReplaceLoad('user', true)).toBe(false);
		expect(shouldRoomReplaceLoad(undefined, true)).toBe(false);
		expect(shouldAutoFollowBroadcaster({ localRole: 'viewer', broadcasterRole: 'owner' })).toBe(
			true,
		);
		expect(
			shouldAutoFollowBroadcaster({ localRole: 'collaborator', broadcasterRole: 'owner' }),
		).toBe(false);
	});
});

describe('lifecycle', () => {
	it('leaves once on pagehide, honours bfcache and rejoin', () => {
		const listeners = new Map<string, TeardownListener>();
		const target = {
			addEventListener: (t: string, l: TeardownListener) => listeners.set(t, l),
			removeEventListener: (t: string) => listeners.delete(t),
		};
		const leave = vi.fn();
		const rejoin = vi.fn();
		const dispose = registerTeardown({ leave, rejoin, target });
		listeners.get('pagehide')!({ persisted: true });
		listeners.get('beforeunload')!({});
		expect(leave).toHaveBeenCalledTimes(1);
		listeners.get('pageshow')!({ persisted: true });
		expect(rejoin).toHaveBeenCalledTimes(1);
		listeners.get('message')!({ data: { type: 'ooxml-core:collab-leave' } });
		expect(leave).toHaveBeenCalledTimes(2);
		dispose();
		expect(listeners.size).toBe(0);
	});
	it('ignores persisted pagehide without rejoin', () => {
		const listeners = new Map<string, TeardownListener>();
		const target = {
			addEventListener: (t: string, l: TeardownListener) => listeners.set(t, l),
			removeEventListener: () => {},
		};
		const leave = vi.fn();
		registerTeardown({ leave, target });
		listeners.get('pagehide')!({ persisted: true });
		expect(leave).not.toHaveBeenCalled();
	});
	it('drops a departed peer announced on the channel, not itself or other rooms', () => {
		const channel: DepartureChannelLike = { postMessage: vi.fn(), close: vi.fn(), onmessage: null };
		const awareness = {
			clientID: 1,
			states: new Map<number, Record<string, unknown>>([
				[2, {}],
				[3, {}],
			]),
		};
		const dep = createDepartureChannel('room', awareness, () => channel, 'chan');
		channel.onmessage!({ data: { channel: 'chan', roomId: 'other', clientId: 2 } });
		channel.onmessage!({ data: { channel: 'chan', roomId: 'room', clientId: 1 } });
		channel.onmessage!({ data: 'junk' });
		expect(awareness.states.size).toBe(2);
		channel.onmessage!({ data: { channel: 'chan', roomId: 'room', clientId: 2 } });
		expect([...awareness.states.keys()]).toEqual([3]);
		dep.announce();
		expect(channel.postMessage).toHaveBeenCalledWith({
			channel: 'chan',
			roomId: 'room',
			clientId: 1,
		});
		dep.dispose();
		dep.announce();
		expect(channel.postMessage).toHaveBeenCalledTimes(1);
	});
});

describe('assets', () => {
	const sync = createAssetSync({ mapName: 'assets', fields: { mediaData: '_mdRef' } });
	it('routes payloads through the asset map and bumps the version on a swap', () => {
		const doc = new Y.Doc();
		const assets = sync.getMap(doc);
		const owner = new Map<string, unknown>();
		const ownerLike = {
			get: (k: string) => owner.get(k),
			set: (k: string, v: unknown) => void owner.set(k, v),
			delete: (k: string) => void owner.delete(k),
			forEach: () => {},
		};
		sync.write('e1', { mediaData: 'AAA' }, ownerLike, assets);
		expect(owner.get('_mdRef')).toBe('e1:mediaData');
		sync.reconcile('e1', { mediaData: 'BBB' }, ownerLike, assets);
		expect(owner.get('mediaData__v')).toBe(1);
		expect(sync.isVersionKey('mediaData__v')).toBe(true);
		expect(sync.isVersionKey('other__v')).toBe(false);
		const target: Record<string, unknown> = {};
		sync.read(ownerLike, assets, target);
		expect(target.mediaData).toBe('BBB');
		sync.reconcile('e1', {}, ownerLike, assets);
		expect(owner.has('_mdRef')).toBe(false);
		expect(assets.get('e1:mediaData')).toBeUndefined();
	});
});

describe('codec', () => {
	it('round-trips snapshots, merges updates and base64', () => {
		const doc = new Y.Doc();
		doc.getText('t').insert(0, 'abc');
		const restored = restoreSnapshot(encodeSnapshot(doc));
		expect(restored.getText('t').toString()).toBe('abc');
		expect(() => restoreSnapshot(new Uint8Array([9, 9, 9, 9]))).toThrow();
		const bytes = new Uint8Array(70_000).map((_, i) => i % 251);
		expect(fromBase64(toBase64(bytes))).toEqual(bytes);
		expect(fromBase64('***')).toBeNull();
		const u1 = encodeSnapshot(doc);
		doc.getText('t').insert(3, 'd');
		const merged = mergeUpdates([u1, encodeSnapshot(doc)]);
		expect(restoreSnapshot(merged).getText('t').toString()).toBe('abcd');
		expect(validateUpdate(u1).ok).toBe(true);
		expect(applyUpdateSafe(new Y.Doc(), u1.slice(0, 3)).ok).toBe(false);
	});
});

describe('external providers', () => {
	it('adapts a y-websocket style provider', () => {
		const handlers = new Map<string, (p: never) => void>();
		const provider = {
			on: (e: string, h: (p: never) => void) => handlers.set(e, h),
			off: vi.fn(),
			connect: vi.fn(),
			destroy: vi.fn(),
		};
		const adapted = adaptYjsProvider(provider);
		const statuses: string[] = [];
		const synced: boolean[] = [];
		adapted.on('status', (s) => statuses.push(s));
		adapted.on('synced', (s) => synced.push(s));
		(handlers.get('status') as (p: unknown) => void)({ status: 'connected' });
		(handlers.get('sync') as (p: unknown) => void)(true);
		(handlers.get('status') as (p: unknown) => void)({ status: 'disconnected' });
		expect(statuses).toEqual(['connected', 'disconnected']);
		expect(synced).toEqual([true, false]);
		adapted.destroy();
		expect(provider.destroy).toHaveBeenCalled();
	});
	it('observes a host session, delivering the first snapshot and only changes', () => {
		let snapshot = { status: 'connecting' as const, synced: false } as {
			status: 'connecting' | 'connected';
			synced: boolean;
		};
		let notify = () => {};
		const session = {
			getSnapshot: () => snapshot,
			subscribe: (l: () => void) => ((notify = l), () => {}),
		} as unknown as ExternalSession;
		const seen: unknown[] = [];
		observeExternalSession(session, (s) => seen.push(s));
		notify();
		snapshot = { status: 'connected', synced: true };
		notify();
		expect(seen).toEqual([
			{ status: 'connecting', synced: false },
			{ status: 'connected', synced: true },
		]);
	});
});

describe('websocket transport', () => {
	class FakeSocket implements WebSocketLike {
		static instances: FakeSocket[] = [];
		binaryType = '';
		readyState = 0;
		sent: Uint8Array[] = [];
		onopen: WebSocketLike['onopen'] = null;
		onclose: WebSocketLike['onclose'] = null;
		onerror: WebSocketLike['onerror'] = null;
		onmessage: WebSocketLike['onmessage'] = null;
		constructor(public url: string) {
			FakeSocket.instances.push(this);
		}
		send(data: Uint8Array) {
			this.sent.push(data);
		}
		close() {}
	}
	it('reconnects with backoff, delivers binary and stops on disconnect', () => {
		vi.useFakeTimers();
		FakeSocket.instances = [];
		const transport = createWebSocketTransport({
			url: roomUrl('wss://h/c/', 'r 1'),
			WebSocket: FakeSocket,
			baseDelayMs: 100,
		});
		const handlers = { open: vi.fn(), close: vi.fn(), message: vi.fn(), error: vi.fn() };
		transport.connect(handlers);
		const first = FakeSocket.instances[0]!;
		expect(first.url).toBe('wss://h/c/r%201');
		first.readyState = 1;
		first.onopen!({});
		first.onmessage!({ data: new Uint8Array([1]).buffer });
		first.onmessage!({ data: 'text' });
		expect(handlers.message).toHaveBeenCalledWith(new Uint8Array([1]));
		expect(handlers.error).toHaveBeenCalledTimes(1);
		transport.send(new Uint8Array([2]));
		expect(first.sent).toHaveLength(1);
		first.onclose!({});
		vi.advanceTimersByTime(100);
		expect(FakeSocket.instances).toHaveLength(2);
		transport.disconnect();
		vi.advanceTimersByTime(10_000);
		expect(FakeSocket.instances).toHaveLength(2);
		vi.useRealTimers();
	});
	it('fails fast on mixed content', () => {
		const handlers = { open: vi.fn(), close: vi.fn(), message: vi.fn(), error: vi.fn() };
		FakeSocket.instances = [];
		createWebSocketTransport({
			url: 'ws://example.com/r',
			WebSocket: FakeSocket,
			pageProtocol: 'https:',
		}).connect(handlers);
		expect(handlers.error).toHaveBeenCalled();
		expect(FakeSocket.instances).toHaveLength(0);
	});
});
