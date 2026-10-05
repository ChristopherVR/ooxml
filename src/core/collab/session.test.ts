import * as Y from 'yjs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { applyUpdateSafe } from './codec.js';
import { createMemoryHub } from './memory-transport.js';
import { createCollabSession, type CollabSession } from './session.js';
import { transportProvider } from './transport-provider.js';

interface Cursor {
	page: number;
}
const sanitize = (raw: Record<string, unknown>): Cursor | null =>
	typeof raw.page === 'number' ? { page: raw.page } : null;

const open: CollabSession<Cursor>[] = [];
function join(
	hub: ReturnType<typeof createMemoryHub>,
	name: string,
	extra: Partial<Parameters<typeof createCollabSession<Cursor>>[0]> = {},
): CollabSession<Cursor> {
	const session = createCollabSession<Cursor>({
		roomId: 'room-1',
		provider: transportProvider({ transport: hub.createTransport('room-1') }),
		user: { name },
		initialPresence: { page: 0 },
		sanitizePayload: sanitize,
		heartbeatMs: 0,
		...extra,
	});
	open.push(session);
	return session;
}
afterEach(() => {
	for (const session of open.splice(0)) session.destroy();
	vi.useRealTimers();
});

describe('two in-memory clients', () => {
	it('converge on concurrent edits to the same text and map', () => {
		const hub = createMemoryHub();
		const a = join(hub, 'Ada');
		const b = join(hub, 'Bob');
		a.doc.getText('t').insert(0, 'hello');
		b.doc.getMap('m').set('k', 1);
		expect(b.doc.getText('t').toString()).toBe('hello');
		expect(a.doc.getMap('m').get('k')).toBe(1);
		a.doc.getText('t').insert(5, ' A');
		b.doc.getText('t').insert(5, ' B');
		expect(a.doc.getText('t').toString()).toBe(b.doc.getText('t').toString());
		expect(a.doc.getText('t').toString()).toHaveLength(9);
	});

	it('also converge with asynchronous delivery', async () => {
		const hub = createMemoryHub({ async: true });
		const a = join(hub, 'Ada');
		const b = join(hub, 'Bob');
		await Promise.resolve();
		await Promise.resolve();
		a.doc.getArray('l').push([1]);
		b.doc.getArray('l').push([2]);
		await new Promise((resolve) => setTimeout(resolve, 5));
		expect(a.doc.getArray('l').toJSON().sort()).toEqual([1, 2]);
		expect(b.doc.getArray('l').toJSON().sort()).toEqual([1, 2]);
	});

	it('gives a late joiner the existing state and syncs', () => {
		const hub = createMemoryHub();
		const a = join(hub, 'Ada');
		a.doc.getText('t').insert(0, 'existing');
		const b = join(hub, 'Bob');
		expect(b.doc.getText('t').toString()).toBe('existing');
		expect(b.synced).toBe(true);
	});

	it('tolerates dropped messages: a later handshake repairs the gap', () => {
		let drop = true;
		const hub = createMemoryHub({ filter: () => !drop });
		const a = join(hub, 'Ada');
		const b = join(hub, 'Bob');
		a.doc.getText('t').insert(0, 'lost');
		expect(b.doc.getText('t').toString()).toBe('');
		drop = false;
		b.disconnect();
		b.connect();
		expect(b.doc.getText('t').toString()).toBe('lost');
	});
});

describe('presence', () => {
	it('shows joins and leaves and sanitises the payload', () => {
		const hub = createMemoryHub();
		const a = join(hub, 'Ada');
		const seen: string[][] = [];
		a.on('peers', (peers) => seen.push(peers.map((p) => p.userName)));
		const b = join(hub, 'Bob');
		expect(a.peers().map((p) => p.userName)).toEqual(['Bob']);
		expect(b.peers().map((p) => p.userName)).toEqual(['Ada']);
		b.updatePresence({ page: 3 });
		expect(a.peers()[0]?.page).toBe(3);
		b.destroy();
		expect(a.peers()).toEqual([]);
		expect(seen.some((names) => names.includes('Bob'))).toBe(true);
	});

	it('hides peers whose payload fails validation', () => {
		const hub = createMemoryHub();
		const a = join(hub, 'Ada');
		const b = join(hub, 'Bob');
		b.awareness.setLocalStateField('presence', {
			userName: 'Bob',
			page: 'x',
			lastUpdated: new Date().toISOString(),
		});
		expect(a.peers()).toEqual([]);
	});

	it('survives a reconnect: presence is withdrawn then restored', () => {
		const hub = createMemoryHub();
		const a = join(hub, 'Ada');
		const b = join(hub, 'Bob');
		b.disconnect();
		expect(a.peers()).toEqual([]);
		b.connect();
		expect(a.peers().map((p) => p.userName)).toEqual(['Bob']);
		expect(b.peers().map((p) => p.userName)).toEqual(['Ada']);
	});

	it('ages out stale peers on heartbeat', () => {
		vi.useFakeTimers();
		const hub = createMemoryHub();
		const a = join(hub, 'Ada', { heartbeatMs: 10_000 });
		const b = join(hub, 'Bob', { heartbeatMs: 0 });
		expect(a.peers()).toHaveLength(1);
		vi.advanceTimersByTime(40_000);
		expect(a.peers()).toHaveLength(0);
		void b;
	});
});

describe('reconnect and offline edits', () => {
	it('merges edits made while disconnected after reconnecting', () => {
		const hub = createMemoryHub();
		const a = join(hub, 'Ada');
		const b = join(hub, 'Bob');
		b.disconnect();
		b.doc.getText('t').insert(0, 'offline-b');
		a.doc.getText('t').insert(0, 'live-a');
		expect(a.doc.getText('t').toString()).toBe('live-a');
		b.connect();
		expect(a.doc.getText('t').toString()).toBe(b.doc.getText('t').toString());
		expect(a.doc.getText('t').toString()).toContain('offline-b');
		expect(a.doc.getText('t').toString()).toContain('live-a');
	});

	it('reports status transitions', () => {
		const hub = createMemoryHub();
		const statuses: string[] = [];
		const a = join(hub, 'Ada', { autoConnect: false });
		a.on('status', (s) => statuses.push(s));
		a.connect();
		a.disconnect();
		expect(statuses).toEqual(['connecting', 'connected', 'disconnected']);
	});
});

describe('malformed input', () => {
	it('rejects garbage, truncated and oversized messages without corrupting the document', () => {
		const hub = createMemoryHub();
		const a = join(hub, 'Ada');
		a.doc.getText('t').insert(0, 'safe');
		const errors: Error[] = [];
		a.on('error', (error) => errors.push(error));
		const evil = hub.createTransport('room-1');
		const sender = { send: (bytes: Uint8Array) => evil.send(bytes) };
		evil.connect({ open: () => {}, close: () => {}, message: () => {}, error: () => {} });
		sender.send(new Uint8Array([0, 2, 5, 1, 2])); // sync update with truncated payload
		sender.send(new Uint8Array([0, 9, 0])); // unknown sync sub-type
		sender.send(new Uint8Array([1, 3, 9])); // truncated awareness
		sender.send(new Uint8Array([0])); // type only
		sender.send(new Uint8Array(0)); // empty
		expect(errors.length).toBeGreaterThanOrEqual(4);
		expect(a.doc.getText('t').toString()).toBe('safe');
		a.doc.getText('t').insert(4, '!');
		expect(a.doc.getText('t').toString()).toBe('safe!');
	});

	it('applyUpdateSafe refuses non-updates and leaves the doc untouched', () => {
		const doc = new Y.Doc();
		doc.getText('t').insert(0, 'x');
		for (const bad of [new Uint8Array([255, 255, 255]), new Uint8Array(0), 'text', null]) {
			const result = applyUpdateSafe(doc, bad);
			expect(result.ok).toBe(false);
		}
		expect(applyUpdateSafe(doc, new Uint8Array(10), undefined, { maxBytes: 5 })).toMatchObject({
			ok: false,
		});
		expect(doc.getText('t').toString()).toBe('x');
	});
});

describe('write gating', () => {
	it('opens the gate on sync, blocks viewers, and falls back to the grace period alone', () => {
		vi.useFakeTimers();
		const hub = createMemoryHub();
		const ready = vi.fn();
		const solo = join(hub, 'Solo', { onReady: ready, syncGraceMs: 500 });
		expect(solo.canWrite()).toBe(false);
		vi.advanceTimersByTime(500);
		expect(solo.canWrite()).toBe(true);
		expect(ready).toHaveBeenCalledTimes(1);
		const viewer = join(hub, 'Viewer', { user: { name: 'Viewer', role: 'viewer' } });
		expect(viewer.synced).toBe(true);
		expect(viewer.canWrite()).toBe(false);
	});

	it('rejects an invalid room id up front', () => {
		const hub = createMemoryHub();
		expect(() => join(hub, 'Ada', { roomId: 'bad room!' })).toThrow(/Invalid collaboration room/);
	});
});
