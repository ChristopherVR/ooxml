import * as Y from 'yjs';
import { afterEach, describe, expect, it } from 'vitest';
import { bindDocument, type DocumentAdapter } from './binding.js';
import { createMemoryHub } from './memory-transport.js';
import { createCollabSession, type CollabSession } from './session.js';
import { transportProvider } from './transport-provider.js';

// A toy product mapping: the model is a list of titles in a Y.Array.
const adapter: DocumentAdapter<string[]> = {
	isEmpty: (doc) => doc.getArray('titles').length === 0,
	read: (doc) => doc.getArray<string>('titles').toArray(),
	write: (doc, model, origin) =>
		doc.transact(() => {
			const array = doc.getArray<string>('titles');
			array.delete(0, array.length);
			array.push(model);
		}, origin),
	observe: (doc, onChange) => {
		const array = doc.getArray('titles');
		const handler = (_events: unknown, tx: Y.Transaction) => onChange(tx.origin);
		array.observeDeep(handler);
		return () => array.unobserveDeep(handler);
	},
};

const sessions: CollabSession[] = [];
afterEach(() => sessions.splice(0).forEach((s) => s.destroy()));
function join(hub: ReturnType<typeof createMemoryHub>, role?: 'viewer') {
	const session = createCollabSession({
		roomId: 'r',
		provider: transportProvider({ transport: hub.createTransport('r') }),
		user: { name: 'u', ...(role ? { role } : {}) },
		heartbeatMs: 0,
		syncGraceMs: 0,
	});
	sessions.push(session);
	return session;
}

describe('bindDocument', () => {
	it('seeds an empty room, adopts it on the late joiner, and skips local echoes', async () => {
		const hub = createMemoryHub();
		const a = join(hub);
		const modelA: string[] = ['one'];
		const adoptedA: string[][] = [];
		const bindA = bindDocument(a, adapter, {
			getLocalModel: () => modelA,
			onRemoteModel: (m) => adoptedA.push(m),
		});
		await new Promise((r) => setTimeout(r, 5));
		expect(a.doc.getArray('titles').toJSON()).toEqual(['one']);
		expect(adoptedA).toEqual([]); // its own seed is not an echo

		const b = join(hub);
		const adoptedB: string[][] = [];
		bindDocument(b, adapter, {
			getLocalModel: () => ['bootstrap'],
			onRemoteModel: (m) => adoptedB.push(m),
		});
		await new Promise((r) => setTimeout(r, 5));
		expect(b.doc.getArray('titles').toJSON()).toEqual(['one']);
		expect(adoptedB.at(-1)).toEqual(['one']);

		expect(bindA.push(['one', 'two'])).toBe(true);
		expect(adoptedB.at(-1)).toEqual(['one', 'two']);
		expect(adoptedA).toEqual([]);
	});

	it('lets the room win a bootstrap load but publishes a user load', async () => {
		const hub = createMemoryHub();
		const a = join(hub);
		const bindA = bindDocument(a, adapter, {
			getLocalModel: () => ['room'],
			onRemoteModel: () => {},
		});
		await new Promise((r) => setTimeout(r, 5));
		expect(bindA.handleLoad('bootstrap', ['other'])).toBe(true);
		expect(a.doc.getArray('titles').toJSON()).toEqual(['room']);
		expect(bindA.handleLoad('user', ['opened'])).toBe(false);
		expect(a.doc.getArray('titles').toJSON()).toEqual(['opened']);
	});

	it('refuses writes from a viewer', async () => {
		const hub = createMemoryHub();
		const viewer = join(hub, 'viewer');
		const bind = bindDocument(viewer, adapter, {
			getLocalModel: () => ['x'],
			onRemoteModel: () => {},
		});
		await new Promise((r) => setTimeout(r, 5));
		expect(bind.push(['y'])).toBe(false);
		expect(viewer.doc.getArray('titles').length).toBe(0);
		bind.dispose();
	});
});
