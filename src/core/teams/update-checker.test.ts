import * as Y from 'yjs';
import { describe, expect, it } from 'vitest';
import { createChatStore } from './chat';
import { createUpdateChecker } from './update-checker';
import { type UpdateActor, validateTeamsUpdate } from './validate-update';

const ann: UpdateActor = { id: 'ann' };
const bob: UpdateActor = { id: 'bob' };

/** The update a client produces while running `fn`. */
function capture(doc: Y.Doc, fn: () => void): Uint8Array {
	const updates: Uint8Array[] = [];
	const on = (u: Uint8Array): void => void updates.push(u);
	doc.on('update', on);
	fn();
	doc.off('update', on);
	return Y.mergeUpdates(updates);
}

/** A server room and a client of `user` that starts from it and sends everything it writes. */
function room() {
	const server = new Y.Doc();
	const client = (user: string) => {
		const doc = new Y.Doc();
		Y.applyUpdate(doc, Y.encodeStateAsUpdate(server));
		server.on('update', (u: Uint8Array) => Y.applyUpdate(doc, u));
		return { doc, store: createChatStore(doc, { id: user, name: user }) };
	};
	return { server, client };
}

describe('createUpdateChecker', () => {
	it('gives the verdicts validateTeamsUpdate gives, over a run of updates', () => {
		const { server, client } = room();
		const checker = createUpdateChecker(server);
		const a = client('ann');
		const b = client('bob');
		const send = (from: { doc: Y.Doc }, actor: UpdateActor, fn: () => void) => {
			const update = capture(from.doc, fn);
			const expected = validateTeamsUpdate(server, update, actor);
			const verdict = checker.validate(update, actor);
			expect(verdict).toEqual(expected);
			if (verdict.ok) Y.applyUpdate(server, update);
			return verdict.ok;
		};
		expect(send(a, ann, () => a.store.createChannel({ id: 'general', name: 'General' }))).toBe(
			true,
		);
		let id = '';
		expect(send(a, ann, () => (id = a.store.post('general', { text: 'one' })!.id))).toBe(true);
		expect(send(b, bob, () => b.store.post('general', { text: 'two' }))).toBe(true);
		// Bob edits Ann's message: refused, and the copy must not keep it.
		expect(
			send(b, bob, () =>
				b.doc.getMap<Y.Map<unknown>>('teams:messages:general').get(id)!.set('text', 'x'),
			),
		).toBe(false);
		// The next update is judged against the room as it is, without Bob's refused edit.
		const fresh = client('ann');
		expect(send(fresh, ann, () => fresh.store.edit('general', id, 'one, edited'))).toBe(true);
		expect(server.getMap<Y.Map<unknown>>('teams:messages:general').get(id)!.get('text')).toBe(
			'one, edited',
		);
		checker.destroy();
	});

	it('follows changes the room makes itself', () => {
		const { server, client } = room();
		const checker = createUpdateChecker(server);
		const a = client('ann');
		const create = capture(a.doc, () => a.store.createChannel({ id: 'general', name: 'General' }));
		expect(checker.validate(create, ann).ok).toBe(true);
		Y.applyUpdate(server, create);
		// A row written by the room itself (as a server seeds a chat), then an edit that depends on it.
		server.getMap('teams:chat-members').set('ann', true);
		const pin = capture(a.doc, () => a.store.post('general', { text: 'after the seed' }));
		expect(checker.validate(pin, ann)).toEqual({ ok: true });
		checker.destroy();
	});

	it('refuses an update that depends on state the room lacks, then recovers', () => {
		const { server, client } = room();
		const checker = createUpdateChecker(server);
		const a = client('ann');
		const first = capture(a.doc, () => a.store.createChannel({ id: 'general', name: 'General' }));
		const second = capture(a.doc, () => a.store.post('general', { text: 'hi' }));
		expect(checker.validate(second, ann)).toEqual({
			ok: false,
			violations: [{ code: 'incomplete-update' }],
		});
		expect(checker.validate(first, ann).ok).toBe(true);
		Y.applyUpdate(server, first);
		expect(checker.validate(second, ann).ok).toBe(true);
		checker.destroy();
	});

	it('rehearses an update that Yjs cannot apply without touching the room', () => {
		const { server } = room();
		const checker = createUpdateChecker(server);
		expect(() => checker.rehearse(new Uint8Array([1, 1, 5]))).toThrow();
		expect(server.store.clients.size).toBe(0);
		checker.destroy();
	});
});
