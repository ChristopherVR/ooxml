import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { describe, expect, it } from 'vitest';
import { createChatStore } from './chat';
import { validateTeamsSyncMessage, validateTeamsUpdate, type UpdateActor } from './validate-update';

const ann: UpdateActor = { id: 'ann' };
const bob: UpdateActor = { id: 'bob' };
const mod: UpdateActor = { id: 'mod', moderator: true };

/** A client document kept in step with the "server" document, as a real client would be. */
function pair(): { server: Y.Doc; client: (id: string) => Y.Doc } {
	const server = new Y.Doc();
	const client = (): Y.Doc => {
		const doc = new Y.Doc();
		Y.applyUpdate(doc, Y.encodeStateAsUpdate(server));
		return doc;
	};
	return { server, client };
}

/** The update a client produces while running `fn`. */
function capture(doc: Y.Doc, fn: () => void): Uint8Array {
	const updates: Uint8Array[] = [];
	const on = (u: Uint8Array): void => void updates.push(u);
	doc.on('update', on);
	fn();
	doc.off('update', on);
	return Y.mergeUpdates(updates);
}

const codes = (v: ReturnType<typeof validateTeamsUpdate>): string[] =>
	v.ok ? [] : v.violations.map((x) => x.code);

/** Seed the server with a channel and one message by ann; returns its id. */
function seeded(): { server: Y.Doc; messageId: string; fresh: () => Y.Doc } {
	const { server, client } = pair();
	const doc = client('ann');
	const store = createChatStore(doc, { id: 'ann', name: 'Ann' });
	store.createChannel({ id: 'general', name: 'General' });
	const message = store.post('general', { text: 'hello' })!;
	Y.applyUpdate(server, Y.encodeStateAsUpdate(doc));
	return { server, messageId: message.id, fresh: () => client('x') };
}

const rows = (doc: Y.Doc): Y.Map<Y.Map<unknown>> => doc.getMap('teams:messages:general');

describe('validateTeamsUpdate: message rows', () => {
	it('accepts what the real store writes for its own user', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const store = createChatStore(doc, { id: 'bob', name: 'Bob' });
		const update = capture(doc, () => {
			store.post('general', { text: 'hi' });
			store.react('x', '👍');
		});
		expect(validateTeamsUpdate(server, update, bob)).toEqual({ ok: true });
	});

	it('accepts an author editing and soft-deleting their own message', () => {
		const { server, messageId, fresh } = seeded();
		const doc = fresh();
		const other = fresh();
		const store = createChatStore(doc, { id: 'ann', name: 'ann' });
		expect(
			validateTeamsUpdate(
				server,
				capture(doc, () => store.edit('general', messageId, 'edited')),
				ann,
			),
		).toEqual({ ok: true });
		expect(
			validateTeamsUpdate(
				server,
				capture(other, () =>
					createChatStore(other, { id: 'ann', name: 'ann' }).remove('general', messageId),
				),
				ann,
			),
		).toEqual({ ok: true });
	});

	it('rejects a message posted under another author', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const update = capture(doc, () => {
			const row = new Y.Map<unknown>();
			row.set('authorId', 'ann');
			row.set('authorName', 'Ann');
			row.set('text', 'I am ann');
			row.set('ts', 1);
			row.set('deleted', false);
			rows(doc).set('forged', row);
		});
		const verdict = validateTeamsUpdate(server, update, bob);
		expect(codes(verdict)).toEqual(['forged-author']);
		expect(verdict).toMatchObject({
			violations: [{ map: 'teams:messages:general', key: 'forged' }],
		});
	});

	it('rejects a row with no author or a non-map row', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const noAuthor = capture(doc, () => rows(doc).set('a', new Y.Map<unknown>()));
		expect(codes(validateTeamsUpdate(server, noAuthor, bob))).toEqual(['forged-author']);
		const other = fresh();
		const notMap = capture(other, () =>
			other.getMap<unknown>('teams:messages:general').set('b', 'x'),
		);
		expect(codes(validateTeamsUpdate(server, notMap, bob))).toEqual(['malformed-row']);
	});

	it('rejects editing, deleting and removing somebody else message', () => {
		const { server, messageId, fresh } = seeded();
		const edit = fresh();
		expect(
			codes(
				validateTeamsUpdate(
					server,
					capture(edit, () => rows(edit).get(messageId)!.set('text', 'tampered')),
					bob,
				),
			),
		).toEqual(['foreign-row']);
		const soft = fresh();
		expect(
			codes(
				validateTeamsUpdate(
					server,
					capture(soft, () => rows(soft).get(messageId)!.set('deleted', true)),
					bob,
				),
			),
		).toEqual(['foreign-row']);
		const hard = fresh();
		expect(
			codes(
				validateTeamsUpdate(
					server,
					capture(hard, () => rows(hard).delete(messageId)),
					bob,
				),
			),
		).toEqual(['foreign-row']);
	});

	it('rejects replacing somebody else row with a copy that claims the actor', () => {
		const { server, messageId, fresh } = seeded();
		const doc = fresh();
		const update = capture(doc, () => {
			const row = new Y.Map<unknown>();
			row.set('authorId', 'bob');
			row.set('text', 'mine now');
			rows(doc).set(messageId, row);
		});
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['foreign-row']);
	});

	it('rejects handing your own row to somebody else', () => {
		const { server, messageId, fresh } = seeded();
		const doc = fresh();
		const update = capture(doc, () => rows(doc).get(messageId)!.set('authorId', 'bob'));
		expect(codes(validateTeamsUpdate(server, update, ann))).toEqual(['author-changed']);
	});

	it('catches a forged row hidden inside a merged update of valid ones', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const store = createChatStore(doc, { id: 'bob', name: 'bob' });
		const update = capture(doc, () => {
			store.post('general', { text: 'fine' });
			const row = new Y.Map<unknown>();
			row.set('authorId', 'ann');
			rows(doc).set('forged', row);
			store.post('general', { text: 'also fine' });
		});
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['forged-author']);
	});

	it('lets a moderator hide or remove a message, but not rewrite it', () => {
		const { server, messageId, fresh } = seeded();
		const hide = fresh();
		expect(
			validateTeamsUpdate(
				server,
				capture(hide, () => rows(hide).get(messageId)!.set('deleted', true)),
				mod,
			),
		).toEqual({ ok: true });
		const remove = fresh();
		expect(
			validateTeamsUpdate(
				server,
				capture(remove, () => rows(remove).delete(messageId)),
				mod,
			),
		).toEqual({ ok: true });
		const rewrite = fresh();
		expect(
			codes(
				validateTeamsUpdate(
					server,
					capture(rewrite, () => rows(rewrite).get(messageId)!.set('text', 'tampered')),
					mod,
				),
			),
		).toEqual(['foreign-row']);
		const hideAndRewrite = fresh();
		expect(
			codes(
				validateTeamsUpdate(
					server,
					capture(hideAndRewrite, () => {
						rows(hideAndRewrite).get(messageId)!.set('deleted', true);
						rows(hideAndRewrite).get(messageId)!.set('text', 'x');
					}),
					mod,
				),
			),
		).toEqual(['foreign-row']);
	});

	it('a moderator still cannot post as somebody else', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const update = capture(doc, () => {
			const row = new Y.Map<unknown>();
			row.set('authorId', 'ann');
			rows(doc).set('forged', row);
		});
		expect(codes(validateTeamsUpdate(server, update, mod))).toEqual(['forged-author']);
	});

	it('lets concurrent clients write the same tab discussion root as themselves', () => {
		const { server, fresh } = seeded();
		const first = fresh();
		const second = fresh();
		const write = (doc: Y.Doc, id: string): Uint8Array =>
			capture(doc, () => {
				const row = new Y.Map<unknown>();
				row.set('authorId', id);
				rows(doc).set('tab:abc', row);
			});
		const a = write(first, 'ann');
		const b = write(second, 'bob');
		expect(validateTeamsUpdate(server, a, ann)).toEqual({ ok: true });
		Y.applyUpdate(server, a);
		expect(validateTeamsUpdate(server, b, bob)).toEqual({ ok: true });
		// ...but not under another name.
		const forger = fresh();
		const forged = capture(forger, () => {
			const row = new Y.Map<unknown>();
			row.set('authorId', 'ann');
			rows(forger).set('tab:other', row);
		});
		expect(codes(validateTeamsUpdate(server, forged, bob))).toEqual(['forged-author']);
	});
});

describe('validateTeamsUpdate: reactions', () => {
	it('accepts own reactions and rejects forged or removed ones', () => {
		const { server, fresh } = seeded();
		const own = fresh();
		expect(
			validateTeamsUpdate(
				server,
				capture(own, () => createChatStore(own, { id: 'ann', name: 'ann' }).react('m1', '👍')),
				ann,
			),
		).toEqual({ ok: true });
		const forged = fresh();
		const forgedUpdate = capture(forged, () =>
			forged.getMap('teams:reactions').set('m1|👍|ann', Date.now()),
		);
		expect(codes(validateTeamsUpdate(server, forgedUpdate, bob))).toEqual(['foreign-reaction']);
		// Put ann's reaction on the server, then try to remove it as bob.
		Y.applyUpdate(server, forgedUpdate);
		const removal = fresh();
		const removeUpdate = capture(removal, () =>
			removal.getMap('teams:reactions').delete('m1|👍|ann'),
		);
		expect(codes(validateTeamsUpdate(server, removeUpdate, bob))).toEqual(['foreign-reaction']);
		expect(validateTeamsUpdate(server, removeUpdate, mod)).toEqual({ ok: true });
		expect(validateTeamsUpdate(server, removeUpdate, ann)).toEqual({ ok: true });
	});

	it('rejects malformed reaction keys and a moderator adding one in another name', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const bad = capture(doc, () => doc.getMap('teams:reactions').set('nonsense', 1));
		expect(codes(validateTeamsUpdate(server, bad, bob))).toEqual(['malformed-row']);
		const other = fresh();
		const add = capture(other, () => other.getMap('teams:reactions').set('m1|👍|ann', 1));
		expect(codes(validateTeamsUpdate(server, add, mod))).toEqual(['foreign-reaction']);
	});
});

describe('validateTeamsUpdate: Yjs semantics and safety', () => {
	it('does not modify the room document', () => {
		const { server, fresh } = seeded();
		const before = Y.encodeStateAsUpdate(server);
		const doc = fresh();
		const update = capture(doc, () =>
			createChatStore(doc, { id: 'bob', name: 'bob' }).post('general', { text: 'x' }),
		);
		validateTeamsUpdate(server, update, bob);
		validateTeamsUpdate(server, update, ann);
		expect(Y.encodeStateAsUpdate(server)).toEqual(before);
	});

	it('passes changes outside the checked maps', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const update = capture(doc, () => {
			const row = new Y.Map<unknown>();
			row.set('name', 'Bob');
			doc.getMap<Y.Map<unknown>>('teams:roster').set('bob', row);
		});
		expect(validateTeamsUpdate(server, update, bob)).toEqual({ ok: true });
	});

	it('refuses an update that depends on state the server lacks', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const store = createChatStore(doc, { id: 'bob', name: 'bob' });
		capture(doc, () => store.post('general', { text: 'first, never sent' }));
		const second = capture(doc, () => store.post('general', { text: 'second' }));
		expect(codes(validateTeamsUpdate(server, second, bob))).toEqual(['incomplete-update']);
	});

	it('refuses garbage and oversized updates', () => {
		const { server } = seeded();
		expect(codes(validateTeamsUpdate(server, new Uint8Array([255, 255, 255, 1]), bob))).toEqual([
			'invalid-update',
		]);
		expect(codes(validateTeamsUpdate(server, new Uint8Array(10), bob, { maxBytes: 5 }))).toEqual([
			'invalid-update',
		]);
	});

	it('does not accept a delete-set entry that removes a row it never saw as a tombstone', () => {
		// A deletion alone (no inserts) is a delete set; it must still be attributed to the row.
		const { server, messageId, fresh } = seeded();
		const doc = fresh();
		const update = capture(doc, () => rows(doc).delete(messageId));
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['foreign-row']);
		expect(validateTeamsUpdate(server, update, ann)).toEqual({ ok: true });
	});
});

describe('validateTeamsSyncMessage', () => {
	const frame = (kind: number, payload: Uint8Array): Uint8Array => {
		const e = encoding.createEncoder();
		encoding.writeVarUint(e, 0);
		encoding.writeVarUint(e, kind);
		encoding.writeVarUint8Array(e, payload);
		return encoding.toUint8Array(e);
	};

	it('validates update and step 2 messages and passes the rest', () => {
		const { server, fresh } = seeded();
		const doc = fresh();
		const update = capture(doc, () => {
			const row = new Y.Map<unknown>();
			row.set('authorId', 'ann');
			rows(doc).set('forged', row);
		});
		expect(codes(validateTeamsSyncMessage(server, frame(2, update), bob))).toEqual([
			'forged-author',
		]);
		expect(codes(validateTeamsSyncMessage(server, frame(1, update), bob))).toEqual([
			'forged-author',
		]);
		expect(validateTeamsSyncMessage(server, frame(0, Y.encodeStateVector(doc)), bob)).toEqual({
			ok: true,
		});
		const awareness = encoding.createEncoder();
		encoding.writeVarUint(awareness, 1);
		expect(validateTeamsSyncMessage(server, encoding.toUint8Array(awareness), bob)).toEqual({
			ok: true,
		});
		expect(codes(validateTeamsSyncMessage(server, new Uint8Array([0, 2]), bob))).toEqual([
			'invalid-update',
		]);
	});
});
