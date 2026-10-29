import * as Y from 'yjs';
import { describe, expect, it } from 'vitest';
import { createChatStore } from './chat';
import { createMeetingStore } from './meetings';
import { createRoster } from './roster';
import { type UpdateActor, validateTeamsUpdate } from './validate-update';

const ann: UpdateActor = { id: 'ann' };
const bob: UpdateActor = { id: 'bob' };
const cy: UpdateActor = { id: 'cy' };
const mod: UpdateActor = { id: 'mod', moderator: true };

const copy = (from: Y.Doc): Y.Doc => {
	const doc = new Y.Doc();
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(from));
	return doc;
};

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

const HOUR = 3_600_000;

describe('validateTeamsUpdate: meetings', () => {
	function seeded(): { server: Y.Doc; id: string } {
		const server = new Y.Doc();
		server.clientID = 1;
		const store = createMeetingStore(server, { id: 'ann', name: 'Ann' });
		const meeting = store.create({
			title: 'Standup',
			start: Date.now() + HOUR,
			end: Date.now() + 2 * HOUR,
		})!;
		return { server, id: meeting.id };
	}

	/** The bytes of one change made on a fresh copy of the server document. */
	const change = (server: Y.Doc, fn: (store: ReturnType<typeof createMeetingStore>) => void) => {
		const doc = copy(server);
		const store = createMeetingStore(doc, { id: 'ann', name: 'Ann' });
		return capture(doc, () => fn(store));
	};

	it('accepts the organiser creating, editing and cancelling', () => {
		const { server, id } = seeded();
		const edit = change(server, (s) => void s.update(id, { title: 'Daily' }));
		const cancel = change(server, (s) => void s.cancel(id));
		const fresh = change(
			server,
			(s) =>
				void s.create({ title: 'Retro', start: Date.now() + HOUR, end: Date.now() + 2 * HOUR }),
		);
		for (const update of [edit, cancel, fresh])
			expect(validateTeamsUpdate(server, update, ann)).toEqual({ ok: true });
	});

	it('refuses a meeting created in somebody else name', () => {
		const { server } = seeded();
		const doc = copy(server);
		const update = capture(
			doc,
			() =>
				void createMeetingStore(doc, { id: 'ann', name: 'Ann' }).create({
					title: 'Forged',
					start: Date.now() + HOUR,
					end: Date.now() + 2 * HOUR,
				}),
		);
		expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['forged-author']);
	});

	it('refuses other people editing or cancelling, and lets a moderator do both', () => {
		const { server, id } = seeded();
		// Written as the organiser would, then presented as bob.
		const edit = change(server, (s) => void s.update(id, { title: 'Hijacked' }));
		const cancel = change(server, (s) => void s.cancel(id));
		for (const update of [edit, cancel]) {
			expect(codes(validateTeamsUpdate(server, update, bob))).toEqual(['foreign-row']);
			expect(validateTeamsUpdate(server, update, mod)).toEqual({ ok: true });
		}
	});
});

describe('validateTeamsUpdate: roster', () => {
	function seeded(): Y.Doc {
		const server = new Y.Doc();
		server.clientID = 1;
		createRoster(server, { id: 'ann', name: 'Ann' }).register();
		return server;
	}

	it('accepts registering yourself', () => {
		const server = seeded();
		const doc = copy(server);
		const update = capture(doc, () => createRoster(doc, { id: 'bob', name: 'Bob' }).register());
		expect(validateTeamsUpdate(server, update, bob)).toEqual({ ok: true });
	});

	it('refuses writing the entry of somebody else', () => {
		const server = seeded();
		const doc = copy(server);
		const update = capture(doc, () => createRoster(doc, { id: 'bob', name: 'Bob' }).register());
		expect(codes(validateTeamsUpdate(server, update, cy))).toEqual(['foreign-row']);
	});

	it('refuses renaming somebody else, lets a moderator remove an entry only', () => {
		const server = seeded();
		const doc = copy(server);
		const rename = capture(doc, () =>
			doc.getMap<Y.Map<unknown>>('teams:roster').get('ann')!.set('name', 'Evil'),
		);
		expect(codes(validateTeamsUpdate(server, rename, bob))).toEqual(['foreign-row']);
		const doc2 = copy(server);
		const remove = capture(doc2, () => doc2.getMap('teams:roster').delete('ann'));
		expect(validateTeamsUpdate(server, remove, mod)).toEqual({ ok: true });
		expect(codes(validateTeamsUpdate(server, remove, bob))).toEqual(['foreign-row']);
		expect(validateTeamsUpdate(server, remove, ann)).toEqual({ ok: true });
	});
});

describe('validateTeamsUpdate: chat members', () => {
	function seeded(): Y.Doc {
		const server = new Y.Doc();
		server.clientID = 1;
		const members = server.getMap<boolean>('teams:chat-members');
		members.set('ann', true);
		members.set('bob', true);
		return server;
	}
	const change = (server: Y.Doc, fn: (m: Y.Map<boolean>) => void): Uint8Array => {
		const doc = copy(server);
		return capture(doc, () => fn(doc.getMap<boolean>('teams:chat-members')));
	};

	it('lets the first writer seed a new chat that lists them', () => {
		const empty = new Y.Doc();
		const update = change(empty, (m) => {
			m.set('ann', true);
			m.set('bob', true);
		});
		expect(validateTeamsUpdate(empty, update, ann)).toEqual({ ok: true });
		expect(codes(validateTeamsUpdate(empty, update, cy))).toEqual(['not-member', 'not-member']);
	});

	it('lets a member add people and refuses a non-member', () => {
		const server = seeded();
		const add = change(server, (m) => m.set('cy', true));
		expect(validateTeamsUpdate(server, add, bob)).toEqual({ ok: true });
		expect(codes(validateTeamsUpdate(server, add, cy))).toEqual(['not-member']);
		expect(validateTeamsUpdate(server, add, mod)).toEqual({ ok: true });
	});

	it('lets a member leave but not remove others; a moderator removes anyone', () => {
		const server = seeded();
		const leave = change(server, (m) => m.delete('bob'));
		expect(validateTeamsUpdate(server, leave, bob)).toEqual({ ok: true });
		expect(codes(validateTeamsUpdate(server, leave, ann))).toEqual(['foreign-row']);
		expect(validateTeamsUpdate(server, leave, mod)).toEqual({ ok: true });
	});

	it('refuses a member entry that is not true', () => {
		const server = seeded();
		const update = change(server, (m) => (m as Y.Map<unknown>).set('cy', 'yes'));
		expect(codes(validateTeamsUpdate(server, update, ann))).toEqual(['malformed-row']);
	});
});

describe('validateTeamsUpdate: content limits', () => {
	function seeded(): { server: Y.Doc; messageId: string } {
		const server = new Y.Doc();
		server.clientID = 1;
		const store = createChatStore(server, { id: 'ann', name: 'Ann' });
		store.createChannel({ id: 'general', name: 'General' });
		return { server, messageId: store.post('general', { text: 'hi' })!.id };
	}

	it('refuses a message longer than the limit and accepts one at the limit', () => {
		const { server } = seeded();
		const post = (text: string): Uint8Array => {
			const doc = copy(server);
			return capture(doc, () => {
				const m = new Y.Map<unknown>();
				m.set('authorId', 'ann');
				m.set('authorName', 'Ann');
				m.set('text', text);
				m.set('ts', Date.now());
				doc.getMap<Y.Map<unknown>>('teams:messages:general').set(`big-${text.length}`, m);
			});
		};
		expect(
			validateTeamsUpdate(server, post('x'.repeat(100)), ann, { maxMessageChars: 100 }),
		).toEqual({
			ok: true,
		});
		expect(
			codes(validateTeamsUpdate(server, post('x'.repeat(101)), ann, { maxMessageChars: 100 })),
		).toEqual(['content-too-large']);
		// The default is 32 KB.
		expect(codes(validateTeamsUpdate(server, post('x'.repeat(32_769)), ann))).toEqual([
			'content-too-large',
		]);
		expect(validateTeamsUpdate(server, post('x'.repeat(32_768)), ann)).toEqual({ ok: true });
	});

	it('refuses an edit that grows a message past the limit', () => {
		const { server, messageId } = seeded();
		const doc = copy(server);
		const update = capture(doc, () =>
			doc
				.getMap<Y.Map<unknown>>('teams:messages:general')
				.get(messageId)!
				.set('text', 'y'.repeat(60)),
		);
		expect(codes(validateTeamsUpdate(server, update, ann, { maxMessageChars: 50 }))).toEqual([
			'content-too-large',
		]);
	});

	it('refuses more reactions on one message than allowed', () => {
		const { server, messageId } = seeded();
		const doc = copy(server);
		const update = capture(doc, () => {
			const reactions = doc.getMap<boolean>('teams:reactions');
			for (const emoji of ['a', 'b', 'c']) reactions.set(`${messageId}|${emoji}|ann`, true);
		});
		expect(validateTeamsUpdate(server, update, ann, { maxReactionsPerMessage: 3 })).toEqual({
			ok: true,
		});
		expect(codes(validateTeamsUpdate(server, update, ann, { maxReactionsPerMessage: 2 }))).toEqual([
			'content-too-large',
		]);
	});
});
