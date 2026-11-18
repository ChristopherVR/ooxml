import * as Y from 'yjs';
import { describe, expect, it } from 'vitest';
import { createChatStore } from './chat';
import { createChatRooms } from './chat-rooms';
import { type ChatMember, createChatDirectory, directChatId, legacyDirectChatId } from './chats';
import { type UpdateActor, validateTeamsUpdate } from './validate-update';

const person = (id: string): ChatMember => ({ id, name: id });
const [ann, bob, cy, eve] = ['ann', 'bob', 'cy', 'eve'].map(person) as [
	ChatMember,
	ChatMember,
	ChatMember,
	ChatMember,
];
const actor = (id: string, moderator = false): UpdateActor => ({ id, moderator });

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
const rowOf = (doc: Y.Doc, id: string): Y.Map<unknown> =>
	doc.getMap<Y.Map<unknown>>('teams:chats').get(id)!;

/** What `by` is told when it sends the change `fn` makes to a copy of the server's document. */
function verdict(
	server: Y.Doc,
	by: UpdateActor,
	fn: (doc: Y.Doc) => void,
): ReturnType<typeof validateTeamsUpdate> {
	const doc = copy(server);
	return validateTeamsUpdate(
		server,
		capture(doc, () => fn(doc)),
		by,
	);
}

function serverWithChats(): { server: Y.Doc; group: string; direct: string } {
	const server = new Y.Doc();
	server.clientID = 1;
	const directory = createChatDirectory(server, ann);
	const group = directory.createGroup({ others: [bob], name: 'Crew' })!.id;
	const direct = directory.ensureDirect(bob)!.id;
	return { server, group, direct };
}

function writeRow(doc: Y.Doc, id: string, fields: Record<string, unknown>): void {
	const row = new Y.Map<unknown>();
	for (const [key, value] of Object.entries(fields)) row.set(key, value);
	doc.getMap('teams:chats').set(id, row);
}

describe('validateTeamsUpdate: creating directory rows', () => {
	it('accepts the rows the client writes', () => {
		const server = new Y.Doc();
		expect(
			verdict(server, actor('ann'), (doc) => {
				const directory = createChatDirectory(doc, ann);
				directory.createGroup({ others: [bob, cy], name: 'Crew', meetingId: 'm-1' });
				directory.ensureDirect(bob);
			}),
		).toEqual({ ok: true });
	});

	it('refuses a row created in somebody else name or under their group id', () => {
		const server = new Y.Doc();
		const written = new Y.Doc();
		const id = createChatDirectory(written, ann).createGroup({ others: [bob] })!.id;
		const update = Y.encodeStateAsUpdate(written);
		expect(codes(validateTeamsUpdate(server, update, actor('bob')))).toEqual(['forged-author']);
		expect(
			codes(
				verdict(server, actor('eve'), (doc) =>
					writeRow(doc, id, {
						kind: 'group',
						name: '',
						members: [eve, ann],
						createdAt: 1,
						createdBy: 'eve',
					}),
				),
			),
		).toEqual(['forged-author']);
	});

	it('refuses a group id that is not random and bound to its creator', () => {
		expect(
			codes(
				verdict(new Y.Doc(), actor('ann'), (doc) =>
					writeRow(doc, 'g-chosen', {
						kind: 'group',
						name: '',
						members: [ann, bob],
						createdAt: 1,
						createdBy: 'ann',
					}),
				),
			),
		).toEqual(['forged-author']);
	});

	it('accepts a 1:1 row only under the id of its two people', () => {
		const row = (createdBy: string, members: ChatMember[]) => ({
			kind: 'direct',
			name: '',
			members,
			createdAt: 1,
			createdBy,
		});
		const server = new Y.Doc();
		const pair = directChatId('ann', 'bob');
		const refused = (id: string, by: string, members: ChatMember[]): string[] =>
			codes(verdict(server, actor(by), (doc) => writeRow(doc, id, row(by, members))));
		expect(refused(pair, 'ann', [ann, bob])).toEqual([]);
		expect(refused(pair, 'eve', [eve, bob])).toEqual(['malformed-row']);
		expect(refused(pair, 'eve', [ann, bob])).toEqual(['not-member']);
		expect(refused(legacyDirectChatId('ann', 'bob'), 'ann', [ann, bob])).toEqual(['malformed-row']);
	});

	it('refuses rows a reader would have to repair', () => {
		const server = new Y.Doc();
		const id = directChatId('ann', 'bob');
		const base = { name: '', createdAt: 1, createdBy: 'ann' };
		for (const fields of [
			{ ...base, kind: 'room', members: [ann, bob] },
			{ ...base, kind: 'direct', members: [ann, bob, { id: 'a,b', name: 'Bad' }] },
			{ ...base, kind: 'direct', members: 'ann,bob' },
		])
			expect(codes(verdict(server, actor('ann'), (doc) => writeRow(doc, id, fields)))).toEqual([
				'malformed-row',
			]);
		const notAMap = copy(server);
		const update = capture(notAMap, () => notAMap.getMap('teams:chats').set(id, 'row'));
		expect(codes(validateTeamsUpdate(server, update, actor('ann')))).toEqual(['malformed-row']);
	});
});

describe('validateTeamsUpdate: changing directory rows', () => {
	it('keeps people who are not members from touching a row', () => {
		const { server, group, direct } = serverWithChats();
		for (const id of [group, direct])
			expect(
				codes(
					verdict(server, actor('eve'), (doc) => {
						const row = rowOf(doc, id);
						doc.transact(() => {
							row.set('kind', 'group');
							row.set('members', [...(row.get('members') as ChatMember[]), eve]);
						});
					}),
				),
			).toEqual(['not-member']);
		expect(
			codes(verdict(server, actor('eve', true), (doc) => rowOf(doc, group).set('name', 'Mine'))),
		).toEqual(['not-member']);
	});

	it('never lets a row change kind or origin, or disappear', () => {
		const { server, group } = serverWithChats();
		const change = (fn: (row: Y.Map<unknown>) => void): string[] =>
			codes(verdict(server, actor('ann'), (doc) => fn(rowOf(doc, group))));
		expect(change((row) => row.set('kind', 'direct'))).toEqual(['malformed-row']);
		expect(change((row) => row.set('createdBy', 'bob'))).toEqual(['foreign-row']);
		expect(change((row) => row.set('createdAt', 5))).toEqual(['foreign-row']);
		expect(change((row) => row.set('meetingId', 'm-2'))).toEqual(['foreign-row']);
		expect(
			codes(verdict(server, actor('ann'), (doc) => doc.getMap('teams:chats').delete(group))),
		).toEqual(['foreign-row']);
	});

	it('keeps a 1:1 row as it was created', () => {
		const { server, direct } = serverWithChats();
		const change = (fn: (row: Y.Map<unknown>) => void): string[] =>
			codes(verdict(server, actor('bob'), (doc) => fn(rowOf(doc, direct))));
		expect(change((row) => row.set('kind', 'group'))).toEqual(['malformed-row']);
		expect(change((row) => row.set('name', 'Renamed'))).toEqual(['foreign-row']);
		expect(change((row) => row.set('members', [ann, bob, cy]))).toEqual(['malformed-row']);
	});

	it('accepts both people opening the same 1:1 chat at once', () => {
		const { server, direct } = serverWithChats();
		const offline = new Y.Doc();
		const update = capture(offline, () => createChatDirectory(offline, bob).ensureDirect(ann));
		expect(createChatDirectory(offline, bob).ensureDirect(ann)!.id).toBe(direct);
		expect(validateTeamsUpdate(server, update, actor('bob'))).toEqual({ ok: true });
	});

	it('lets a member rename, add people and leave, but not remove others', () => {
		const { server, group } = serverWithChats();
		const by = (who: ChatMember, fn: (d: ReturnType<typeof createChatDirectory>) => void) =>
			verdict(server, actor(who.id), (doc) => fn(createChatDirectory(doc, who)));
		expect(by(bob, (d) => d.rename(group, 'Renamed'))).toEqual({ ok: true });
		expect(by(bob, (d) => d.add(group, [cy]))).toEqual({ ok: true });
		expect(by(bob, (d) => d.remove(group, 'bob'))).toEqual({ ok: true });
		expect(codes(by(bob, (d) => d.remove(group, 'ann')))).toEqual(['foreign-row']);
		expect(
			verdict(server, actor('bob', true), (doc) =>
				createChatDirectory(doc, bob).remove(group, 'ann'),
			),
		).toEqual({ ok: true });
	});

	it('leaves rows from before these rules alone until somebody changes them', () => {
		const server = new Y.Doc();
		writeRow(server, 'g-g-1-1', {
			kind: 'group',
			name: 'Old',
			members: [ann, bob],
			createdAt: 1,
			createdBy: 'ann',
		});
		expect(
			verdict(server, actor('bob'), (doc) => rowOf(doc, 'g-g-1-1').set('name', 'New')),
		).toEqual({ ok: true });
	});
});

describe('the client and the room member list', () => {
	function roomsFor(user: ChatMember, doc: Y.Doc) {
		const chatDocs = new Map<string, Y.Doc>();
		const rooms = createChatRooms({
			doc,
			base: createChatStore(doc, user),
			user,
			openRoom: (chat, chatDoc) => {
				chatDocs.set(chat.id, chatDoc);
				return { destroy: () => undefined };
			},
		});
		const members = (id: string): string[] =>
			[...(chatDocs.get(id)?.getMap<boolean>('teams:chat-members').keys() ?? [])].sort();
		return { rooms, members };
	}

	it('writes the members of a group its own user created or added', () => {
		const { rooms, members } = roomsFor(ann, new Y.Doc());
		const chat = rooms.createGroup({ others: [bob] })!;
		expect(members(chat.id)).toEqual(['ann', 'bob']);
		rooms.addPeople(chat.id, [cy]);
		expect(members(chat.id)).toEqual(['ann', 'bob', 'cy']);
		rooms.leaveChat(chat.id);
		expect(members(chat.id)).toEqual(['bob', 'cy']);
		rooms.destroy();
	});

	it('does not copy into the room what somebody else wrote in the directory', () => {
		const doc = new Y.Doc();
		const { rooms, members } = roomsFor(bob, doc);
		const creator = createChatDirectory(doc, ann);
		const group = creator.createGroup({ others: [bob] })!;
		const direct = creator.ensureDirect(bob)!;
		const row = rowOf(doc, group.id);
		row.set('members', [...(row.get('members') as ChatMember[]), eve]);
		expect(rooms.chat(group.id)?.members.map((m) => m.id)).toContain('eve');
		expect(members(group.id)).toEqual([]);
		expect(members(direct.id)).toEqual([]);
		rooms.destroy();
	});

	it('keeps using a 1:1 chat made before the id changed', () => {
		const doc = new Y.Doc();
		const old = legacyDirectChatId('ann', 'bob');
		writeRow(doc, old, {
			kind: 'direct',
			name: '',
			members: [ann, bob],
			createdAt: 1,
			createdBy: 'ann',
		});
		expect(createChatDirectory(doc, bob).ensureDirect(ann)?.id).toBe(old);
	});
});
