import * as Y from 'yjs';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createChatStore } from './chat';
import { createChatRooms } from './chat-rooms';
import {
	type ChatInfo,
	type ChatMember,
	createChatDirectory,
	createGroupChatId,
	directChatId,
	legacyDirectChatId,
} from './chats';
import { type UpdateActor, applyUpdateReverts, validateTeamsUpdate } from './validate-update';

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
		expect(change((row) => row.set('kind', 'direct'))).toEqual(['foreign-row']);
		expect(change((row) => row.set('members', [ann]))).toEqual(['foreign-row']);
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

	it('lets a member rename, add people and leave, and a moderator remove others', () => {
		const { server, group } = serverWithChats();
		const by = (who: ChatMember, fn: (d: ReturnType<typeof createChatDirectory>) => void) =>
			verdict(server, actor(who.id), (doc) => fn(createChatDirectory(doc, who)));
		expect(by(bob, (d) => d.rename(group, 'Renamed'))).toEqual({ ok: true });
		expect(by(bob, (d) => d.add(group, [cy]))).toEqual({ ok: true });
		expect(by(bob, (d) => d.remove(group, 'bob'))).toEqual({ ok: true });
		expect(
			verdict(server, actor('bob', true), (doc) =>
				createChatDirectory(doc, bob).remove(group, 'ann'),
			),
		).toEqual({ ok: true });
	});

	it('refuses member fields no client writes', () => {
		const { server, group } = serverWithChats();
		const change = (fn: (row: Y.Map<unknown>) => void): string[] =>
			codes(verdict(server, actor('bob'), (doc) => fn(rowOf(doc, group))));
		expect(change((row) => row.set('member:a,b', 'Bad'))).toEqual(['malformed-row']);
		expect(change((row) => row.set('member:cy', 7))).toEqual(['malformed-row']);
		expect(change((row) => row.set('member:cy', ''))).toEqual(['malformed-row']);
		expect(change((row) => row.set('member:cy', ' <b>x</b> '))).toEqual(['malformed-row']);
		expect(change((row) => row.set('extra', 1))).toEqual(['malformed-row']);
		expect(
			codes(
				verdict(server, actor('eve'), (doc) => {
					const id = createGroupChatId('eve');
					writeRow(doc, id, {
						kind: 'group',
						name: '',
						members: [eve, ann],
						createdAt: 1,
						createdBy: 'eve',
						'member:bob': 'Bob',
					});
				}),
			),
		).toEqual(['malformed-row']);
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

describe('changing a group from several clients at once', () => {
	/** Send `client`'s changes to `server` as the reference server would take them. */
	function push(server: Y.Doc, client: Y.Doc, by: UpdateActor): void {
		const update = Y.encodeStateAsUpdate(client, Y.encodeStateVector(server));
		const result = validateTeamsUpdate(server, update, by);
		if (!result.ok) throw new Error(`refused: ${codes(result).join()}`);
		Y.applyUpdate(server, update);
		if (result.reverts) applyUpdateReverts(server, result.reverts);
	}
	const pull = (server: Y.Doc, client: Y.Doc): void =>
		Y.applyUpdate(client, Y.encodeStateAsUpdate(server, Y.encodeStateVector(client)));
	const ids = (doc: Y.Doc, id: string): string[] =>
		createChatDirectory(doc, ann)
			.get(id)!
			.members.map((m) => m.id)
			.sort();

	/** Every order of client ids, because the older of two rewrites used to win by client id. */
	function scenarios(run: (clients: { annDoc: Y.Doc; bobDoc: Y.Doc }) => void): void {
		for (const [annId, bobId] of [
			[10, 20],
			[20, 10],
		] as const) {
			const annDoc = new Y.Doc();
			const bobDoc = new Y.Doc();
			annDoc.clientID = annId;
			bobDoc.clientID = bobId;
			run({ annDoc, bobDoc });
		}
	}
	function start(annDoc: Y.Doc, bobDoc: Y.Doc): { server: Y.Doc; group: string } {
		const server = new Y.Doc();
		const group = createChatDirectory(annDoc, ann).createGroup({ others: [bob] })!.id;
		push(server, annDoc, actor('ann'));
		pull(server, bobDoc);
		return { server, group };
	}

	it('keeps both of two people added at once', () =>
		scenarios(({ annDoc, bobDoc }) => {
			const { server, group } = start(annDoc, bobDoc);
			createChatDirectory(annDoc, ann).add(group, [cy]);
			createChatDirectory(bobDoc, bob).add(group, [eve]);
			push(server, annDoc, actor('ann'));
			push(server, bobDoc, actor('bob'));
			pull(server, annDoc);
			pull(server, bobDoc);
			for (const doc of [server, annDoc, bobDoc])
				expect(ids(doc, group)).toEqual(['ann', 'bob', 'cy', 'eve']);
		}));

	it('keeps an addition and a departure made at once', () =>
		scenarios(({ annDoc, bobDoc }) => {
			const { server, group } = start(annDoc, bobDoc);
			createChatDirectory(annDoc, ann).add(group, [cy]);
			createChatDirectory(bobDoc, bob).remove(group, 'bob');
			push(server, bobDoc, actor('bob'));
			push(server, annDoc, actor('ann'));
			pull(server, annDoc);
			pull(server, bobDoc);
			for (const doc of [server, annDoc, bobDoc]) expect(ids(doc, group)).toEqual(['ann', 'cy']);
		}));

	it('undoes what someone removed while offline changed, and keeps syncing them', () =>
		scenarios(({ annDoc, bobDoc }) => {
			const { server, group } = start(annDoc, bobDoc);
			createChatDirectory(annDoc, { ...ann }).remove(group, 'bob');
			push(server, annDoc, actor('ann', true));
			const directory = createChatDirectory(bobDoc, bob);
			directory.rename(group, 'Taken over');
			directory.add(group, [eve]);
			bobDoc.getMap('later').set('note', 'written offline too');
			push(server, bobDoc, actor('bob'));
			pull(server, annDoc);
			pull(server, bobDoc);
			for (const doc of [server, annDoc, bobDoc]) {
				expect(ids(doc, group)).toEqual(['ann']);
				expect(rowOf(doc, group).get('name')).toBe('');
			}
			expect(server.getMap('later').get('note')).toBe('written offline too');
			bobDoc.getMap('later').set('note', 'still syncing');
			push(server, bobDoc, actor('bob'));
			expect(server.getMap('later').get('note')).toBe('still syncing');
		}));

	it('undoes a removal by someone who is no longer a moderator', () => {
		const { server, group } = serverWithChats();
		const result = verdict(server, actor('bob'), (doc) =>
			createChatDirectory(doc, bob).remove(group, 'ann'),
		);
		expect(result).toEqual({
			ok: true,
			reverts: [{ chat: group, key: 'member:ann', value: undefined }],
		});
	});

	it('still refuses someone who was never in the group', () =>
		scenarios(({ annDoc, bobDoc }) => {
			const { server, group } = start(annDoc, bobDoc);
			const update = capture(bobDoc, () => rowOf(bobDoc, group).set('member:eve', 'eve'));
			expect(codes(validateTeamsUpdate(server, update, actor('eve')))).toEqual(['not-member']);
		}));

	it('reads the same people everywhere, whatever order the fields arrived in', () => {
		const { server, group } = serverWithChats();
		const first = copy(server);
		const second = copy(server);
		const added = capture(first, () => {
			const directory = createChatDirectory(first, ann);
			directory.add(group, [eve]);
			directory.add(group, [cy]);
		});
		Y.applyUpdate(second, added);
		const list = (doc: Y.Doc) => createChatDirectory(doc, ann).get(group)!.members;
		expect(list(second)).toEqual(list(first));
		expect(list(first).map((m) => m.id)).toEqual(['ann', 'bob', 'cy', 'eve']);
	});
});

describe('values a row cannot hold', () => {
	const big = BigInt(1);
	const groupOf = (doc: Y.Doc): string =>
		createChatDirectory(doc, ann).createGroup({ others: [bob] })!.id;

	it('refuses a new row that carries one, and does not throw', () => {
		const server = new Y.Doc();
		const client = copy(server);
		const id = createGroupChatId('ann');
		const update = capture(client, () =>
			writeRow(client, id, {
				kind: 'group',
				name: '',
				members: [{ ...ann, extra: big }, bob],
				createdAt: 1,
				createdBy: 'ann',
			}),
		);
		expect(codes(validateTeamsUpdate(server, update, actor('ann')))).toEqual(['malformed-row']);
	});

	it('refuses one written over a field of a row someone may edit', () => {
		const server = new Y.Doc();
		const id = groupOf(server);
		const client = copy(server);
		const update = capture(client, () => rowOf(client, id).set('name', big));
		expect(codes(validateTeamsUpdate(server, update, actor('ann')))).toEqual(['malformed-row']);
	});

	it('does not stop the people of a row that already holds one', () => {
		const server = new Y.Doc();
		const id = groupOf(server);
		rowOf(server, id).set('members', [{ ...ann, extra: big }, bob]);
		const client = copy(server);
		const update = capture(client, () => createChatDirectory(client, bob).rename(id, 'Crew'));
		expect(validateTeamsUpdate(server, update, actor('bob'))).toEqual({ ok: true });
	});

	it('answers with a refusal whatever goes wrong during the check', () => {
		const server = new Y.Doc();
		const client = copy(server);
		const update = capture(client, () => client.getMap('teams:roster').set('ann', new Y.Map()));
		const broken = {
			get id(): string {
				throw new Error('unreadable');
			},
		};
		expect(codes(validateTeamsUpdate(server, update, broken))).toEqual(['invalid-update']);
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
				return { synced: () => true, destroy: () => undefined };
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
		creator.add(group.id, [eve]);
		expect(rooms.chat(group.id)?.members.map((m) => m.id)).toContain('eve');
		expect(members(group.id)).toEqual([]);
		expect(members(direct.id)).toEqual([]);
		rooms.destroy();
	});

	describe('after the creator crashed before admitting anyone', () => {
		function reopen(user: ChatMember, doc: Y.Doc) {
			const chatDocs = new Map<string, Y.Doc>();
			const roomSynced = new Map<string, () => void>();
			const directoryListeners = new Set<() => void>();
			let directorySynced = false;
			const rooms = createChatRooms({
				doc,
				base: createChatStore(doc, user),
				user,
				directorySync: {
					synced: () => directorySynced,
					onSynced: (listener) => {
						directoryListeners.add(listener);
						return () => directoryListeners.delete(listener);
					},
				},
				openRoom: (chat, chatDoc, onSynced) => {
					chatDocs.set(chat.id, chatDoc);
					roomSynced.set(chat.id, onSynced);
					return { synced: () => false, destroy: () => undefined };
				},
			});
			return {
				rooms,
				members: (id: string): string[] =>
					[...(chatDocs.get(id)?.getMap<boolean>('teams:chat-members').keys() ?? [])].sort(),
				roomSynced: (id: string) => roomSynced.get(id)!(),
				directorySynced: () => {
					directorySynced = true;
					for (const listener of directoryListeners) listener();
				},
			};
		}

		it('admits the listed people once the room and the directory are current', () => {
			const doc = new Y.Doc();
			const id = createChatDirectory(doc, ann).createGroup({ others: [bob, cy] })!.id;
			createChatDirectory(doc, ann).remove(id, 'cy');
			const client = reopen(ann, doc);
			expect(client.members(id)).toEqual([]);
			client.roomSynced(id);
			expect(client.members(id)).toEqual([]);
			client.directorySynced();
			expect(client.members(id)).toEqual(['ann', 'bob']);
			client.rooms.destroy();
		});

		it('leaves it to the creator', () => {
			const doc = new Y.Doc();
			const id = createChatDirectory(doc, ann).createGroup({ others: [bob] })!.id;
			const client = reopen(bob, doc);
			client.roomSynced(id);
			client.directorySynced();
			expect(client.members(id)).toEqual([]);
			client.rooms.destroy();
		});
	});

	describe('after leaving while out of touch', () => {
		function device(user: ChatMember, doc: Y.Doc, restore?: (chat: ChatInfo) => Y.Doc) {
			const rooms = new Map<
				string,
				{ doc: Y.Doc; synced: boolean; onSynced: () => void; destroyed: boolean }
			>();
			const directoryListeners = new Set<() => void>();
			let directorySynced = false;
			const chat = createChatRooms({
				doc,
				base: createChatStore(doc, user),
				user,
				...(restore ? { createDoc: restore } : {}),
				directorySync: {
					synced: () => directorySynced,
					onSynced: (listener) => {
						directoryListeners.add(listener);
						return () => directoryListeners.delete(listener);
					},
				},
				openRoom: (info, chatDoc, onSynced) => {
					const room = { doc: chatDoc, synced: false, onSynced, destroyed: false };
					rooms.set(info.id, room);
					return {
						synced: () => room.synced,
						destroy: () => {
							room.destroyed = true;
						},
					};
				},
			});
			return {
				chat,
				room: (id: string) => rooms.get(id)!,
				roomSyncs: (id: string) => {
					const room = rooms.get(id)!;
					room.synced = true;
					room.onSynced();
				},
				directorySyncs: () => {
					directorySynced = true;
					for (const listener of directoryListeners) listener();
				},
			};
		}
		const listed = (doc: Y.Doc, id: string): boolean =>
			doc.getMap<boolean>('teams:chat-members').get(id) === true;

		beforeEach(() => void vi.useFakeTimers());
		afterEach(() => void vi.useRealTimers());

		function leftOffline() {
			const doc = new Y.Doc();
			const id = createChatDirectory(doc, ann).createGroup({ others: [bob, cy] })!.id;
			const client = device(bob, doc);
			client.room(id).doc.getMap<boolean>('teams:chat-members').set('bob', true);
			expect(client.chat.leaveChat(id)).toBe(true);
			return { doc, id, client };
		}

		it('keeps the room open until it syncs, however long that takes', () => {
			const { id, client } = leftOffline();
			vi.advanceTimersByTime(3_600_000);
			expect(client.room(id).destroyed).toBe(false);
			expect(client.chat.leavingChats()).toEqual([id]);
			expect(listed(client.room(id).doc, 'bob')).toBe(false);

			client.directorySyncs();
			client.roomSyncs(id);
			vi.advanceTimersByTime(1500);
			expect(client.room(id).destroyed).toBe(true);
			expect(client.chat.leavingChats()).toEqual([]);
			client.chat.destroy();
		});

		it('leaves a room that only shows them as listed once it syncs', () => {
			const doc = new Y.Doc();
			const id = createChatDirectory(doc, ann).createGroup({ others: [bob, cy] })!.id;
			const client = device(bob, doc);
			expect(client.chat.leaveChat(id)).toBe(true);
			client.room(id).doc.getMap<boolean>('teams:chat-members').set('bob', true);
			client.roomSyncs(id);
			expect(listed(client.room(id).doc, 'bob')).toBe(false);
			client.chat.destroy();
		});

		it('stops waiting for a room the server refuses once the workspace is back', () => {
			const { id, client } = leftOffline();
			client.directorySyncs();
			vi.advanceTimersByTime(29_000);
			expect(client.room(id).destroyed).toBe(false);
			vi.advanceTimersByTime(1000);
			expect(client.room(id).destroyed).toBe(true);
			client.chat.destroy();
		});

		it('takes up an undelivered leave again after a restart', () => {
			const { doc, id, client } = leftOffline();
			const kept = Y.encodeStateAsUpdate(client.room(id).doc);
			client.chat.destroy();

			const next = device(bob, doc, () => {
				const restored = new Y.Doc();
				Y.applyUpdate(restored, kept);
				return restored;
			});
			expect(next.chat.chat(id)).toBeNull();
			expect(next.chat.leavingChats()).toEqual([id]);
			next.directorySyncs();
			next.roomSyncs(id);
			vi.advanceTimersByTime(1500);
			expect(next.room(id).destroyed).toBe(true);
			next.chat.destroy();
		});

		it('opens nothing for a chat left long ago or lost to a removal', () => {
			const { doc, id, client } = leftOffline();
			client.chat.destroy();
			const empty = device(bob, doc, () => new Y.Doc());
			expect(empty.chat.leavingChats()).toEqual([]);
			empty.chat.destroy();

			const stillListed = device(bob, doc, () => {
				const restored = new Y.Doc();
				restored.getMap<boolean>('teams:chat-members').set('bob', true);
				return restored;
			});
			expect(stillListed.chat.leavingChats()).toEqual([]);
			expect(stillListed.chat.chat(id)).toBeNull();
			stillListed.chat.destroy();
		});

		it('carries on in the same room when added back before the leave went through', async () => {
			const { doc, id, client } = leftOffline();
			createChatDirectory(doc, ann).add(id, [bob]);
			await Promise.resolve();
			expect(client.chat.chat(id)?.id).toBe(id);
			expect(client.chat.leavingChats()).toEqual([]);
			expect(client.room(id).destroyed).toBe(false);
			expect(listed(client.room(id).doc, 'bob')).toBe(true);
			client.chat.destroy();
		});
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
