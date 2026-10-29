import * as Y from 'yjs';
import { createChatDirectory, chatRoomId, chatTitle, directChatId, sanitizeMembers } from './chats';
import { createTeamsClient, type TeamsClient } from './store';

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));
const until = async (fn: () => boolean, ms = 3000): Promise<boolean> => {
	const end = Date.now() + ms;
	while (Date.now() < end) {
		if (fn()) return true;
		await sleep(10);
	}
	return fn();
};
const clients: TeamsClient[] = [];
let room = 0;
const workspace = (): string => `chats-${Date.now()}-${room++}`;
const make = (name: string, workspaceId: string, moderator = false): TeamsClient => {
	const c = createTeamsClient({
		workspaceId,
		user: { id: name, name: name[0]!.toUpperCase() + name.slice(1) },
		moderator,
		reminderCheckMs: 0,
	});
	clients.push(c);
	return c;
};
afterEach(() => {
	for (const c of clients.splice(0)) c.destroy();
});

describe('chat model', () => {
	it('names a 1:1 chat the same for both people, whoever opens it', () => {
		expect(directChatId('ada', 'bob')).toBe(directChatId('bob', 'ada'));
		expect(directChatId('ada', 'bob')).not.toBe(directChatId('ada', 'cy'));
		expect(chatRoomId('acme', 'dm-1').length).toBeLessThanOrEqual(128);
		expect(chatRoomId('x'.repeat(200), 'dm-1').length).toBeLessThanOrEqual(128);
	});

	it('titles chats by the other people', () => {
		const members = sanitizeMembers([
			{ id: 'ada', name: 'Ada' },
			{ id: 'bob', name: 'Bob' },
			{ id: 'cy', name: 'Cy' },
		]);
		expect(chatTitle({ kind: 'direct', name: '', members: members.slice(0, 2) }, 'ada')).toBe(
			'Bob',
		);
		expect(chatTitle({ kind: 'group', name: '', members }, 'ada')).toBe('Bob, Cy');
		expect(chatTitle({ kind: 'group', name: 'Launch', members }, 'ada')).toBe('Launch');
	});

	it('drops members with unsafe ids, duplicates and malformed rows', () => {
		expect(
			sanitizeMembers([
				{ id: 'ada', name: 'Ada' },
				{ id: 'ada', name: 'Again' },
				{ id: 'a,b', name: 'Comma' },
				{ id: 'x y', name: 'Space' },
				'nope',
				null,
			]),
		).toEqual([{ id: 'ada', name: 'Ada' }]);
	});

	it('lets only members rename, add and remove', () => {
		const doc = new Y.Doc();
		const ada = createChatDirectory(doc, { id: 'ada', name: 'Ada' });
		const eve = createChatDirectory(doc, { id: 'eve', name: 'Eve' });
		const group = ada.createGroup({
			others: [
				{ id: 'bob', name: 'Bob' },
				{ id: 'cy', name: 'Cy' },
			],
			name: 'Team',
		})!;
		expect(eve.rename(group.id, 'Mine now')).toBe(false);
		expect(eve.add(group.id, [{ id: 'eve', name: 'Eve' }])).toEqual([]);
		expect(eve.remove(group.id, 'bob')).toBe(false);
		expect(ada.rename(group.id, 'Renamed')).toBe(true);
		expect(ada.add(group.id, [{ id: 'dan', name: 'Dan' }]).map((m) => m.id)).toEqual(['dan']);
		expect(ada.remove(group.id, 'dan')).toBe(true);
		expect(ada.get(group.id)!.members.map((m) => m.id)).toEqual(['ada', 'bob', 'cy']);
	});

	it('does not rename or grow a 1:1 chat and ignores yourself', () => {
		const ada = createChatDirectory(new Y.Doc(), { id: 'ada', name: 'Ada' });
		expect(ada.ensureDirect({ id: 'ada', name: 'Ada' })).toBeNull();
		const dm = ada.ensureDirect({ id: 'bob', name: 'Bob' })!;
		expect(ada.ensureDirect({ id: 'bob', name: 'Bob' })!.id).toBe(dm.id);
		expect(ada.rename(dm.id, 'x')).toBe(false);
		expect(ada.add(dm.id, [{ id: 'cy', name: 'Cy' }])).toEqual([]);
		expect(ada.remove(dm.id, 'bob')).toBe(false);
	});
});

describe('chats in the client', () => {
	it('opens a 1:1 chat that only its two people can see or read', async () => {
		const ws = workspace();
		const ada = make('ada', ws);
		const bob = make('bob', ws);
		const cy = make('cy', ws);
		expect(await until(() => ada.getState().contacts.length === 2)).toBe(true);
		const id = ada.openDirectChat('bob')!;
		expect(id).toBeTruthy();
		await ada.send({ text: 'secret plans' });
		expect(await until(() => bob.getState().chats.some((c) => c.id === id))).toBe(true);
		bob.select(id);
		expect(await until(() => bob.getState().messages.some((m) => m.text === 'secret plans'))).toBe(
			true,
		);
		expect(bob.getState().chats[0]!.title).toBe('Ada');
		// Cy shares the workspace and can see the directory row, but never the chat or its messages.
		await sleep(100);
		expect(cy.getState().chats).toEqual([]);
		expect(cy.workspace.chat.messages(id)).toEqual([]);
		expect(cy.workspace.chat.post(id, { text: 'let me in' })).toBeNull();
		expect(cy.workspace.chat.react('anything', '👍', id)).toBeUndefined();
		cy.select(id);
		await sleep(20);
		expect(cy.getState().messages).toEqual([]);
		expect(cy.getState().selectedChannelId).not.toBe(id);
	});

	it('creates a group chat with a name, system messages, rename, add, remove and leave', async () => {
		const ws = workspace();
		const ada = make('ada', ws, true);
		const bob = make('bob', ws);
		const cy = make('cy', ws);
		const dan = make('dan', ws);
		expect(await until(() => ada.getState().contacts.length === 3)).toBe(true);
		const id = ada.createGroupChat(['bob', 'cy'], 'Launch')!;
		expect(await until(() => cy.getState().chats.some((c) => c.id === id))).toBe(true);
		expect(cy.getState().chats[0]!.title).toBe('Launch');
		expect(await until(() => dan.getState().chats.length === 0)).toBe(true);

		ada.renameChat(id, 'Launch team');
		expect(await until(() => bob.getState().chats[0]?.title === 'Launch team')).toBe(true);
		ada.addToChat(id, ['dan']);
		expect(await until(() => dan.getState().chats.some((c) => c.id === id))).toBe(true);
		dan.select(id);
		await until(() => dan.getState().messages.some((m) => m.text.includes('joined the chat')));
		const notice = dan.getState().messages.find((m) => m.text.includes('joined the chat'))!;
		expect(notice.system).toBe(true);
		expect(notice.text).toBe('Dan joined the chat');

		bob.leaveChat(id);
		expect(await until(() => bob.getState().chats.length === 0)).toBe(true);
		expect(await until(() => ada.getState().chats[0]?.members.length === 3)).toBe(true);
		ada.select(id);
		await until(() => ada.getState().messages.some((m) => m.text === 'Bob left the chat'));
		expect(ada.getState().messages.some((m) => m.text === 'Bob left the chat')).toBe(true);

		// Only a moderator removes somebody else: the server refuses it from a plain member.
		cy.removeFromChat(id, 'dan');
		await sleep(100);
		expect(dan.getState().chats.some((c) => c.id === id)).toBe(true);
		ada.removeFromChat(id, 'dan');
		expect(await until(() => dan.getState().chats.length === 0)).toBe(true);
		expect(dan.workspace.chat.messages(id)).toEqual([]);
	});

	it('lists chats recent first with unread counts and a last message preview', async () => {
		const ws = workspace();
		const ada = make('ada', ws);
		const bob = make('bob', ws);
		const cy = make('cy', ws);
		expect(await until(() => ada.getState().contacts.length === 2)).toBe(true);
		const withBob = ada.openDirectChat('bob')!;
		await ada.send({ text: 'first to bob' });
		const withCy = ada.openDirectChat('cy')!;
		await ada.send({ text: 'hello cy' });
		expect(ada.getState().chats.map((c) => c.id)).toEqual([withCy, withBob]);
		expect(ada.getState().chats[0]!.lastMessage).toMatchObject({ text: 'hello cy', mine: true });
		expect(await until(() => bob.getState().chats.some((c) => c.unread === 1))).toBe(true);
		bob.select(withBob);
		await sleep(10);
		expect(bob.getState().chats[0]!.unread).toBe(0);
		expect(await until(() => cy.getState().chats[0]?.lastMessage?.text === 'hello cy')).toBe(true);
		expect(cy.getState().chats[0]!.members.find((m) => m.id === 'ada')?.availability).toBeDefined();
	});

	it('keeps reactions, pins and mentions working inside a chat', async () => {
		const ws = workspace();
		const ada = make('ada', ws);
		const bob = make('bob', ws);
		expect(await until(() => ada.getState().contacts.length === 1)).toBe(true);
		const id = ada.openDirectChat('bob')!;
		await ada.send({ text: 'ping @Bob', mentions: ada.getState().mentionTargets });
		const message = ada.getState().messages[0]!;
		expect(message.mentions?.map((m) => m.name)).toEqual(['Bob']);
		ada.toggleReaction(message.id, '👍');
		ada.pinMessage(message.id);
		expect(await until(() => ada.getState().pinned.length === 1)).toBe(true);
		expect(await until(() => bob.getState().chats.some((c) => c.id === id))).toBe(true);
		bob.select(id);
		expect(await until(() => bob.getState().messages[0]?.reactions['👍']?.length === 1)).toBe(true);
		expect(bob.getState().activity.some((a) => a.kind === 'mention' && a.chat === true)).toBe(true);
		expect(bob.getState().mentionTargets.map((t) => t.name)).toEqual(['Ada']);
	});
});
