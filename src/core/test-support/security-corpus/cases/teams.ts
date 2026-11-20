import * as Y from 'yjs';

import { type ChatMember, createChatDirectory } from '../../../teams/chats';
import { type UpdateActor, validateTeamsUpdate } from '../../../teams/validate-update';
import type { SecurityCase } from '../case';

const person = (id: string): ChatMember => ({ id, name: id });
const actor = (id: string): UpdateActor => ({ id });

/** The update `by` sends after `change` edits a copy of the server's document, and its verdict. */
function verdict(server: Y.Doc, by: string, change: (doc: Y.Doc) => void): string[] {
	const doc = new Y.Doc();
	Y.applyUpdate(doc, Y.encodeStateAsUpdate(server));
	const updates: Uint8Array[] = [];
	doc.on('update', (update: Uint8Array) => void updates.push(update));
	change(doc);
	const result = validateTeamsUpdate(server, Y.mergeUpdates(updates), actor(by));
	return result.ok ? [] : result.violations.map((violation) => violation.code);
}

const LIMITS = { ms: 2_000, heapMb: 96 };

export const cases: readonly SecurityCase[] = [
	{
		name: 'a non-member cannot turn a private chat into a group that lists them',
		limits: LIMITS,
		run() {
			const server = new Y.Doc();
			server.clientID = 1;
			const directory = createChatDirectory(server, person('ann'));
			const direct = directory.ensureDirect(person('bob'))!.id;
			return verdict(server, 'eve', (doc) => {
				const row = doc.getMap<Y.Map<unknown>>('teams:chats').get(direct)!;
				doc.transact(() => {
					row.set('kind', 'group');
					row.set('members', [...(row.get('members') as ChatMember[]), person('eve')]);
				});
			});
		},
		expected: ['not-member'],
	},
	{
		name: 'nobody becomes a member by writing themselves into an empty room',
		limits: LIMITS,
		run() {
			return verdict(new Y.Doc(), 'eve', (doc) =>
				doc.getMap<boolean>('teams:chat-members').set('eve', true),
			);
		},
		expected: ['not-member'],
	},
	{
		name: 'a member cannot remove somebody else by rewriting their entry',
		limits: LIMITS,
		run() {
			const server = new Y.Doc();
			server.clientID = 1;
			const members = server.getMap<boolean>('teams:chat-members');
			members.set('ann', true);
			members.set('bob', true);
			return verdict(server, 'ann', (doc) =>
				doc.getMap<boolean>('teams:chat-members').set('bob', false),
			);
		},
		expected: ['malformed-row'],
	},
];
