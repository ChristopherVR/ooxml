// The Chat app's actions for the client: start a 1:1 chat, create a group, rename it, add and remove
// people, leave. People are named by user id; the roster supplies the display names, so a UI only
// ever passes ids. Kept out of `store.ts` so that file stays about channels and calls. DOM-free.
import type { ChatMember } from './chats';
import type { TeamsWorkspace } from './workspace';

export interface ChatActions {
	/** Open (creating if needed) the 1:1 chat with a person and select it. Returns the chat id. */
	openDirectChat: (userId: string) => string | null;
	/** A group of you and the people listed (at least one), optionally named. Selects it. */
	createGroupChat: (userIds: readonly string[], name?: string) => string | null;
	renameChat: (chatId: string, name: string) => boolean;
	addToChat: (chatId: string, userIds: readonly string[]) => boolean;
	removeFromChat: (chatId: string, userId: string) => boolean;
	/** Leave a group chat. 1:1 chats cannot be left. */
	leaveChat: (chatId: string) => boolean;
}

export interface ChatActionDeps {
	ws: TeamsWorkspace;
	select: (id: string) => void;
	notice: (text: string) => void;
	refresh: () => void;
}

export function createChatActions({ ws, select, notice, refresh }: ChatActionDeps): ChatActions {
	const people = (ids: readonly string[]): ChatMember[] => {
		const known = new Map(ws.roster.list().map((r) => [r.id, r]));
		const out: ChatMember[] = [];
		for (const id of ids) {
			const found = known.get(id);
			if (found && id !== ws.user.id) out.push({ id: found.id, name: found.name });
		}
		return out;
	};
	const done = (ok: boolean, failure: string): boolean => {
		if (!ok) notice(failure);
		refresh();
		return ok;
	};
	return {
		openDirectChat(userId) {
			const other = people([userId])[0];
			const chat = other ? ws.chat.openDirect(other) : null;
			if (!chat) {
				notice('That person is not available to chat with');
				return null;
			}
			select(chat.id);
			return chat.id;
		},
		createGroupChat(userIds, name = '') {
			const others = people(userIds);
			if (others.length === 0) {
				notice('Pick at least one person');
				return null;
			}
			const chat =
				others.length === 1 && !name.trim()
					? ws.chat.openDirect(others[0]!)
					: ws.chat.createGroup({ others, name });
			if (!chat) {
				notice('The chat could not be created');
				return null;
			}
			select(chat.id);
			return chat.id;
		},
		renameChat: (chatId, name) =>
			done(ws.chat.renameChat(chatId, name), 'Only group chats can be renamed'),
		addToChat: (chatId, userIds) =>
			done(ws.chat.addPeople(chatId, people(userIds)).length > 0, 'Nobody was added to the chat'),
		removeFromChat: (chatId, userId) =>
			done(ws.chat.removePerson(chatId, userId), 'That person cannot be removed'),
		leaveChat: (chatId) => done(ws.chat.leaveChat(chatId), 'You cannot leave this chat'),
	};
}
