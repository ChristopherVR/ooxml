// Private chats on top of the channel store. The workspace document holds the channels and the
// chat directory; every chat the person belongs to gets a Y.Doc of its own (its own sync room),
// so people who are not members never receive its messages: the reference server refuses them the
// room, and this model never opens, reads or writes a chat the person is not a member of. The
// returned object is a drop-in `ChatStore`, so everything built on channels (threads, pins,
// reactions, mentions, drafts, activity) works unchanged in a chat. DOM-free. New code.
import * as Y from 'yjs';
import type { ChatStore, ChatUser } from './chat';
import { createChatStore } from './chat';
import {
	type ChatDirectory,
	type ChatInfo,
	type ChatMember,
	chatTitle,
	createChatDirectory,
	isChatMember,
} from './chats';
import type { Channel, Message } from './model';
import type { PinnedMessage } from './pins';
import type { ReactionEvent } from './chat';

/** The members map inside a chat document; the reference server reads it to admit sockets. */
export const CHAT_MEMBERS_MAP = 'teams:chat-members';

export interface ChatRoomHandle {
	destroy: () => void;
}

export interface ChatRoomsOptions {
	/** The workspace document (holds the directory) and the channel store built on it. */
	doc: Y.Doc;
	base: ChatStore;
	user: ChatUser;
	/**
	 * Start syncing a chat's own document. Called once per chat the person belongs to, with a
	 * document the caller may seed from storage. Return a handle that stops the sync.
	 */
	openRoom: (chat: ChatInfo, doc: Y.Doc) => ChatRoomHandle;
	/** A document for a chat, for example restored from a local snapshot. */
	createDoc?: (chat: ChatInfo) => Y.Doc;
}

export interface ChatRooms extends ChatStore {
	/** The chats the person belongs to, in creation order. */
	chats: () => ChatInfo[];
	chat: (id: string) => ChatInfo | null;
	/** The 1:1 chat with someone, created when missing. */
	openDirect: (other: ChatMember) => ChatInfo | null;
	createGroup: (input: {
		others: readonly ChatMember[];
		name?: string;
		id?: string;
		meetingId?: string;
	}) => ChatInfo | null;
	renameChat: (id: string, name: string) => boolean;
	addPeople: (id: string, people: readonly ChatMember[]) => ChatMember[];
	removePerson: (id: string, userId: string) => boolean;
	leaveChat: (id: string) => boolean;
	/** Whether a message id is a chat the person belongs to. */
	isChat: (id: string) => boolean;
}

interface Open {
	doc: Y.Doc;
	store: ChatStore;
	handle: ChatRoomHandle;
	off: () => void;
}

const CLOSE_GRACE_MS = 1500;
const systemText = (verb: string, names: string[]): string => `${names.join(', ')} ${verb}`;

export function createChatRooms(options: ChatRoomsOptions): ChatRooms {
	const { doc, base, user, openRoom } = options;
	const directory: ChatDirectory = createChatDirectory(doc, user);
	const open = new Map<string, Open>();
	const listeners = new Set<() => void>();
	const closing = new Set<() => void>();
	const mineNow = (): ChatInfo[] => directory.list().filter((c) => isChatMember(c, user.id));
	const notify = (): void => {
		for (const l of [...listeners]) l();
	};

	/** Keep the chat document's member map equal to the directory: the server admits by it. */
	const mirrorMembers = (chat: ChatInfo, chatDoc: Y.Doc): void => {
		const map = chatDoc.getMap<boolean>(CHAT_MEMBERS_MAP);
		const want = new Set(chat.members.map((m) => m.id));
		const stale = [...map.keys()].filter((k) => !want.has(k));
		const missing = [...want].filter((k) => map.get(k) !== true);
		if (!stale.length && !missing.length) return;
		chatDoc.transact(() => {
			for (const k of stale) map.delete(k);
			for (const k of missing) map.set(k, true);
		});
	};

	const sync = (): void => {
		const wanted = new Map(mineNow().map((c) => [c.id, c]));
		for (const [id, entry] of open)
			if (!wanted.has(id)) {
				// Take the person off the member list the server admits by, then let a last notice flush.
				const info = directory.get(id);
				if (info) mirrorMembers(info, entry.doc);
				open.delete(id);
				entry.off();
				const close = (): void => {
					closing.delete(close);
					entry.handle.destroy();
					entry.store.destroy();
					entry.doc.destroy();
				};
				closing.add(close);
				setTimeout(close, CLOSE_GRACE_MS);
			}
		let changed = false;
		for (const [id, chat] of wanted) {
			const existing = open.get(id);
			if (existing) {
				mirrorMembers(chat, existing.doc);
				continue;
			}
			const chatDoc = options.createDoc?.(chat) ?? new Y.Doc();
			mirrorMembers(chat, chatDoc);
			const store = createChatStore(chatDoc, user);
			store.createChannel({ id, name: 'Chat', kind: 'direct' });
			const handle = openRoom(chat, chatDoc);
			open.set(id, { doc: chatDoc, store, handle, off: store.observe(notify) });
			changed = true;
		}
		if (changed) notify();
	};
	const stopBase = base.observe(() => {
		sync();
		notify();
	});
	sync();

	const storeOf = (id: string): ChatStore | undefined => open.get(id)?.store;
	const route = <T>(id: string, fallback: T, fn: (s: ChatStore) => T, other: () => T): T => {
		if (open.has(id)) return fn(open.get(id)!.store);
		if (directory.get(id)) return fallback; // a chat the person is not in
		return other();
	};
	const asChannel = (c: ChatInfo): Channel => ({
		id: c.id,
		name: chatTitle(c, user.id),
		topic: '',
		kind: 'direct',
		createdAt: c.createdAt,
		createdBy: c.createdBy,
		archived: false,
	});
	const notice = (id: string, text: string): void => {
		storeOf(id)?.post(id, { text, system: true });
	};
	const names = (people: readonly ChatMember[]): string[] => people.map((p) => p.name);

	const removePerson = (id: string, userId: string): boolean => {
		const chat = directory.get(id);
		const gone = chat?.members.find((m) => m.id === userId);
		if (!chat || !gone || !isChatMember(chat, user.id)) return false;
		// Say it while the chat is still open on this device.
		if (userId === user.id) notice(id, `${user.name} left the chat`);
		const ok = directory.remove(id, userId);
		if (ok && userId !== user.id) notice(id, `${user.name} removed ${gone.name} from the chat`);
		if (ok) sync();
		return ok;
	};

	return {
		chats: mineNow,
		chat: (id) => mineNow().find((c) => c.id === id) ?? null,
		isChat: (id) => open.has(id),
		openDirect(other) {
			const info = directory.ensureDirect(other);
			sync();
			return info;
		},
		createGroup(input) {
			const info = directory.createGroup(input);
			sync();
			return info;
		},
		renameChat(id, name) {
			const ok = directory.rename(id, name);
			if (ok) {
				const to = directory.get(id)?.name;
				notice(
					id,
					to ? `${user.name} named the chat "${to}"` : `${user.name} removed the chat name`,
				);
			}
			return ok;
		},
		addPeople(id, people) {
			const added = directory.add(id, people);
			if (added.length) {
				sync();
				notice(id, systemText('joined the chat', names(added)));
			}
			return added;
		},
		removePerson,
		leaveChat: (id) => removePerson(id, user.id),
		channels() {
			const real = base.channels().filter((c) => c.kind !== 'direct');
			return [...real, ...mineNow().map(asChannel)];
		},
		createChannel: (input) => base.createChannel(input),
		archiveChannel: (id, archived) => base.archiveChannel(id, archived),
		messages: (id) =>
			route<Message[]>(
				id,
				[],
				(s) => s.messages(id),
				() => base.messages(id),
			),
		post: (id, input) =>
			route(
				id,
				null,
				(s) => s.post(id, input),
				() => base.post(id, input),
			),
		ensureTabConversation: (tab) => base.ensureTabConversation(tab),
		edit: (id, mid, text, mentions) =>
			route(
				id,
				false,
				(s) => s.edit(id, mid, text, mentions),
				() => base.edit(id, mid, text, mentions),
			),
		remove: (id, mid) =>
			route(
				id,
				false,
				(s) => s.remove(id, mid),
				() => base.remove(id, mid),
			),
		pin: (id, mid) =>
			route(
				id,
				false,
				(s) => s.pin(id, mid),
				() => base.pin(id, mid),
			),
		unpin: (id, mid) =>
			route(
				id,
				false,
				(s) => s.unpin(id, mid),
				() => base.unpin(id, mid),
			),
		pinned: (id) =>
			route<PinnedMessage[]>(
				id,
				[],
				(s) => s.pinned(id),
				() => base.pinned(id),
			),
		reactionEvents(): ReactionEvent[] {
			return [
				...base.reactionEvents(),
				...[...open.values()].flatMap((o) => o.store.reactionEvents()),
			];
		},
		react(mid, emoji, channelId) {
			if (channelId && open.has(channelId)) open.get(channelId)!.store.react(mid, emoji);
			else base.react(mid, emoji);
		},
		unreact(mid, emoji, channelId) {
			if (channelId && open.has(channelId)) open.get(channelId)!.store.unreact(mid, emoji);
			else base.unreact(mid, emoji);
		},
		observe(listener) {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
		destroy() {
			stopBase();
			for (const close of [...closing]) close();
			for (const entry of open.values()) {
				entry.off();
				entry.handle.destroy();
				entry.store.destroy();
			}
			open.clear();
			listeners.clear();
			base.destroy();
		},
	};
}
