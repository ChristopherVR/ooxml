// Private chats on top of the channel store. The workspace document holds the channels and the
// chat directory; every chat the person belongs to gets a Y.Doc of its own (its own sync room),
// so people who are not members never receive its messages: the reference server refuses them the
// room, and this model never opens, reads or writes a chat the person is not a member of. The
// room's member list changes only through this person's own actions (creating a group, adding
// people, leaving), never by copying what somebody else wrote in the directory. One repair: the
// creator of a group admits its listed people again when the room opens, in case a crash came
// between creating the group and admitting them. Leaving is written to the room like any other
// change, so a room left while offline stays open until it has synced, and its local copy outlives
// a restart for the same reason. The returned object is a drop-in `ChatStore`, so everything built
// on channels (threads, pins, reactions, mentions, drafts, activity) works unchanged in a chat.
// DOM-free. New code.
import * as Y from 'yjs';
import type { ChatStore, ChatUser } from './chat';
import { createChatStore } from './chat';
import { CHATS_MAP, memberStanding } from './chat-records';
import {
	type ChatDirectory,
	type ChatInfo,
	type ChatMember,
	chatTitle,
	createChatDirectory,
	isChatMember,
	isGroupChatIdOf,
} from './chats';
import type { Channel, Message } from './model';
import type { PinnedMessage } from './pins';
import type { ReactionEvent } from './chat';

/** The members map inside a chat document; the reference server reads it to admit sockets. */
export const CHAT_MEMBERS_MAP = 'teams:chat-members';

export interface ChatRoomHandle {
	/** Whether the room has caught up with the server over its current connection. */
	synced: () => boolean;
	destroy: () => void;
}

export interface ChatRoomsOptions {
	/** The workspace document (holds the directory) and the channel store built on it. */
	doc: Y.Doc;
	base: ChatStore;
	user: ChatUser;
	/**
	 * Start syncing a chat's own document. Called once per chat the person belongs to, with a
	 * document the caller may seed from storage. Call `onSynced` when the document has caught up
	 * with the server. Return a handle that stops the sync.
	 */
	openRoom: (chat: ChatInfo, doc: Y.Doc, onSynced: () => void) => ChatRoomHandle;
	/** Whether the workspace document has caught up with the server, and news of it doing so. */
	directorySync?: { synced: () => boolean; onSynced: (listener: () => void) => () => void };
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
		meetingId?: string;
	}) => ChatInfo | null;
	renameChat: (id: string, name: string) => boolean;
	addPeople: (id: string, people: readonly ChatMember[]) => ChatMember[];
	removePerson: (id: string, userId: string) => boolean;
	leaveChat: (id: string) => boolean;
	/** Chats the person left whose room may not know yet. Their local copy must be kept. */
	leavingChats: () => string[];
	/** Whether a message id is a chat the person belongs to. */
	isChat: (id: string) => boolean;
}

interface Open {
	doc: Y.Doc;
	store: ChatStore;
	handle: ChatRoomHandle;
	off: () => void;
}

/** A room this person left, kept open until it has what they wrote there. */
interface Leaving {
	room: Open;
	/** The room has synced since they left: only the grace period remains. */
	delivered: boolean;
	timer: ReturnType<typeof setTimeout> | undefined;
}

const CLOSE_GRACE_MS = 1500;
/**
 * How long a left room may stay out of reach while the workspace is in sync. Both use one server,
 * so by then the server is refusing this room, and the person is out of it already.
 */
const LEAVE_TIMEOUT_MS = 30_000;
const systemText = (verb: string, names: string[]): string => `${names.join(', ')} ${verb}`;

export function createChatRooms(options: ChatRoomsOptions): ChatRooms {
	const { doc, base, user, openRoom } = options;
	const directory: ChatDirectory = createChatDirectory(doc, user);
	const open = new Map<string, Open>();
	const listeners = new Set<() => void>();
	const closing = new Set<() => void>();
	const leaving = new Map<string, Leaving>();
	const mineNow = (): ChatInfo[] => directory.list().filter((c) => isChatMember(c, user.id));
	const notify = (): void => {
		for (const l of [...listeners]) l();
	};

	const membersOf = (id: string): Y.Map<boolean> | undefined =>
		open.get(id)?.doc.getMap<boolean>(CHAT_MEMBERS_MAP);

	/** This person lets `people` into the chat's room: the list the server admits sockets by. */
	const admit = (id: string, people: readonly ChatMember[]): void => {
		const map = membersOf(id);
		const missing = people.map((p) => p.id).filter((k) => map?.get(k) !== true);
		if (!map || !missing.length) return;
		open.get(id)?.doc.transact(() => {
			for (const k of missing) map.set(k, true);
		});
	};

	/** Group rooms this person created and has not re-admitted people to since opening them. */
	const toReadmit = new Set<string>();
	const syncedRooms = new Set<string>();
	/** Waits until both documents are current, so nobody who has left is let back in. */
	const readmit = (id: string): void => {
		if (!toReadmit.has(id) || !syncedRooms.has(id) || !options.directorySync?.synced()) return;
		toReadmit.delete(id);
		const chat = directory.get(id);
		if (chat && isChatMember(chat, user.id)) admit(id, chat.members);
	};
	const directoryCurrent = (): boolean => options.directorySync?.synced() ?? true;
	const release = (room: Open): void => {
		room.handle.destroy();
		room.store.destroy();
		room.doc.destroy();
	};
	const finishLeave = (id: string): void => {
		const entry = leaving.get(id);
		if (!entry) return;
		leaving.delete(id);
		release(entry.room);
	};
	/** Sets the timer a left room now waits on. Offline there is none: the next sync comes back here. */
	const settleLeave = (id: string): void => {
		const entry = leaving.get(id);
		if (!entry || entry.delivered) return;
		clearTimeout(entry.timer);
		entry.timer = undefined;
		if (entry.room.handle.synced()) {
			// A room that had not synced before they left only now shows them as listed.
			entry.room.doc.getMap<boolean>(CHAT_MEMBERS_MAP).delete(user.id);
			entry.delivered = true;
			entry.timer = setTimeout(() => finishLeave(id), CLOSE_GRACE_MS);
		} else if (directoryCurrent()) {
			entry.timer = setTimeout(() => {
				if (directoryCurrent()) finishLeave(id);
			}, LEAVE_TIMEOUT_MS);
		}
	};
	const startLeave = (id: string, room: Open): void => {
		leaving.set(id, { room, delivered: false, timer: undefined });
		settleLeave(id);
	};
	/** Added back before the leave was finished: the room carries on, with this person in it. */
	const rejoin = (id: string, { room, timer }: Leaving): void => {
		leaving.delete(id);
		clearTimeout(timer);
		const members = room.doc.getMap<boolean>(CHAT_MEMBERS_MAP);
		if (members.get(user.id) !== true) members.set(user.id, true);
		open.set(id, { ...room, off: room.store.observe(notify) });
	};
	const stopDirectorySync = options.directorySync?.onSynced(() => {
		for (const id of [...toReadmit]) readmit(id);
		for (const id of [...leaving.keys()]) settleLeave(id);
	});

	const openChat = (chat: ChatInfo, chatDoc: Y.Doc): Open => {
		const { id } = chat;
		const store = createChatStore(chatDoc, user);
		store.createChannel({ id, name: 'Chat', kind: 'direct' });
		const handle = openRoom(chat, chatDoc, () => {
			syncedRooms.add(id);
			readmit(id);
			settleLeave(id);
		});
		return { doc: chatDoc, store, handle, off: store.observe(notify) };
	};

	/** Opens and closes rooms to match the directory. `left` is a chat this person just left. */
	const sync = (left?: string): void => {
		const wanted = new Map(mineNow().map((c) => [c.id, c]));
		for (const [id, entry] of open)
			if (!wanted.has(id)) {
				open.delete(id);
				toReadmit.delete(id);
				syncedRooms.delete(id);
				entry.off();
				if (id === left) {
					startLeave(id, entry);
					continue;
				}
				// Let a last notice and member list change flush before the room closes.
				const close = (): void => {
					closing.delete(close);
					release(entry);
				};
				closing.add(close);
				setTimeout(close, CLOSE_GRACE_MS);
			}
		let changed = false;
		for (const [id, chat] of wanted) {
			if (open.has(id)) continue;
			changed = true;
			const pending = leaving.get(id);
			if (pending) {
				rejoin(id, pending);
				continue;
			}
			if (chat.kind === 'group' && isGroupChatIdOf(id, user.id)) toReadmit.add(id);
			open.set(id, openChat(chat, options.createDoc?.(chat) ?? new Y.Doc()));
			readmit(id);
		}
		if (changed) notify();
	};

	/**
	 * Picks up leaves an earlier session could not deliver. The local copy of such a room was kept,
	 * and no longer lists this person. Any other chat they are out of has nothing left to send.
	 */
	const resumeLeaves = (): void => {
		const { createDoc } = options;
		if (!createDoc) return;
		doc.getMap<unknown>(CHATS_MAP).forEach((row, id) => {
			if (!(row instanceof Y.Map) || memberStanding(row, user.id) !== 'former') return;
			const chat = directory.get(id);
			if (!chat || open.has(id)) return;
			const chatDoc = createDoc(chat);
			const listed = chatDoc.getMap<boolean>(CHAT_MEMBERS_MAP).get(user.id) === true;
			if (listed || chatDoc.store.clients.size === 0) return chatDoc.destroy();
			const room = openChat(chat, chatDoc);
			room.off();
			startLeave(id, room);
		});
	};

	const stopBase = base.observe(() => {
		sync();
		notify();
	});
	sync();
	resumeLeaves();

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
		// The server lets a member remove only themselves; removing others is a moderator action.
		if (userId !== user.id && user.moderator !== true) return false;
		// Say it while the chat is still open on this device.
		if (userId === user.id) notice(id, `${user.name} left the chat`);
		// Held from here: the directory change below can close this person's room right away.
		const members = membersOf(id);
		const ok = directory.remove(id, userId);
		if (ok && userId !== user.id) notice(id, `${user.name} removed ${gone.name} from the chat`);
		if (ok) {
			members?.delete(userId);
			sync(userId === user.id ? id : undefined);
		}
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
			if (info) admit(info.id, info.members);
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
				admit(id, added);
				notice(id, systemText('joined the chat', names(added)));
			}
			return added;
		},
		removePerson,
		leaveChat: (id) => removePerson(id, user.id),
		leavingChats: () => [...leaving.keys()],
		channels() {
			const real = base.channels().filter((c) => c.kind !== 'direct');
			return [...real, ...mineNow().map(asChannel)];
		},
		createChannel: (input) => base.createChannel(input),
		archiveChannel: (id, archived) => base.archiveChannel(id, archived),
		renameChannel: (id, name) => base.renameChannel(id, name),
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
			stopDirectorySync?.();
			for (const close of [...closing]) close();
			for (const { room, timer } of leaving.values()) {
				clearTimeout(timer);
				release(room);
			}
			leaving.clear();
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
