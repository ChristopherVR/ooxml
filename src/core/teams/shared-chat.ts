// Assistant conversations shared over the collaboration engine. A shared chat is its own room, so
// anyone who has its id and reaches the same sync server (or, without one, another tab of the same
// browser) sees the same messages, whichever workspace or profile they are in. The person who asks
// runs the model with their own connection and streams the reply into the room; everyone else
// watches it arrive. Messages are rows keyed by id plus an order list, so two people asking at the
// same time both land. Everything a peer wrote is re-validated when read back. DOM-free. New code.
import * as Y from 'yjs';
import type { BroadcastChannelConstructor } from '../collab/broadcast-transport';
import { createClientId } from '../collab/identity';
import type { Transport } from '../collab/provider';
import { type CollabSession, createCollabSession } from '../collab/session';
import { transportProvider } from '../collab/transport-provider';
import { isValidId, isValidRoomId, sanitizeUserName } from '../collab/validation';
import type { WebSocketConstructor } from '../collab/websocket-transport';
import { type TeamsServerConfig, localServerConfig } from './server-config';
import { createSyncTransport } from './sync-transport';

export type SharedChatRole = 'user' | 'assistant';

export interface SharedChatMessage {
	id: string;
	role: SharedChatRole;
	content: string;
	authorId: string;
	authorName: string;
	ts: number;
	/** The asker's client is still writing this reply. */
	streaming: boolean;
	/** The host's own fields (attached file names, proposed actions, versions). JSON only; re-validate it. */
	data: Record<string, unknown>;
}

export type SharedChatInput = Omit<SharedChatMessage, 'authorId' | 'authorName' | 'ts'> &
	Partial<Pick<SharedChatMessage, 'authorId' | 'authorName' | 'ts'>>;

/** What a person in the chat is doing right now. */
export type SharedChatActivity = 'idle' | 'typing' | 'asking';

export interface SharedChatPerson {
	clientId: number;
	name: string;
	color: string;
	activity: SharedChatActivity;
}

export const MAX_SHARED_CHAT_CHARS = 200_000;
export const MAX_SHARED_CHAT_DATA_CHARS = 1_000_000;
export const MAX_SHARED_CHAT_TITLE_CHARS = 120;
const MESSAGES = 'chat:messages';
const ORDER = 'chat:order';
const META = 'chat:meta';

/** A new, unguessable room id: whoever holds it can join, so it carries 128 random bits. */
export function sharedChatRoomId(): string {
	const bytes = crypto.getRandomValues(new Uint8Array(16));
	return `aichat-${Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('')}`;
}

/** True for a room id made by {@link sharedChatRoomId}. */
export const isSharedChatRoomId = (value: unknown): value is string =>
	isValidRoomId(value) && /^aichat-[0-9a-f]{32}$/u.test(value);

function plainData(raw: unknown): Record<string, unknown> {
	if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return {};
	try {
		const text = JSON.stringify(raw);
		if (text.length > MAX_SHARED_CHAT_DATA_CHARS) return {};
		return JSON.parse(text) as Record<string, unknown>;
	} catch {
		return {};
	}
}

function readMessage(id: string, raw: unknown): SharedChatMessage | null {
	if (!raw || typeof raw !== 'object' || !isValidId(id)) return null;
	const r = raw as Record<string, unknown>;
	if (r.role !== 'user' && r.role !== 'assistant') return null;
	return {
		id,
		role: r.role,
		content: typeof r.content === 'string' ? r.content.slice(0, MAX_SHARED_CHAT_CHARS) : '',
		authorId: typeof r.authorId === 'string' ? r.authorId.slice(0, 160) : '',
		authorName: sanitizeUserName(r.authorName),
		ts: typeof r.ts === 'number' && Number.isFinite(r.ts) ? r.ts : 0,
		streaming: r.streaming === true,
		data: plainData(r.data),
	};
}

export interface SharedChatStore {
	title: () => string;
	setTitle: (title: string) => void;
	/** The messages in order, each re-validated. */
	messages: () => SharedChatMessage[];
	/** Adds a message, or replaces the one with its id. Unchanged messages are not written again. */
	put: (message: SharedChatInput) => void;
	/** Removes a message and everything after it, as editing and resending an earlier request does. */
	truncateFrom: (id: string) => void;
	/** Fires after any change, local or remote. */
	observe: (listener: () => void) => () => void;
}

export function createSharedChatStore(
	doc: Y.Doc,
	user: { id: string; name: string },
): SharedChatStore {
	const rows = doc.getMap<unknown>(MESSAGES);
	const order = doc.getArray<string>(ORDER);
	const meta = doc.getMap<unknown>(META);
	const ids = (): string[] => {
		// Two people adding the same id at once leave it twice in the list; show it once.
		const seen = new Set<string>();
		return order.toArray().filter((id) => typeof id === 'string' && !seen.has(id) && seen.add(id));
	};
	return {
		title: () => {
			const raw = meta.get('title');
			return typeof raw === 'string' ? raw.slice(0, MAX_SHARED_CHAT_TITLE_CHARS) : '';
		},
		setTitle(title) {
			const next = title.replace(/\s+/gu, ' ').trim().slice(0, MAX_SHARED_CHAT_TITLE_CHARS);
			if (next && next !== meta.get('title')) meta.set('title', next);
		},
		messages: () =>
			ids()
				.map((id) => readMessage(id, rows.get(id)))
				.filter((m): m is SharedChatMessage => m !== null),
		put(message) {
			if (!isValidId(message.id)) throw new Error('A shared message needs an id.');
			const previous = readMessage(message.id, rows.get(message.id));
			const next = readMessage(message.id, {
				authorId: previous?.authorId ?? user.id,
				authorName: previous?.authorName ?? user.name,
				ts: previous?.ts ?? Date.now(),
				...message,
			});
			if (!next) throw new Error('A shared message needs a role.');
			if (previous && JSON.stringify(previous) === JSON.stringify(next)) return;
			doc.transact(() => {
				const { id, ...row } = next;
				rows.set(id, row);
				if (!previous) order.push([id]);
			});
		},
		truncateFrom(id) {
			const list = order.toArray();
			const at = list.indexOf(id);
			if (at < 0) return;
			doc.transact(() => {
				for (const gone of list.slice(at)) rows.delete(gone);
				order.delete(at, list.length - at);
			});
		},
		observe(listener) {
			let queued = false;
			const handler = (): void => {
				if (queued) return;
				queued = true;
				queueMicrotask(() => {
					queued = false;
					listener();
				});
			};
			doc.on('update', handler);
			return () => doc.off('update', handler);
		},
	};
}

export interface SharedChatOptions {
	roomId: string;
	user: { id?: string; name: string; color?: string };
	config?: TeamsServerConfig;
	/** Use this transport instead of the one the config selects (tests, custom hosts). */
	transport?: Transport;
	doc?: Y.Doc;
	WebSocket?: WebSocketConstructor;
	BroadcastChannel?: BroadcastChannelConstructor;
}

export interface SharedChat {
	readonly roomId: string;
	readonly user: { id: string; name: string };
	readonly session: CollabSession<{ activity?: SharedChatActivity }>;
	readonly store: SharedChatStore;
	/** The other people in the chat now. */
	people: () => SharedChatPerson[];
	onPeople: (listener: (people: SharedChatPerson[]) => void) => () => void;
	setActivity: (activity: SharedChatActivity) => void;
	destroy: () => void;
}

const ACTIVITIES: readonly SharedChatActivity[] = ['idle', 'typing', 'asking'];

export function createSharedChat(options: SharedChatOptions): SharedChat {
	if (!isValidRoomId(options.roomId)) throw new Error('That is not a shared chat link.');
	const config = options.config ?? localServerConfig();
	const user = {
		id: options.user.id ?? createClientId(),
		name: sanitizeUserName(options.user.name),
	};
	const transport =
		options.transport ??
		createSyncTransport(config, options.roomId, {
			WebSocket: options.WebSocket,
			BroadcastChannel: options.BroadcastChannel,
		});
	const session = createCollabSession<{ activity?: SharedChatActivity }>({
		roomId: options.roomId,
		...(options.doc ? { doc: options.doc } : {}),
		provider: transportProvider({ transport }),
		user: { name: user.name, ...(options.user.color ? { color: options.user.color } : {}) },
		initialPresence: { activity: 'idle' },
		sanitizePayload: (raw) => ({
			activity: ACTIVITIES.includes(raw.activity as SharedChatActivity)
				? (raw.activity as SharedChatActivity)
				: 'idle',
		}),
		teardown: true,
	});
	const people = (): SharedChatPerson[] =>
		session.peers().map((p) => ({
			clientId: p.clientId,
			name: p.userName,
			color: p.userColor,
			activity: p.activity ?? 'idle',
		}));
	return {
		roomId: options.roomId,
		user,
		session,
		store: createSharedChatStore(session.doc, user),
		people,
		onPeople: (listener) => session.on('peers', () => listener(people())),
		setActivity: (activity) => session.updatePresence({ activity }),
		destroy: () => session.destroy(),
	};
}
