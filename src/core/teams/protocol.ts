// The wire contract of OpenTeams in one place: the protocol version a client announces to the
// server and the naming rules of every room. Both are pinned by tests, because a server, a
// stored snapshot and a client from another release must keep agreeing on them (see
// docs/teams-area.md, "Compatibility policy"). DOM-free. New code.
import { chatRoomId, directChatId } from './chats';
import { callRoomId, chatCallRoomId } from './view';

export { chatRoomId, directChatId, callRoomId, chatCallRoomId };

/**
 * Version of the sync and signaling protocol: the room names, the query parameters and the y-websocket
 * message framing. It changes only when an old client or server could misread the other.
 */
export const PROTOCOL_VERSION = 2;

/** Query parameter that carries {@link PROTOCOL_VERSION} on every sync and signaling connection. */
export const PROTOCOL_PARAM = 'proto';

/** The longest room name the reference server accepts. */
export const MAX_ROOM_ID_LENGTH = 128;

/** Whether `id` can name a workspace room: 1 to {@link MAX_ROOM_ID_LENGTH} of `[A-Za-z0-9_-]`. */
export const isWorkspaceId = (id: string): boolean =>
	id.length > 0 && id.length <= MAX_ROOM_ID_LENGTH && /^[\w-]+$/u.test(id);

/** `chat-<16 hex of the workspace>-<chat id>`: a private chat's own sync room. */
export const CHAT_ROOM_PATTERN = /^chat-[0-9a-f]{16}-[\w-]+$/u;
/** `dm-<32 hex>`: the chat id of a 1:1 chat (the same two people always get the same id). */
export const DIRECT_CHAT_ID_PATTERN = /^dm-[0-9a-f]{32}$/u;
/** `g-<32 hex>-<24 hex>`: a group chat id, random and bound to the person who created it. */
export const GROUP_CHAT_ID_PATTERN = /^g-[0-9a-f]{32}-[0-9a-f]{24}$/u;
/** Whether a chat id has a format only protocol 2 issues: no older room can rightly carry it. */
export const isCurrentChatId = (id: string): boolean =>
	DIRECT_CHAT_ID_PATTERN.test(id) || GROUP_CHAT_ID_PATTERN.test(id);
/** `call-<workspace>-<channel>`: the signaling room of a call. */
export const CALL_ROOM_PATTERN = /^call-[\w-]+$/u;
/** The end of a call room an earlier client used for a current chat id: `call-<workspace>-<chat id>`. */
const CHAT_ID_SUFFIX = /-(?:dm-[0-9a-f]{32}|g-[0-9a-f]{32}-[0-9a-f]{24})$/u;

export type RoomKind = 'chat' | 'call' | 'workspace';

export const isChatRoom = (room: string): boolean => CHAT_ROOM_PATTERN.test(room);

/** The chat id inside a chat room's name, or null when `room` is not a chat room. */
export const chatIdOfRoom = (room: string): string | null =>
	isChatRoom(room) ? room.slice('chat-'.length + 16 + 1) : null;

/**
 * A name that differs from a chat room only by letter case. It is never a room of its own: on a
 * case-insensitive disk it would otherwise be the private room it resembles.
 */
export const isChatRoomAlias = (room: string): boolean =>
	/^chat-[0-9a-f]{16}-/iu.test(room) && !isChatRoom(room);

/**
 * The chat room whose members may join the call room `room` (`call-chat-<16 hex>-<chat id>`, see
 * `chatCallRoomId`), or null when `room` is not a private chat's call.
 */
export const chatRoomOfCall = (room: string): string | null =>
	room.startsWith('call-') && isChatRoom(room.slice('call-'.length))
		? room.slice('call-'.length)
		: null;

/**
 * A call room no current client opens for a private chat: a letter-case variant of a chat's call
 * room, or the `call-<workspace>-<chat id>` form, which names a current chat id but carries no chat
 * room a server could check members against.
 */
export const isStrayChatCallRoom = (room: string): boolean =>
	room.startsWith('call-') &&
	chatRoomOfCall(room) === null &&
	(isChatRoomAlias(room.slice('call-'.length)) || CHAT_ID_SUFFIX.test(room));

/**
 * What a room name is. A workspace room is the workspace id itself (1-128 of `[\w-]`), so a
 * workspace whose id looks like a chat or call room is read as that kind: do not name workspaces
 * `chat-...` or `call-...`.
 */
export function classifyRoom(room: string): RoomKind | null {
	if (room.length === 0 || room.length > MAX_ROOM_ID_LENGTH) return null;
	if (isChatRoom(room)) return 'chat';
	if (isChatRoomAlias(room)) return null;
	if (CALL_ROOM_PATTERN.test(room)) return 'call';
	return /^[\w-]+$/u.test(room) ? 'workspace' : null;
}

/** Query parameters every sync and signaling connection sends. */
export const protocolParams = (): Record<string, string> => ({
	[PROTOCOL_PARAM]: String(PROTOCOL_VERSION),
});

/** The protocol version a connection URL announces, or null when it does not (an older client). */
export function protocolOfUrl(url: string): number | null {
	const raw = new URL(url, 'http://placeholder.invalid').searchParams.get(PROTOCOL_PARAM);
	return raw !== null && /^[1-9]\d{0,5}$/u.test(raw) ? Number(raw) : null;
}
