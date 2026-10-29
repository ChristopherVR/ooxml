// The wire contract of OpenTeams in one place: the protocol version a client announces to the
// server and the naming rules of every room. Both are pinned by tests, because a server, a
// stored snapshot and a client from another release must keep agreeing on them (see
// docs/teams-area.md, "Compatibility policy"). DOM-free. New code.
import { chatRoomId, directChatId } from './chats';
import { callRoomId } from './view';

export { chatRoomId, directChatId, callRoomId };

/**
 * Version of the sync and signaling protocol: the room names, the query parameters and the y-websocket
 * message framing. It changes only when an old client or server could misread the other.
 */
export const PROTOCOL_VERSION = 1;

/** Query parameter that carries {@link PROTOCOL_VERSION} on every sync and signaling connection. */
export const PROTOCOL_PARAM = 'proto';

/** The longest room name the reference server accepts. */
export const MAX_ROOM_ID_LENGTH = 128;

/** `chat-<16 hex of the workspace>-<chat id>`: a private chat's own sync room. */
export const CHAT_ROOM_PATTERN = /^chat-[0-9a-f]{16}-[\w-]+$/u;
/** `dm-<16 hex>`: the chat id of a 1:1 chat (the same two people always get the same id). */
export const DIRECT_CHAT_ID_PATTERN = /^dm-[0-9a-f]{16}$/u;
/** `call-<workspace>-<channel>`: the signaling room of a call. */
export const CALL_ROOM_PATTERN = /^call-[\w-]+$/u;

export type RoomKind = 'chat' | 'call' | 'workspace';

/**
 * What a room name is. A workspace room is the workspace id itself (1-128 of `[\w-]`), so a
 * workspace whose id looks like a chat or call room is read as that kind: do not name workspaces
 * `chat-...` or `call-...`.
 */
export function classifyRoom(room: string): RoomKind | null {
	if (room.length === 0 || room.length > MAX_ROOM_ID_LENGTH) return null;
	if (CHAT_ROOM_PATTERN.test(room)) return 'chat';
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
