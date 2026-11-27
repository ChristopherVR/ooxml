import { describe, expect, it } from 'vitest';
import {
	CALL_ROOM_PATTERN,
	CHAT_ROOM_PATTERN,
	DIRECT_CHAT_ID_PATTERN,
	GROUP_CHAT_ID_PATTERN,
	PROTOCOL_PARAM,
	PROTOCOL_VERSION,
	callRoomId,
	chatCallRoomId,
	chatRoomId,
	chatRoomOfCall,
	classifyRoom,
	directChatId,
	chatIdOfRoom,
	isChatRoomAlias,
	isStrayChatCallRoom,
	isWorkspaceId,
	MAX_ROOM_ID_LENGTH,
	protocolOfUrl,
	protocolParams,
} from './protocol';
import { createGroupChatId, isGroupChatIdOf, legacyDirectChatId } from './chats';
import { endpointUrl } from './server-config';

describe('protocol version and room names', () => {
	it('accepts workspace ids up to the server room limit, as the server does', () => {
		expect(MAX_ROOM_ID_LENGTH).toBe(128);
		expect(isWorkspaceId('team_1-a')).toBe(true);
		expect(isWorkspaceId('w'.repeat(128))).toBe(true);
		expect(isWorkspaceId('w'.repeat(101))).toBe(true);
		expect(isWorkspaceId('w'.repeat(129))).toBe(false);
		expect(isWorkspaceId('')).toBe(false);
		expect(isWorkspaceId('a b')).toBe(false);
		expect(isWorkspaceId('a/b')).toBe(false);
		expect(classifyRoom('w'.repeat(128))).toBe('workspace');
	});

	it('pins the protocol version and its query parameter', () => {
		expect(PROTOCOL_VERSION).toBe(2);
		expect(PROTOCOL_PARAM).toBe('proto');
		expect(protocolParams()).toEqual({ proto: '2' });
	});

	it('pins the chat room format', () => {
		const room = chatRoomId('acme', 'g1');
		expect(room).toMatch(CHAT_ROOM_PATTERN);
		expect(room).toMatch(/^chat-[0-9a-f]{16}-g1$/u);
		expect(chatRoomId('acme', 'g1')).toBe(room);
		expect(chatRoomId('other', 'g1')).not.toBe(room);
		expect(chatRoomId('x'.repeat(300), 'g1').length).toBeLessThanOrEqual(128);
	});

	it('pins the hash behind the names (a change here renames every existing room)', () => {
		// Fixed vectors: if these change, every stored chat room is orphaned.
		expect(directChatId('ann', 'bob')).toBe('dm-c56f0deee04ba579e3730ecc643f5676');
		expect(directChatId('bob', 'ann')).toBe('dm-c56f0deee04ba579e3730ecc643f5676');
		expect(directChatId('ann', 'bob')).toMatch(DIRECT_CHAT_ID_PATTERN);
		expect(legacyDirectChatId('bob', 'ann')).toBe('dm-d50d61311f9ec6df');
		expect(chatRoomId('acme', 'g1')).toBe('chat-45fd71af8d5cbb5d-g1');
	});

	it('makes group ids random and bound to their creator', () => {
		const id = createGroupChatId('ann');
		expect(id).toMatch(GROUP_CHAT_ID_PATTERN);
		expect(id.length).toBeLessThanOrEqual(64);
		expect(createGroupChatId('ann')).not.toBe(id);
		expect(isGroupChatIdOf(id, 'ann')).toBe(true);
		expect(isGroupChatIdOf(id, 'bob')).toBe(false);
		expect(isGroupChatIdOf('g-g-1-1', 'ann')).toBe(false);
	});

	it('finds the chat id in a room name and refuses case variants of chat rooms', () => {
		const room = chatRoomId('acme', 'g1');
		expect(chatIdOfRoom(room)).toBe('g1');
		expect(chatIdOfRoom('acme')).toBeNull();
		expect(isChatRoomAlias(room)).toBe(false);
		expect(isChatRoomAlias(room.toUpperCase())).toBe(true);
		expect(isChatRoomAlias(room.replace('chat', 'Chat'))).toBe(true);
		expect(isChatRoomAlias('general')).toBe(false);
		expect(classifyRoom(room.toUpperCase())).toBeNull();
	});

	it('keeps direct chats apart per pair', () => {
		expect(directChatId('ann', 'bob')).not.toBe(directChatId('ann', 'cy'));
	});

	it('pins the call room format', () => {
		expect(callRoomId('acme', 'general')).toBe('call-acme-general');
		expect(callRoomId('acme', 'general')).toMatch(CALL_ROOM_PATTERN);
		expect(callRoomId('w'.repeat(200), 'c').length).toBe(128);
	});

	it('names a private chat call after the chat room, so a server can check its members', () => {
		const chat = createGroupChatId('ann');
		const room = chatCallRoomId('w'.repeat(128), chat);
		expect(room).toBe(`call-${chatRoomId('w'.repeat(128), chat)}`);
		expect(room.length).toBeLessThanOrEqual(128);
		expect(room).toMatch(CALL_ROOM_PATTERN);
		expect(chatRoomOfCall(room)).toBe(chatRoomId('w'.repeat(128), chat));
		expect(chatRoomOfCall(callRoomId('acme', 'general'))).toBeNull();
		expect(isStrayChatCallRoom(room)).toBe(false);
		expect(isStrayChatCallRoom(callRoomId('acme', 'general'))).toBe(false);
	});

	it('marks call rooms no current client opens for a chat', () => {
		const direct = directChatId('ann', 'bob');
		// The shape earlier clients used: the chat id after the workspace, with no room to check.
		expect(isStrayChatCallRoom(callRoomId('acme', direct))).toBe(true);
		expect(isStrayChatCallRoom(callRoomId('acme', createGroupChatId('ann')))).toBe(true);
		expect(
			isStrayChatCallRoom(chatCallRoomId('acme', direct).toUpperCase().replace('CALL', 'call')),
		).toBe(true);
	});

	it('classifies rooms', () => {
		expect(classifyRoom('acme')).toBe('workspace');
		expect(classifyRoom(chatRoomId('acme', 'g1'))).toBe('chat');
		expect(classifyRoom('call-acme-general')).toBe('call');
		expect(classifyRoom('')).toBeNull();
		expect(classifyRoom('a/b')).toBeNull();
		expect(classifyRoom('a'.repeat(129))).toBeNull();
	});

	it('reads the announced protocol version back from a URL', () => {
		const url = endpointUrl('wss://h/sync', 'room', 't', protocolParams());
		expect(url).toBe('wss://h/sync/room?token=t&proto=2');
		expect(protocolOfUrl(url)).toBe(2);
		expect(protocolOfUrl('wss://h/sync/room?token=t')).toBeNull();
		expect(protocolOfUrl('wss://h/sync/room?proto=abc')).toBeNull();
		expect(protocolOfUrl('/sync/room?proto=7')).toBe(7);
	});
});
