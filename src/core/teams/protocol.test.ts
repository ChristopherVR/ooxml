import { describe, expect, it } from 'vitest';
import {
	CALL_ROOM_PATTERN,
	CHAT_ROOM_PATTERN,
	DIRECT_CHAT_ID_PATTERN,
	PROTOCOL_PARAM,
	PROTOCOL_VERSION,
	callRoomId,
	chatRoomId,
	classifyRoom,
	directChatId,
	protocolOfUrl,
	protocolParams,
} from './protocol';
import { endpointUrl } from './server-config';

describe('protocol version and room names', () => {
	it('pins the protocol version and its query parameter', () => {
		expect(PROTOCOL_VERSION).toBe(1);
		expect(PROTOCOL_PARAM).toBe('proto');
		expect(protocolParams()).toEqual({ proto: '1' });
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
		expect(directChatId('ann', 'bob')).toBe('dm-d50d61311f9ec6df');
		expect(directChatId('bob', 'ann')).toBe('dm-d50d61311f9ec6df');
		expect(directChatId('ann', 'bob')).toMatch(DIRECT_CHAT_ID_PATTERN);
		expect(chatRoomId('acme', 'g1')).toBe('chat-45fd71af8d5cbb5d-g1');
	});

	it('keeps direct chats apart per pair', () => {
		expect(directChatId('ann', 'bob')).not.toBe(directChatId('ann', 'cy'));
	});

	it('pins the call room format', () => {
		expect(callRoomId('acme', 'general')).toBe('call-acme-general');
		expect(callRoomId('acme', 'general')).toMatch(CALL_ROOM_PATTERN);
		expect(callRoomId('w'.repeat(200), 'c').length).toBe(128);
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
		expect(url).toBe('wss://h/sync/room?token=t&proto=1');
		expect(protocolOfUrl(url)).toBe(1);
		expect(protocolOfUrl('wss://h/sync/room?token=t')).toBeNull();
		expect(protocolOfUrl('wss://h/sync/room?proto=abc')).toBeNull();
		expect(protocolOfUrl('/sync/room?proto=7')).toBe(7);
	});
});
