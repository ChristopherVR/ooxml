import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { createMemoryHub } from '../collab/memory-transport';
import {
	MAX_SHARED_CHAT_CHARS,
	createSharedChat,
	createSharedChatStore,
	isSharedChatRoomId,
	sharedChatRoomId,
} from './shared-chat';

const ada = { id: 'u1', name: 'Ada' };

describe('shared chat store', () => {
	it('keeps messages in order and skips unchanged writes', () => {
		const doc = new Y.Doc();
		const store = createSharedChatStore(doc, ada);
		store.put({ id: 'm1', role: 'user', content: 'Hello', streaming: false, data: { files: [] } });
		store.put({ id: 'm2', role: 'assistant', content: '', streaming: true, data: {} });
		store.put({ id: 'm2', role: 'assistant', content: 'Hi there', streaming: false, data: {} });
		let updates = 0;
		doc.on('update', () => updates++);
		store.put({ id: 'm1', role: 'user', content: 'Hello', streaming: false, data: { files: [] } });
		expect(updates).toBe(0);
		expect(store.messages().map((m) => [m.id, m.content, m.authorName, m.streaming])).toEqual([
			['m1', 'Hello', 'Ada', false],
			['m2', 'Hi there', 'Ada', false],
		]);
	});

	it('keeps the original author when someone else updates a message', () => {
		const doc = new Y.Doc();
		createSharedChatStore(doc, ada).put({
			id: 'm1',
			role: 'assistant',
			content: 'A',
			streaming: false,
			data: {},
		});
		createSharedChatStore(doc, { id: 'u2', name: 'Ben' }).put({
			id: 'm1',
			role: 'assistant',
			content: 'A',
			streaming: false,
			data: { applied: [true] },
		});
		const [message] = createSharedChatStore(doc, ada).messages();
		expect(message?.authorId).toBe('u1');
		expect(message?.data).toEqual({ applied: [true] });
	});

	it('truncates from a message for edit and resend', () => {
		const store = createSharedChatStore(new Y.Doc(), ada);
		for (const id of ['a', 'b', 'c'])
			store.put({ id, role: 'user', content: id, streaming: false, data: {} });
		store.truncateFrom('b');
		expect(store.messages().map((m) => m.id)).toEqual(['a']);
		store.truncateFrom('missing');
		expect(store.messages()).toHaveLength(1);
	});

	it('re-validates what a peer wrote', () => {
		const doc = new Y.Doc();
		const rows = doc.getMap<unknown>('chat:messages');
		const order = doc.getArray<string>('chat:order');
		rows.set('bad-role', { role: 'system', content: 'x' });
		rows.set('long', {
			role: 'user',
			content: 'x'.repeat(MAX_SHARED_CHAT_CHARS + 10),
			authorName: '<b>Eve</b>',
			data: [1],
		});
		order.push(['bad-role', 'long', 'long', 'ghost']);
		const messages = createSharedChatStore(doc, ada).messages();
		expect(messages).toHaveLength(1);
		expect(messages[0]?.content).toHaveLength(MAX_SHARED_CHAT_CHARS);
		expect(messages[0]?.authorName).not.toContain('<');
		expect(messages[0]?.data).toEqual({});
	});

	it('sets a single-line title', () => {
		const store = createSharedChatStore(new Y.Doc(), ada);
		store.setTitle('  Launch\n plan ');
		expect(store.title()).toBe('Launch plan');
	});
});

describe('shared chat rooms', () => {
	it('makes unguessable room ids', () => {
		const id = sharedChatRoomId();
		expect(isSharedChatRoomId(id)).toBe(true);
		expect(id).not.toBe(sharedChatRoomId());
		expect(isSharedChatRoomId('general')).toBe(false);
		expect(isSharedChatRoomId('aichat-../../x')).toBe(false);
	});

	it('syncs messages between two people in one room', async () => {
		const hub = createMemoryHub();
		const roomId = sharedChatRoomId();
		const a = createSharedChat({ roomId, user: ada, transport: hub.createTransport(roomId) });
		const b = createSharedChat({
			roomId,
			user: { id: 'u2', name: 'Ben' },
			transport: hub.createTransport(roomId),
		});
		const seen = new Promise<void>((resolve) => {
			const stop = b.store.observe(() => {
				if (b.store.messages().length === 1) {
					stop();
					resolve();
				}
			});
		});
		a.store.put({
			id: 'm1',
			role: 'user',
			content: 'Plan the offsite',
			streaming: false,
			data: {},
		});
		await seen;
		expect(b.store.messages()[0]).toMatchObject({ content: 'Plan the offsite', authorName: 'Ada' });
		a.setActivity('asking');
		await new Promise((resolve) => setTimeout(resolve, 80));
		expect(b.people()).toEqual([expect.objectContaining({ name: 'Ada', activity: 'asking' })]);
		a.destroy();
		b.destroy();
	});

	it('rejects a malformed room id', () => {
		expect(() => createSharedChat({ roomId: 'a/b', user: ada })).toThrow(/shared chat link/u);
	});
});
