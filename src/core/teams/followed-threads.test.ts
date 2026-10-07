import { createThreadFollows } from './followed-threads.js';
import { channelThreads } from './threads.js';
import type { Channel, Message } from './model.js';

const channel: Channel = {
	id: 'channel',
	name: 'Project',
	topic: '',
	kind: 'channel',
	createdAt: 1,
	createdBy: 'ada',
	archived: false,
};
function message(id: string, ts: number, replyTo?: string): Message {
	return {
		id,
		ts,
		channelId: channel.id,
		authorId: 'ada',
		authorName: 'Ada',
		text: id,
		deleted: false,
		attachments: [],
		reactions: {},
		...(replyTo ? { replyTo } : {}),
	};
}

describe('personal followed threads', () => {
	it('tracks foreign unread messages independently and persists explicit unread marks', () => {
		const values = new Map<string, string>();
		const storage = {
			getItem: (key: string) => values.get(key) ?? null,
			setItem: (key: string, value: string) => {
				values.set(key, value);
			},
		};
		const follows = createThreadFollows(storage, 'ada');
		follows.set(channel.id, 'root', true);
		const messages = [
			message('root', 1),
			{ ...message('reply', 2, 'root'), authorId: 'bob' },
			{ ...message('deleted', 3, 'root'), authorId: 'bob', deleted: true },
		];
		expect(follows.view([channel], () => messages, 'ada')[0]?.unread).toBe(1);
		expect(follows.mark(channel.id, 'root', 3, false)).toBe(true);
		expect(follows.view([channel], () => messages, 'ada')[0]?.unread).toBe(0);
		expect(follows.mark(channel.id, 'root', 3, false)).toBe(false);
		messages.push({ ...message('new', 4, 'root'), authorId: 'bob' });
		expect(follows.view([channel], () => messages, 'ada')[0]?.unread).toBe(1);
		follows.mark(channel.id, 'root', 4, false);
		follows.mark(channel.id, 'root', 0, true);
		expect(follows.view([channel], () => messages, 'ada')[0]?.unread).toBe(1);
		expect(
			createThreadFollows(storage, 'ada').view([channel], () => [message('root', 1)], 'ada')[0]
				?.unread,
		).toBe(1);
	});
	it('persists references, derives live counts and orders by activity', () => {
		const values = new Map<string, string>();
		const storage = {
			getItem: (key: string) => values.get(key) ?? null,
			setItem: (key: string, value: string) => {
				values.set(key, value);
			},
		};
		const follows = createThreadFollows(storage, 'ada');
		follows.set(channel.id, 'old', true);
		follows.set(channel.id, 'new', true);
		const messages = [message('old', 1), message('new', 2), message('reply', 3, 'old')];
		const restored = createThreadFollows(storage, 'ada');
		expect(
			restored.view([channel], () => messages).map((row) => [row.root.id, row.replyCount]),
		).toEqual([
			['old', 1],
			['new', 0],
		]);
		expect(createThreadFollows(storage, 'bob').view([channel], () => messages)).toEqual([]);
		expect(restored.view([{ ...channel, archived: true }], () => messages)).toEqual([]);
	});
	it('retains follows across late parent arrival and removes every merged reference', () => {
		const follows = createThreadFollows(undefined, 'ada');
		follows.set(channel.id, 'child', true);
		follows.set(channel.id, 'root', true);
		const messages = [message('root', 1), message('child', 2, 'root')];
		const index = channelThreads(messages);
		expect(follows.view([channel], () => messages)).toHaveLength(1);
		expect(follows.has(channel.id, 'root', (id) => index.thread(id)?.root.id)).toBe(true);
		expect(follows.set(channel.id, 'root', false, (id) => index.thread(id)?.root.id)).toBe(true);
		expect(follows.view([channel], () => messages)).toEqual([]);
	});
	it('ignores corrupt references and survives blocked storage', () => {
		const follows = createThreadFollows(
			{
				getItem: () => '[null,{}, {"channelId":"channel","messageId":5}]',
				setItem: () => {
					throw new Error('blocked');
				},
			},
			'ada',
		);
		expect(follows.view([channel], () => [message('root', 1)])).toEqual([]);
		expect(follows.set(channel.id, 'root', true)).toBe(true);
		expect(follows.has(channel.id, 'root')).toBe(true);
		expect(follows.set('', 'root', true)).toBe(false);
	});
});
