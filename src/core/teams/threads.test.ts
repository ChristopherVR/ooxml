import { channelThreads } from './threads';
import type { Message } from './model';

function message(id: string, replyTo?: string, deleted = false): Message {
	return {
		id,
		channelId: 'channel',
		authorId: 'ada',
		authorName: 'Ada',
		text: id,
		ts: 1,
		deleted,
		reactions: {},
		attachments: [],
		...(replyTo ? { replyTo } : {}),
	};
}

describe('channel threads', () => {
	it('groups nested replies and retains deleted parents without counting deleted replies', () => {
		const root = message('root', undefined, true);
		const index = channelThreads([
			root,
			message('first', 'root', true),
			message('nested', 'first'),
			message('other'),
		]);
		expect(index.posts.map((post) => post.id)).toEqual(['root', 'other']);
		expect(index.replyCounts.root).toBe(1);
		expect(index.thread('nested')).toEqual({
			root,
			replies: [message('first', 'root', true), message('nested', 'first')],
		});
		expect(index.thread('unknown')).toBeNull();
	});
	it('retains orphaned conversations and gives cycles a stable visible root', () => {
		const index = channelThreads([
			message('orphan', 'missing'),
			message('child', 'orphan'),
			message('b', 'a'),
			message('a', 'b'),
			message('c', 'b'),
		]);
		expect(index.posts.map((post) => post.id)).toEqual(['orphan', 'a']);
		expect(index.thread('child')?.root.id).toBe('orphan');
		expect(index.thread('c')?.root.id).toBe('a');
		expect(index.replyCounts.a).toBe(2);
	});
	it('handles deep reply chains without recursive traversal', () => {
		const messages = Array.from({ length: 5000 }, (_, i) =>
			message(`message-${i}`, i ? `message-${i - 1}` : undefined),
		).reverse();
		const index = channelThreads(messages);
		expect(index.posts).toHaveLength(1);
		expect(index.thread('message-4999')?.replies).toHaveLength(4999);
	});
});
