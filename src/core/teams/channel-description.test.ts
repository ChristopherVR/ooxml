import * as Y from 'yjs';
import { describe, expect, it } from 'vitest';
import { createChatStore } from './chat';
import { createTeamsClient } from './store';
import { MAX_CHANNEL_TOPIC_CHARS, sanitizeChannelTopic } from './model';

describe('channel descriptions', () => {
	it('preserves descriptions beyond the name limit across snapshots', () => {
		const doc = new Y.Doc();
		const chat = createChatStore(doc, { id: 'ada', name: 'Ada' });
		const topic = 'Budget planning\n' + 'a'.repeat(300);
		const channel = chat.createChannel({ name: 'Planning', topic })!;
		expect(channel.topic).toBe(topic);
		const restored = new Y.Doc();
		Y.applyUpdate(restored, Y.encodeStateAsUpdate(doc));
		const peer = createChatStore(restored, { id: 'bob', name: 'Bob' });
		expect(peer.channels()[0]?.topic).toBe(topic);
		doc.destroy();
		restored.destroy();
	});
	it('restores selection per user and workspace and falls back when unavailable', async () => {
		const values = new Map<string, string>();
		const storage = {
			getItem: (key: string) => values.get(key) ?? null,
			setItem: (key: string, value: string) => void values.set(key, value),
		};
		const make = (id: string, workspaceId = 'channel-selection') =>
			createTeamsClient({ workspaceId, user: { id, name: id }, storage });
		const ada = make('ada');
		ada.createChannel('First');
		await new Promise((resolve) => setTimeout(resolve, 0));
		ada.createChannel('Planning', 'Keep this description');
		await new Promise((resolve) => setTimeout(resolve, 0));
		ada.destroy();
		const restored = make('ada');
		expect(restored.getState().channel?.name).toBe('Planning');
		expect(restored.getState().channel?.topic).toBe('Keep this description');
		const bob = make('bob');
		expect(bob.getState().channel?.name).toBe('First');
		const elsewhere = make('ada', 'other-selection');
		expect(elsewhere.getState().selectedChannelId).toBe('');
		restored.workspace.chat.archiveChannel(restored.getState().selectedChannelId);
		await new Promise((resolve) => setTimeout(resolve, 0));
		expect(restored.getState().channel?.name).toBe('First');
		for (const client of [restored, bob, elsewhere]) client.destroy();
	});
	it('sanitizes descriptions with their own length limit', () => {
		expect(sanitizeChannelTopic('<b>Budget</b>\u0000\nReview')).toBe('Budget\nReview');
		expect(sanitizeChannelTopic('a'.repeat(2000))).toHaveLength(MAX_CHANNEL_TOPIC_CHARS);
		expect(sanitizeChannelTopic(null)).toBe('');
	});
});
