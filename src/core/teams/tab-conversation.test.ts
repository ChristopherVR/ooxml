import * as Y from 'yjs';
import { createChatStore } from './chat';
import { createTabStore } from './tabs';
import { tabConversationId } from './tab-conversation';
import { channelThreads } from './threads';
import { createTeamsClient } from './store';

describe('tab conversations', () => {
	it('converges concurrent starts without losing replies and survives snapshots', () => {
		const a = new Y.Doc(),
			b = new Y.Doc(),
			restored = new Y.Doc();
		const ada = createChatStore(a, { id: 'ada', name: 'Ada' });
		const bob = createChatStore(b, { id: 'bob', name: 'Bob' });
		ada.createChannel({ name: 'General', id: 'general' });
		const tab = createTabStore(a, { id: 'ada', name: 'Ada' }, () => true).add('general', 'Budget', {
			type: 'website',
			url: 'https://example.com',
		})!;
		Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
		const first = ada.ensureTabConversation(tab)!;
		const second = bob.ensureTabConversation(tab)!;
		expect(first.id).toBe(second.id);
		ada.post('general', { text: 'Ada reply', replyTo: first.id });
		bob.post('general', { text: 'Bob reply', replyTo: second.id });
		Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
		Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
		expect(ada.messages('general')).toEqual(bob.messages('general'));
		const index = channelThreads(ada.messages('general'));
		expect(index.posts).toHaveLength(1);
		expect(index.posts[0]!.tabId).toBe(tab.id);
		expect(
			index
				.thread(first.id)!
				.replies.map((message) => message.text)
				.sort(),
		).toEqual(['Ada reply', 'Bob reply']);
		const before = Y.encodeStateAsUpdate(a);
		ada.ensureTabConversation(tab);
		expect(Y.encodeStateAsUpdate(a)).toEqual(before);
		Y.applyUpdate(restored, before);
		const recovered = createChatStore(restored, { id: 'ada', name: 'Ada' });
		expect(recovered.messages('general')).toEqual(ada.messages('general'));
		const root = index.posts[0]!;
		(root.authorId === 'ada' ? ada : bob).remove('general', root.id);
		Y.applyUpdate(a, Y.encodeStateAsUpdate(b));
		Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
		expect(ada.ensureTabConversation(tab)!.deleted).toBe(true);
		expect(channelThreads(ada.messages('general')).thread(root.id)!.replies).toHaveLength(2);
		ada.destroy();
		bob.destroy();
		recovered.destroy();
		a.destroy();
		b.destroy();
		restored.destroy();
	});
	it('rejects malformed associations and unknown channels, and supports maximum tab IDs', () => {
		const doc = new Y.Doc();
		const chat = createChatStore(doc, { id: 'ada', name: 'Ada' });
		chat.createChannel({ name: 'General', id: 'general' });
		const message = chat.post('general', { text: 'Ordinary post' })!;
		doc.getMap<Y.Map<unknown>>('teams:messages:general').get(message.id)!.set('tabId', 'forged');
		expect(chat.messages('general')[0]).not.toHaveProperty('tabId');
		expect(tabConversationId('bad|id')).toBeNull();
		expect(tabConversationId('x'.repeat(161))).toBeNull();
		expect(tabConversationId('x'.repeat(160))!.length).toBe(68);
		expect(tabConversationId('x'.repeat(160))).not.toBe(tabConversationId('x'.repeat(159)));
		const tab = {
			id: 'valid',
			channelId: 'missing',
			name: 'Site',
			content: { type: 'website' as const, url: 'https://example.com' },
			createdBy: 'ada',
			createdAt: 0,
		};
		expect(chat.ensureTabConversation(tab)).toBeNull();
		chat.destroy();
		doc.destroy();
	});
	it('posts only when requested and opens one shared thread without resetting its draft', async () => {
		const tick = () => new Promise((resolve) => setTimeout(resolve, 0));
		const client = createTeamsClient({
			workspaceId: `tab-conversation-${Date.now()}`,
			user: { id: 'ada', name: 'Ada' },
			config: { mode: 'local', iceServers: [] },
		});
		const website = { type: 'website' as const, url: 'https://example.com' };
		client.createChannel('General');
		await tick();
		const quiet = client.addTab('Quiet', website)!;
		expect(client.getState().posts).toHaveLength(0);
		const announced = client.addTab('Budget', website, { postToChannel: true })!;
		await tick();
		expect(client.getState().posts.map((message) => message.tabId)).toEqual([announced.id]);
		expect(client.openTabConversation(quiet.id)).toBe(true);
		await tick();
		client.setDraft({ text: 'Keep tab reply', files: [], missingFiles: [] });
		await tick();
		client.closeThread();
		expect(client.openTabConversation(quiet.id)).toBe(true);
		await tick();
		expect(client.getState().draft.text).toBe('Keep tab reply');
		expect(client.getState().thread!.root.tabId).toBe(quiet.id);
		expect(client.getState().posts).toHaveLength(2);
		client.removeTab(quiet.id);
		expect(client.openTabConversation(quiet.id)).toBe(false);
		expect(client.getState().posts).toHaveLength(2);
		client.destroy();
		expect(client.addTab('After teardown', website)).toBeNull();
		expect(client.openTabConversation(announced.id)).toBe(false);
	});
});
