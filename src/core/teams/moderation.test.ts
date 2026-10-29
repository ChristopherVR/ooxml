import { afterEach, describe, expect, it } from 'vitest';
import { type TeamsClient, createTeamsClient } from './store';
import { requestTeamsSession } from './server-session';
import { localServerConfig } from './server-config';

const clients: TeamsClient[] = [];
let room = 0;
const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));
async function until(fn: () => boolean): Promise<boolean> {
	for (let i = 0; i < 100; i++) {
		if (fn()) return true;
		await sleep(10);
	}
	return fn();
}
const make = (name: string, workspaceId: string, moderator = false): TeamsClient => {
	const c = createTeamsClient({
		workspaceId,
		user: { id: name, name },
		moderator,
		reminderCheckMs: 0,
	});
	clients.push(c);
	return c;
};
afterEach(() => {
	for (const c of clients.splice(0)) c.destroy();
});

describe('moderation in the client', () => {
	it('exposes the role and lets only a moderator hide, rename and archive what is not theirs', async () => {
		const ws = `moderation-${Date.now()}-${room++}`;
		const ann = make('ann', ws);
		const mod = make('mod', ws, true);
		const bob = make('bob', ws);
		expect(ann.getState().moderator).toBe(false);
		expect(mod.getState().moderator).toBe(true);

		ann.createChannel('Plans');
		await until(() => ann.getState().channel?.name === 'Plans');
		const channelId = ann.getState().selectedChannelId;
		expect(await until(() => bob.getState().channels.some((c) => c.id === channelId))).toBe(true);
		await ann.send({ text: 'my message' });
		const messageId = ann.getState().messages.find((m) => m.text === 'my message')!.id;
		bob.select(channelId);
		mod.select(channelId);
		expect(await until(() => bob.getState().messages.some((m) => m.id === messageId))).toBe(true);
		expect(await until(() => mod.getState().messages.some((m) => m.id === messageId))).toBe(true);

		// A plain member cannot hide it; a moderator can.
		bob.deleteMessage(messageId);
		expect(bob.getState().messages.find((m) => m.id === messageId)!.deleted).toBe(false);
		mod.deleteMessage(messageId);
		expect(
			await until(() => mod.getState().messages.find((m) => m.id === messageId)?.deleted === true),
		).toBe(true);
		expect(
			await until(() => ann.getState().messages.find((m) => m.id === messageId)!.deleted),
		).toBe(true);

		// Channels: the creator and a moderator manage them, others do not.
		expect(bob.renameChannel(channelId, 'Hijack')).toBe(false);
		expect(bob.archiveChannel(channelId)).toBe(false);
		expect(mod.renameChannel(channelId, 'Plans 2')).toBe(true);
		expect(await until(() => ann.getState().channels.some((c) => c.name === 'Plans 2'))).toBe(true);
		expect(ann.archiveChannel(channelId)).toBe(true);
		expect(await until(() => !mod.getState().channels.some((c) => c.id === channelId))).toBe(true);
	});

	it('lets a moderator rename and remove a tab they did not create', async () => {
		const ws = `moderation-${Date.now()}-${room++}`;
		const ann = make('ann', ws);
		const mod = make('mod', ws, true);
		const bob = make('bob', ws);
		await until(() => ann.getState().channel !== null);
		const tab = ann.addTab('Docs', { type: 'website', url: 'https://example.com/' })!;
		expect(await until(() => mod.getState().tabs.some((t) => t.id === tab.id))).toBe(true);
		expect(await until(() => bob.getState().tabs.some((t) => t.id === tab.id))).toBe(true);
		expect(bob.renameTab(tab.id, 'Nope')).toBe(false);
		expect(bob.removeTab(tab.id)).toBe(false);
		expect(mod.renameTab(tab.id, 'Renamed')).toBe(true);
		expect(mod.removeTab(tab.id)).toBe(true);
	});
});

describe('session role', () => {
	const respond = (body: Record<string, unknown>): typeof fetch =>
		(async () => new Response(JSON.stringify(body), { status: 200 })) as unknown as typeof fetch;
	const config = {
		...localServerConfig(),
		mode: 'server' as const,
		syncUrl: 'wss://example.com/sync',
	};

	it('reads the role the server granted', async () => {
		const session = await requestTeamsSession(
			config,
			{ userId: 'ann', credential: 'x' },
			{ fetch: respond({ session: 's', user: 'ann', expiresAt: 1, role: 'moderator' }) },
		);
		expect(session.role).toBe('moderator');
	});

	it('ignores an unknown or missing role', async () => {
		for (const role of [undefined, 'admin', 5]) {
			const session = await requestTeamsSession(
				config,
				{ userId: 'ann', credential: 'x' },
				{ fetch: respond({ session: 's', user: 'ann', expiresAt: 1, role }) },
			);
			expect(session.role).toBeUndefined();
		}
	});
});
