import { createTeamsClient, type TeamsClient } from './store';

const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));
const clients: TeamsClient[] = [];
const make = (name: string, workspaceId: string, extra = {}): TeamsClient => {
	const c = createTeamsClient({ workspaceId, user: { id: name, name }, ...extra });
	clients.push(c);
	return c;
};
afterEach(() => {
	for (const c of clients.splice(0)) c.destroy();
});

describe('mentions through the client', () => {
	it('offers channel members and groups, then stores and syncs mentions', async () => {
		const ada = make('ada', 'mention-client');
		const bob = make('bob', 'mention-client');
		ada.createChannel('Ops');
		await tick(150);
		await bob.send({ text: 'hello from bob' });
		await tick(150);
		expect(ada.getState().mentionTargets.map((t) => t.name)).toEqual(['bob', 'channel', 'team']);
		await ada.send({ text: 'ping @bob and @channel' });
		await tick(150);
		const message = bob.getState().messages.at(-1)!;
		expect(message.mentions?.map((m) => [m.kind, m.id])).toEqual([
			['user', 'bob'],
			['channel', 'channel'],
		]);
	});

	it('keeps mentions when a message is edited', async () => {
		const ada = make('ada', 'mention-edit');
		const bob = make('bob', 'mention-edit');
		ada.createChannel('Ops');
		await tick(150);
		await bob.send({ text: 'hi' });
		await tick(150);
		await ada.send({ text: 'ping @bob' });
		await tick(100);
		ada.startEdit(ada.getState().messages.at(-1)!.id);
		await tick(50);
		await ada.send({ text: 'ping @bob, please look' });
		await tick(100);
		const edited = ada.getState().messages.at(-1)!;
		expect(edited.text).toBe('ping @bob, please look');
		expect(edited.mentions).toHaveLength(1);
		expect(edited.editedAt).toBeGreaterThan(0);
	});
});
