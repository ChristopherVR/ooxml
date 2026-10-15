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

	it('stores picked names with the @ and shows them without it while editing', async () => {
		const ada = make('ada', 'mention-plain');
		const bob = make('bob', 'mention-plain');
		ada.createChannel('Ops');
		await tick(150);
		await bob.send({ text: 'hi' });
		await tick(150);
		await ada.send({
			text: 'hello bob, the channel is quiet',
			mentions: [{ kind: 'user', id: 'bob', name: 'bob' }],
		});
		await tick(100);
		const message = ada.getState().messages.at(-1)!;
		expect(message.text).toBe('hello @bob, the channel is quiet');
		expect(message.mentions).toHaveLength(1);
		ada.startEdit(message.id);
		await tick(50);
		expect(ada.getState().draft.text).toBe('hello bob, the channel is quiet');
		await ada.send({ text: 'hello bob, all quiet' });
		await tick(100);
		expect(ada.getState().messages.at(-1)!.text).toBe('hello @bob, all quiet');
	});
});
