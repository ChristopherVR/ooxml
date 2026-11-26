import type { ActivityItem } from './activity';
import { createTeamsClient, type TeamsClient } from './store';

const waitFor = (assertion: () => void) => vi.waitFor(assertion, { interval: 5, timeout: 5_000 });
const ready = (client: TeamsClient) =>
	waitFor(() => {
		expect(client.workspace.session.canWrite()).toBe(true);
		expect(client.getState().selectedChannelId).toBe('general');
	});
const clients: TeamsClient[] = [];
const make = (name: string, workspaceId: string, extra = {}): TeamsClient => {
	const c = createTeamsClient({ workspaceId, user: { id: name, name }, ...extra });
	clients.push(c);
	return c;
};
afterEach(() => {
	for (const c of clients.splice(0)) c.destroy();
});

const ada = { kind: 'user', id: 'ada', name: 'ada' } as const;

/** Ada and Bob in one workspace, both in the seeded General channel. */
async function pair(room: string, adaExtra = {}) {
	const a = make('ada', room, adaExtra);
	const b = make('bob', room);
	await Promise.all([ready(a), ready(b)]);
	await a.send({ text: 'first post' });
	await waitFor(() => expect(b.getState().messages[0]?.text).toBe('first post'));
	const heard: ActivityItem[] = [];
	a.on('notification', (item) => heard.push(item));
	return { a, b, heard };
}

describe('activity through the client', () => {
	it('collects mentions, replies, reactions and pins for you and announces them', async () => {
		const { a, b, heard } = await pair('activity-flow');
		const mine = a.getState().messages[0]!;
		await b.send({ text: 'ping @ada', mentions: [ada] });
		b.startReply(mine.id);
		await b.send({ text: 'a reply' });
		b.toggleReaction(mine.id, '👍');
		b.pinMessage(mine.id);
		await waitFor(() => expect(a.getState().activity).toHaveLength(4));
		const state = a.getState();
		expect(state.activity.map((i) => i.kind).sort()).toEqual([
			'mention',
			'pin',
			'reaction',
			'reply',
		]);
		expect(state.activityUnread).toBe(4);
		expect(heard).toHaveLength(4);
		expect(b.getState().activity).toEqual([]);
	});

	it('marks items and everything read, and keeps the state per user on the device', async () => {
		const values = new Map<string, string>();
		const storage = {
			getItem: (key: string) => values.get(key) ?? null,
			setItem: (key: string, value: string) => void values.set(key, value),
		};
		const { a, b } = await pair('activity-read', { storage });
		await b.send({ text: 'one @ada', mentions: [ada] });
		await b.send({ text: 'two @ada', mentions: [ada] });
		await waitFor(() => expect(a.getState().activity).toHaveLength(2));
		const [newest, older] = a.getState().activity;
		expect(a.getState().activityUnread).toBe(2);
		expect(a.markActivityRead(newest!.id)).toBe(true);
		await waitFor(() => expect(a.getState().activityUnread).toBe(1));
		expect(a.markActivityRead(newest!.id, false)).toBe(true);
		expect(a.markAllActivityRead()).toBe(true);
		await waitFor(() => expect(a.getState().activityUnread).toBe(0));
		expect(a.getState().activity.every((i) => i.read)).toBe(true);
		expect(older?.read).toBe(false);
		a.destroy();
		const again = make('ada', 'activity-read', { storage });
		await ready(again);
		await waitFor(() => expect(again.getState().activity).toHaveLength(2));
		expect(again.getState().activityUnread).toBe(0);
	});

	it('suppresses notifications during Do not disturb but still lists the activity', async () => {
		const { a, b, heard } = await pair('activity-dnd');
		a.setAvailability('dnd');
		await waitFor(() => expect(a.getState().availability).toBe('dnd'));
		await b.send({ text: 'urgent @ada', mentions: [ada] });
		await waitFor(() => expect(a.getState().activity).toHaveLength(1));
		expect(heard).toEqual([]);
		expect(a.getState().activityUnread).toBe(1);
		a.setAvailability('available');
		await waitFor(() => expect(a.getState().availability).toBe('available'));
		await b.send({ text: 'again @ada', mentions: [ada] });
		await waitFor(() => expect(a.getState().activity).toHaveLength(2));
		expect(heard).toHaveLength(1);
	});

	it('respects the mentions-only and off levels', async () => {
		const { a, b, heard } = await pair('activity-levels');
		const mine = a.getState().messages[0]!;
		a.setNotificationSettings({ level: 'mentions', sound: false });
		await waitFor(() =>
			expect(a.getState().notificationSettings).toEqual({ level: 'mentions', sound: false }),
		);
		b.toggleReaction(mine.id, '👍');
		await waitFor(() => expect(a.getState().activity).toHaveLength(1));
		expect(heard).toEqual([]);
		await b.send({ text: 'hey @ada', mentions: [ada] });
		await waitFor(() => expect(a.getState().activity).toHaveLength(2));
		expect(heard.map((i) => i.kind)).toEqual(['mention']);
		a.setNotificationSettings({ level: 'off' });
		await b.send({ text: 'more @ada', mentions: [ada] });
		await waitFor(() => expect(a.getState().activityUnread).toBe(3));
		expect(heard).toHaveLength(1);
	});

	it('does not announce history that was there before the client started', async () => {
		const b = make('bob', 'activity-history');
		const seed = make('seed', 'activity-history');
		await Promise.all([ready(b), ready(seed)]);
		seed.destroy();
		await b.send({ text: 'old @ada', mentions: [ada] });
		await waitFor(() => expect(b.getState().messages[0]?.text).toBe('old @ada'));
		const a = make('ada', 'activity-history');
		const heard: ActivityItem[] = [];
		a.on('notification', (item) => heard.push(item));
		await waitFor(() => expect(a.getState().messages[0]?.text).toBe('old @ada'));
		expect(heard).toEqual([]);
		expect(a.getState().activityUnread).toBe(0);
	});
});
