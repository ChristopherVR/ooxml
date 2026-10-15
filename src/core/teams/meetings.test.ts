import * as Y from 'yjs';
import { createMeetingReminders } from './meeting-reminders';
import { createMeetingStore, meetingVisibleTo } from './meetings';
import { createTeamsClient, type TeamsClient } from './store';
import type { ActivityItem } from './activity';

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));
const until = async (fn: () => boolean, ms = 3000): Promise<boolean> => {
	const end = Date.now() + ms;
	while (Date.now() < end) {
		if (fn()) return true;
		await sleep(10);
	}
	return fn();
};
const HOUR = 3_600_000;
const clients: TeamsClient[] = [];
let room = 0;
const workspace = (): string => `meet-${Date.now()}-${room++}`;
const make = (name: string, workspaceId: string, extra = {}): TeamsClient => {
	const c = createTeamsClient({
		workspaceId,
		user: { id: name, name: name[0]!.toUpperCase() + name.slice(1) },
		reminderCheckMs: 0,
		...extra,
	});
	clients.push(c);
	return c;
};
afterEach(() => {
	for (const c of clients.splice(0)) c.destroy();
});

describe('meeting store', () => {
	const ada = { id: 'ada', name: 'Ada' };
	it('validates, stores UTC with the organizer zone and restricts changes to the organizer', () => {
		const doc = new Y.Doc();
		const store = createMeetingStore(doc, ada);
		expect(store.create({ title: '  ', start: 1, end: 2 })).toBeNull();
		expect(store.create({ title: 'Bad', start: 5000, end: 4000 })).toBeNull();
		expect(store.create({ title: 'Long', start: 1000, end: 1000 + 25 * HOUR })).toBeNull();
		const start = Date.UTC(2026, 9, 20, 13);
		const m = store.create({
			title: 'Review <b>now</b>',
			start,
			end: start + HOUR,
			zone: 'Europe/Berlin',
			repeat: 'weekly',
			attendees: [
				{ id: 'bob', name: 'Bob' },
				{ id: 'ada', name: 'Ada' },
			],
		})!;
		expect(m).toMatchObject({
			title: 'Review now',
			zone: 'Europe/Berlin',
			start,
			repeat: 'weekly',
		});
		expect(m.attendees.map((a) => a.id)).toEqual(['bob']);
		const bob = createMeetingStore(doc, { id: 'bob', name: 'Bob' });
		expect(bob.update(m.id, { title: 'Hijacked' })).toBeNull();
		expect(bob.cancel(m.id)).toBe(false);
		expect(store.update(m.id, { title: 'Review 2', repeat: 'daily' })).toMatchObject({
			title: 'Review 2',
			repeat: 'daily',
		});
		expect(meetingVisibleTo(m, 'bob')).toBe(true);
		expect(meetingVisibleTo(m, 'cy')).toBe(false);
		expect(store.cancel(m.id)).toBe(true);
		expect(meetingVisibleTo(store.get(m.id)!, 'ada')).toBe(false);
	});

	it('ignores malformed rows written by a peer', () => {
		const doc = new Y.Doc();
		const rows = doc.getMap<Y.Map<unknown>>('teams:meetings');
		const bad = new Y.Map<unknown>();
		bad.set('title', 'No times');
		bad.set('organizerId', 'eve');
		rows.set('bad', bad);
		expect(createMeetingStore(doc, ada).list()).toEqual([]);
	});
});

describe('reminders', () => {
	const m = (over = {}) => ({
		id: 'r1',
		title: 'Sync',
		description: '',
		organizerId: 'ada',
		organizerName: 'Ada',
		zone: 'UTC',
		start: 10 * HOUR,
		end: 11 * HOUR,
		repeat: 'none' as const,
		attendees: [],
		createdAt: 0,
		cancelled: false,
		...over,
	});
	it('fires once per occurrence five minutes ahead and survives a reload', () => {
		const values = new Map<string, string>();
		const storage = {
			getItem: (k: string) => values.get(k) ?? null,
			setItem: (k: string, v: string) => void values.set(k, v),
		};
		const allow = () => true;
		const first = createMeetingReminders(storage, 'k');
		expect(first.due([m()], 'ada', 10 * HOUR - 6 * 60_000, allow)).toEqual([]);
		const due = first.due([m()], 'ada', 10 * HOUR - 4 * 60_000, allow);
		expect(due).toHaveLength(1);
		expect(due[0]).toMatchObject({
			kind: 'meeting',
			channelName: 'Sync',
			preview: 'Starts in 4 minutes',
		});
		expect(first.due([m()], 'ada', 10 * HOUR - 3 * 60_000, allow)).toEqual([]);
		expect(
			createMeetingReminders(storage, 'k').due([m()], 'ada', 10 * HOUR - 2 * 60_000, allow),
		).toEqual([]);
	});

	it('holds a reminder back while it is not allowed and releases it when it is', () => {
		const r = createMeetingReminders(undefined, 'k');
		const now = 10 * HOUR - 4 * 60_000;
		expect(r.due([m()], 'ada', now, () => false)).toEqual([]);
		expect(r.due([m()], 'ada', now + 60_000, () => true)).toHaveLength(1);
	});

	it('does not remind people who are not on the meeting', () => {
		const r = createMeetingReminders(undefined, 'k');
		expect(r.due([m()], 'cy', 10 * HOUR - 60_000, () => true)).toEqual([]);
	});
});

describe('meetings in the client', () => {
	const soon = (): { start: number; end: number } => {
		const start = Date.now() + 2 * HOUR;
		return { start, end: start + HOUR };
	};

	it('schedules a meeting with attendees, a meeting chat and a card in it', async () => {
		const ws = workspace();
		const ada = make('ada', ws);
		const bob = make('bob', ws);
		const cy = make('cy', ws);
		expect(await until(() => ada.getState().contacts.length === 2)).toBe(true);
		const meeting = ada.createMeeting({
			title: 'Planning',
			description: 'Q4',
			...soon(),
			repeat: 'daily',
			attendeeIds: ['bob'],
		})!;
		expect(meeting.chatId).toBeTruthy();
		expect(await until(() => bob.getState().meetings.some((x) => x.id === meeting.id))).toBe(true);
		expect(bob.getState().meetings[0]).toMatchObject({
			title: 'Planning',
			mine: false,
			live: false,
		});
		await sleep(50);
		expect(cy.getState().meetings).toEqual([]);
		// The meeting chat is a group chat tied to the meeting, with a card as its first message.
		expect(await until(() => bob.getState().chats.some((c) => c.meetingId === meeting.id))).toBe(
			true,
		);
		bob.select(meeting.chatId!);
		expect(await until(() => bob.getState().messages.some((x) => x.meetingId === meeting.id))).toBe(
			true,
		);
		expect(cy.getState().chats).toEqual([]);
		expect(bob.updateMeeting(meeting.id, { title: 'x' })).toBeNull();
		expect(ada.getState().meetings[0]!.mine).toBe(true);
	});

	it('posts a meeting card in the channel for a channel meeting', async () => {
		const ws = workspace();
		const ada = make('ada', ws);
		const bob = make('bob', ws);
		await until(() => ada.getState().channels.length > 0 && bob.getState().channels.length > 0);
		const channelId = ada.getState().channels[0]!.id;
		const meeting = ada.createMeeting({ title: 'All hands', ...soon(), channelId })!;
		expect(meeting.channelId).toBe(channelId);
		bob.select(channelId);
		expect(await until(() => bob.getState().messages.some((x) => x.meetingId === meeting.id))).toBe(
			true,
		);
		expect(bob.getState().meetings.map((x) => x.id)).toEqual([meeting.id]);
	});

	it('joins a meeting through the call pre-join and exports an .ics', async () => {
		const ws = workspace();
		const ada = make('ada', ws);
		const meeting = ada.createMeeting({ title: 'Solo focus', ...soon() })!;
		expect(await ada.joinMeeting(meeting.id)).toBe(true);
		await sleep(10);
		expect(ada.getState().call).toMatchObject({
			kind: 'meeting',
			channelName: 'Solo focus',
			phase: 'prejoin',
		});
		expect(await ada.joinMeeting('missing')).toBe(false);
		const ics = ada.meetingIcs(meeting.id)!;
		expect(ics.name).toBe('Solo-focus.ics');
		expect(ics.text).toContain('SUMMARY:Solo focus');
		expect(ada.cancelMeeting(meeting.id)).toBe(true);
		await sleep(10);
		expect(ada.getState().meetings).toEqual([]);
	});

	it('announces a reminder five minutes ahead, but not in Do not disturb', async () => {
		let now = Date.now();
		const ws = workspace();
		const ada = make('ada', ws, { now: () => now, reminderCheckMs: 15 });
		const seen: ActivityItem[] = [];
		ada.on('notification', (item) => seen.push(item));
		const start = now + 20 * 60_000;
		const meeting = ada.createMeeting({ title: 'Reminder me', start, end: start + HOUR })!;
		await sleep(60);
		expect(seen).toEqual([]);
		ada.setAvailability('dnd');
		now = start - 4 * 60_000;
		await sleep(80);
		expect(seen).toEqual([]);
		ada.setAvailability('available');
		await sleep(80);
		expect(seen).toHaveLength(1);
		expect(seen[0]).toMatchObject({
			kind: 'meeting',
			channelName: 'Reminder me',
			messageId: meeting.id,
		});
		await sleep(60);
		expect(seen).toHaveLength(1);
	});
});
