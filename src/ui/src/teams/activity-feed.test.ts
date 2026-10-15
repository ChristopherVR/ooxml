import type { ActivityItem } from 'ooxml-core/teams';
import { registerOfficeUi } from '../index';
import { type OfficeUiActivityFeed, activityTime } from './activity-feed';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

const item = (over: Partial<ActivityItem> & { id: string }): ActivityItem => ({
	kind: 'mention',
	channelId: 'ops',
	channelName: 'Ops',
	messageId: 'm1',
	rootId: 'm1',
	actorId: 'bob',
	actorName: 'Bob',
	ts: Date.now(),
	preview: 'hello there',
	read: false,
	...over,
});

async function mount(items: ActivityItem[]): Promise<OfficeUiActivityFeed> {
	const el = document.createElement('office-ui-activity-feed') as OfficeUiActivityFeed;
	el.items = items;
	document.body.append(el);
	await el.updateComplete;
	return el;
}
const qa = <T extends Element>(el: HTMLElement, selector: string): T[] => [
	...el.shadowRoot!.querySelectorAll<T>(selector),
];
const events = (el: HTMLElement, type: string): unknown[] => {
	const out: unknown[] = [];
	el.addEventListener(type, (e) => out.push((e as CustomEvent).detail));
	return out;
};

describe('office-ui-activity-feed', () => {
	const items = [
		item({ id: 'a', kind: 'mention' }),
		item({ id: 'b', kind: 'reply', read: true }),
		item({ id: 'c', kind: 'reaction', emoji: '👍', read: true }),
		item({ id: 'd', kind: 'pin' }),
	];

	it('lists what happened, flags unread rows and describes each kind', async () => {
		const el = await mount(items);
		const lines = qa(el, '.line').map((l) => l.textContent);
		expect(lines).toEqual([
			'Bob mentioned you in # Ops',
			'Bob replied in # Ops',
			'Bob reacted 👍 to your message',
			'Bob pinned a message in # Ops',
		]);
		expect(qa<HTMLElement>(el, 'li').map((li) => li.dataset.unread)).toEqual([
			'true',
			'false',
			'false',
			'true',
		]);
		expect(qa(el, '.dot')).toHaveLength(2);
	});

	it('filters by Unread, Mentions, Replies and Reactions and explains an empty filter', async () => {
		const el = await mount(items);
		const filters = qa<HTMLButtonElement>(el, '.filters button');
		expect(filters.map((b) => b.textContent!.trim())).toEqual([
			'All',
			'Unread 2',
			'Mentions',
			'Replies',
			'Reactions',
		]);
		const picked = events(el, 'office-activity-filter');
		filters[1]!.click();
		await el.updateComplete;
		expect(qa(el, 'li')).toHaveLength(2);
		filters[3]!.click();
		await el.updateComplete;
		expect(qa(el, '.line').map((l) => l.textContent)).toEqual(['Bob replied in # Ops']);
		expect(picked).toEqual([{ filter: 'unread' }, { filter: 'replies' }]);
		el.items = [item({ id: 'x', read: true })];
		el.filter = 'unread';
		await el.updateComplete;
		expect(el.shadowRoot!.querySelector('.empty')?.textContent).toBe('You are all caught up.');
	});

	it('emits open, per-row read toggles and mark all as read', async () => {
		const el = await mount(items);
		const opened = events(el, 'office-activity-open');
		const toggled = events(el, 'office-activity-read');
		const all = events(el, 'office-activity-read-all');
		qa<HTMLButtonElement>(el, '.item')[0]!.click();
		qa<HTMLButtonElement>(el, '.toggle')[0]!.click();
		qa<HTMLButtonElement>(el, '.toggle')[1]!.click();
		el.shadowRoot!.querySelector<HTMLButtonElement>('.read-all')!.click();
		expect(opened).toEqual([{ id: 'a' }]);
		expect(toggled).toEqual([
			{ id: 'a', read: true },
			{ id: 'b', read: false },
		]);
		expect(all).toHaveLength(1);
	});

	it('disables Mark all as read when everything is read and renders text only', async () => {
		const el = await mount([
			item({
				id: 'x',
				read: true,
				preview: '<img src=x onerror=alert(1)>',
				actorName: '<b>Bob</b>',
			}),
		]);
		expect(el.shadowRoot!.querySelector<HTMLButtonElement>('.read-all')!.disabled).toBe(true);
		expect(el.shadowRoot!.querySelector('img, b')).toBeNull();
		expect(el.shadowRoot!.querySelector('.preview')!.textContent).toContain('<img');
	});

	it('shows the time today and the date otherwise', () => {
		const now = new Date(2026, 9, 15, 14, 5).getTime();
		expect(activityTime(now, now)).toMatch(/\d/u);
		expect(activityTime(now - 86_400_000, now)).toBe('Yesterday');
		expect(activityTime(now - 5 * 86_400_000, now)).not.toBe('Yesterday');
	});
});
