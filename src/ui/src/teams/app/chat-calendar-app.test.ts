import type { TeamsApp } from './teams-app';
import { defineTeamsApp } from './teams-app';

const LOCAL = JSON.stringify({ mode: 'local', iceServers: [] });
const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));
const q = <T extends Element>(root: ParentNode | null | undefined, s: string): T =>
	root!.querySelector<T>(s)!;
const until = async (fn: () => boolean, ms = 3000): Promise<void> => {
	const end = Date.now() + ms;
	while (!fn() && Date.now() < end) await tick(10);
};

beforeAll(() => defineTeamsApp());
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

let workspace = 0;
async function mount(): Promise<TeamsApp> {
	const el = document.createElement('teams-app') as TeamsApp;
	el.setAttribute('workspace-id', `ccapp-${Date.now()}-${workspace++}`);
	el.setAttribute('user-name', 'Ada');
	el.setAttribute('user-id', 'ada');
	el.setAttribute('server-config', LOCAL);
	document.body.append(el);
	await el.updateComplete;
	return el;
}
/** Put a colleague on the roster, as if they had opened the workspace. */
function addColleague(el: TeamsApp, id: string, name: string): void {
	const doc = el.client!.workspace.session.doc;
	const people = doc.getMap<unknown>('teams:roster');
	// The roster rows are Y.Maps; build one from the same class as the map itself.
	const Row = people.constructor as new () => { set: (k: string, v: unknown) => void };
	doc.transact(() => {
		const row = new Row();
		row.set('name', name);
		row.set('seen', Date.now());
		people.set(id, row);
	});
}
const select = (el: TeamsApp, id: string): void => {
	q<HTMLElement>(el.shadowRoot, 'office-ui-app-rail').dispatchEvent(
		new CustomEvent('office-rail-select', { detail: { id } }),
	);
};

describe('<teams-app> Chat and Calendar', () => {
	it('puts Chat after Activity and Calendar after Teams in the app rail', async () => {
		const el = await mount();
		const rail = q<HTMLElement & { items: { id: string; label: string }[] }>(
			el.shadowRoot,
			'office-ui-app-rail',
		);
		expect(rail.items.map((i) => i.id)).toEqual([
			'activity',
			'chat',
			'teams',
			'calendar',
			'calls',
			'files',
			'followed',
			'drafts',
		]);
		el.remove();
	});

	it('starts a 1:1 chat from New chat and shows it in the chat list and header', async () => {
		const el = await mount();
		el.client!.createChannel('Ops');
		addColleague(el, 'bob', 'Bob');
		select(el, 'chat');
		await el.updateComplete;
		expect(el.rail).toBe('chat');
		expect(q(el.shadowRoot, 'office-ui-chat-sidebar')).not.toBeNull();
		expect(el.shadowRoot!.querySelector('.empty-main')).not.toBeNull();
		await until(() => el.client!.getState().contacts.length === 1);
		q<HTMLElement>(el.shadowRoot, 'office-ui-chat-sidebar').dispatchEvent(
			new CustomEvent('office-chat-new'),
		);
		await el.updateComplete;
		const dialog = q<
			HTMLElement & { create: (ids: string[], name: string) => boolean; contacts: unknown[] }
		>(el.shadowRoot, 'teams-new-chat-dialog');
		expect(dialog.hasAttribute('open')).toBe(true);
		expect(dialog.contacts).toHaveLength(1);
		expect(dialog.create(['bob'], '')).toBe(true);
		await tick(10);
		await el.updateComplete;
		const state = el.client!.getState();
		expect(state.chats.map((c) => c.title)).toEqual(['Bob']);
		expect(el.shadowRoot!.querySelector('.chat-head h1')!.textContent).toBe('Bob');
		expect(
			q<HTMLElement>(el.shadowRoot, 'office-ui-chat-composer').getAttribute('placeholder'),
		).toBe('Message Bob');
		// Back on the Teams rail the channel you were in returns.
		select(el, 'teams');
		await tick(10);
		await el.updateComplete;
		expect(el.client!.getState().channel?.name).toBe('Ops');
		el.remove();
	});

	it('creates a named group chat and manages its people from the header', async () => {
		const el = await mount();
		addColleague(el, 'bob', 'Bob');
		addColleague(el, 'cy', 'Cy');
		await until(() => el.client!.getState().contacts.length === 2);
		select(el, 'chat');
		await el.updateComplete;
		const dialog = q<HTMLElement & { create: (ids: string[], name: string) => boolean }>(
			el.shadowRoot,
			'teams-new-chat-dialog',
		);
		expect(dialog.create(['bob', 'cy'], 'Launch')).toBe(true);
		await tick(10);
		await el.updateComplete;
		expect(el.shadowRoot!.querySelector('.chat-head h1')!.textContent).toBe('Launch');
		(el.shadowRoot!.querySelector('.chat-people') as HTMLButtonElement).click();
		await el.updateComplete;
		const people = q<HTMLElement>(el.shadowRoot, 'teams-chat-people-dialog');
		expect(people.hasAttribute('open')).toBe(true);
		const id = el.client!.getState().selectedChannelId;
		people.dispatchEvent(
			new CustomEvent('teams-chat-rename', { detail: { id, name: 'Launch crew' } }),
		);
		await tick(10);
		expect(el.client!.getState().chats[0]!.title).toBe('Launch crew');
		people.dispatchEvent(new CustomEvent('teams-chat-remove', { detail: { id, userId: 'cy' } }));
		await tick(10);
		expect(el.client!.getState().chats[0]!.members.map((m) => m.id)).toEqual(['ada', 'bob']);
		el.remove();
	});

	it('shows system notices as plain lines and a meeting card with Join', async () => {
		const el = await mount();
		addColleague(el, 'bob', 'Bob');
		addColleague(el, 'cy', 'Cy');
		await until(() => el.client!.getState().contacts.length === 2);
		const client = el.client!;
		const start = Date.now() + 3_600_000;
		const meeting = client.createMeeting({
			title: 'Design review',
			start,
			end: start + 3_600_000,
			attendeeIds: ['bob', 'cy'],
		})!;
		await tick(10);
		select(el, 'chat');
		el.client!.select(meeting.chatId!);
		await tick(20);
		await el.updateComplete;
		const list = q<HTMLElement>(el.shadowRoot, 'office-ui-chat-list');
		await (list as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		const card = list.shadowRoot!.querySelector('.meeting-card');
		expect(card?.querySelector('.meeting-title')?.textContent).toBe('Design review');
		client.addToChat(meeting.chatId!, []);
		client.renameChat(meeting.chatId!, 'Design review chat');
		await tick(20);
		await (list as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		expect(list.shadowRoot!.querySelector('.system')?.textContent).toContain('named the chat');
		list.dispatchEvent(
			new CustomEvent('office-chat-meeting', { detail: { meetingId: meeting.id, action: 'join' } }),
		);
		await until(() => el.client!.getState().call !== null);
		expect(el.client!.getState().call).toMatchObject({
			kind: 'meeting',
			channelName: 'Design review',
		});
		el.remove();
	});

	it('schedules a meeting from the calendar and joins it through the pre-join screen', async () => {
		const el = await mount();
		select(el, 'calendar');
		await el.updateComplete;
		expect(el.rail).toBe('calendar');
		expect(q(el.shadowRoot, 'office-ui-calendar-agenda')).not.toBeNull();
		expect(q(el.shadowRoot, 'office-ui-calendar-view')).not.toBeNull();
		q<HTMLElement>(el.shadowRoot, 'office-ui-calendar-view').dispatchEvent(
			new CustomEvent('office-calendar-new', { detail: { start: 0 } }),
		);
		await el.updateComplete;
		const form = q<HTMLElement & { save: (f: unknown, id?: string) => boolean }>(
			el.shadowRoot,
			'teams-meeting-dialog',
		);
		expect(form.hasAttribute('open')).toBe(true);
		const start = Date.now() + 1_800_000;
		expect(form.save({ title: 'Stand-up', start, end: start + 1_800_000, repeat: 'daily' })).toBe(
			true,
		);
		await tick(10);
		await el.updateComplete;
		expect(form.hasAttribute('open')).toBe(false);
		const meeting = el.client!.getState().meetings[0]!;
		expect(meeting).toMatchObject({ title: 'Stand-up', repeat: 'daily', mine: true });
		const agenda = q<HTMLElement>(el.shadowRoot, 'office-ui-calendar-agenda');
		await (agenda as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		expect(agenda.shadowRoot!.querySelector('.row .title')?.textContent).toBe('Stand-up');
		agenda.dispatchEvent(
			new CustomEvent('office-agenda-open', { detail: { meetingId: meeting.id, start } }),
		);
		await el.updateComplete;
		const details = q<HTMLElement>(el.shadowRoot, 'teams-meeting-details');
		expect(details.hasAttribute('open')).toBe(true);
		details.dispatchEvent(
			new CustomEvent('teams-meeting-join', { detail: { meetingId: meeting.id } }),
		);
		await until(() => el.client!.getState().call?.phase === 'prejoin');
		await el.updateComplete;
		expect(q<HTMLElement>(el.shadowRoot, 'office-ui-prejoin').getAttribute('heading')).toBe(
			'Stand-up',
		);
		el.remove();
	});

	it('downloads an .ics for a meeting', async () => {
		const el = await mount();
		const start = Date.now() + 1_800_000;
		const meeting = el.client!.createMeeting({
			title: 'Export me',
			start,
			end: start + 1_800_000,
		})!;
		const urls: Blob[] = [];
		const create = URL.createObjectURL;
		URL.createObjectURL = ((blob: Blob) => (
			urls.push(blob),
			'blob:test'
		)) as typeof URL.createObjectURL;
		const click = HTMLAnchorElement.prototype.click;
		let downloaded = '';
		HTMLAnchorElement.prototype.click = function (this: HTMLAnchorElement) {
			downloaded = this.download;
		};
		try {
			select(el, 'calendar');
			await el.updateComplete;
			q<HTMLElement>(el.shadowRoot, 'teams-meeting-details').dispatchEvent(
				new CustomEvent('teams-meeting-ics', { detail: { meetingId: meeting.id } }),
			);
			expect(downloaded).toBe('Export-me.ics');
			expect(await urls[0]!.text()).toContain('SUMMARY:Export me');
		} finally {
			URL.createObjectURL = create;
			HTMLAnchorElement.prototype.click = click;
		}
		el.remove();
	});
});
