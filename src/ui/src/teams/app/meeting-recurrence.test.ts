import { type MeetingView, zonedToUtc } from 'ooxml-core/teams';
import { defineTeamsMeetingDetails, type TeamsMeetingDetails } from './meeting-details';
import { defineTeamsMeetingDialog, type TeamsMeetingDialog } from './meeting-dialog';
import { defineTeamsApp, type TeamsApp } from './teams-app';

const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));
const HOUR = 3_600_000;
const NY = 'America/New_York';
const at = (month: number, day: number, hour = 9): number =>
	zonedToUtc({ year: 2026, month, day, hour }, NY);

const view = (over: Partial<MeetingView> = {}): MeetingView => ({
	id: 'm1',
	title: 'Stand-up',
	description: '',
	organizerId: 'ada',
	organizerName: 'Ada',
	zone: NY,
	start: at(3, 2),
	end: at(3, 2) + HOUR,
	repeat: 'daily',
	attendees: [],
	createdAt: 0,
	cancelled: false,
	live: false,
	mine: true,
	...over,
});

beforeAll(() => {
	defineTeamsMeetingDialog();
	defineTeamsMeetingDetails();
	defineTeamsApp();
});
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

async function dialog(
	meeting: MeetingView | null,
	extra: Partial<TeamsMeetingDialog> = {},
): Promise<TeamsMeetingDialog> {
	const el = document.createElement('teams-meeting-dialog') as TeamsMeetingDialog;
	Object.assign(el, { meeting, ...extra });
	document.body.append(el);
	el.open = true;
	await el.updateComplete;
	return el;
}
const field = <T extends Element>(el: Element, selector: string): T =>
	el.shadowRoot!.querySelector<T>(selector)!;
const labelled = <T extends Element>(el: Element, text: string, selector: string): T => {
	const label = Array.from(el.shadowRoot!.querySelectorAll('label')).find((l) =>
		l.textContent?.trim().startsWith(text),
	);
	return label!.querySelector<T>(selector)!;
};
const choose = async (el: TeamsMeetingDialog, text: string, value: string): Promise<void> => {
	const select = labelled<HTMLSelectElement>(el, text, 'select');
	select.value = value;
	select.dispatchEvent(new Event('change'));
	await el.updateComplete;
};
const submit = async (el: TeamsMeetingDialog): Promise<void> => {
	field<HTMLFormElement>(el, 'form').dispatchEvent(
		new Event('submit', { cancelable: true, bubbles: true }),
	);
	await el.updateComplete;
};

describe('meeting dialog: recurrence end', () => {
	it('shows the end controls only for a repeating meeting', async () => {
		const el = await dialog(null, { start: at(3, 2) });
		expect(el.shadowRoot!.textContent).not.toContain('Ends');
		await choose(el, 'Repeat', 'weekly');
		expect(el.shadowRoot!.textContent).toContain('Ends');
		expect(field(el, 'input[type=date]')).toBeNull();
		await choose(el, 'Ends', 'date');
		expect(field(el, 'input[type=date]')).not.toBeNull();
		await choose(el, 'Ends', 'count');
		expect(field<HTMLInputElement>(el, 'input[type=number]').value).toBe('10');
		el.remove();
	});

	it('saves a count, a date or no end for a repeat', async () => {
		const save = vi.fn(() => true);
		const el = await dialog(null, { start: at(3, 2), save });
		labelled<HTMLInputElement>(el, 'Title', 'input').value = 'Review';
		labelled<HTMLInputElement>(el, 'Title', 'input').dispatchEvent(new Event('input'));
		await choose(el, 'Repeat', 'daily');
		await choose(el, 'Ends', 'count');
		const count = labelled<HTMLInputElement>(el, 'Number of events', 'input');
		count.value = '4';
		count.dispatchEvent(new Event('input'));
		await submit(el);
		expect(save).toHaveBeenLastCalledWith(
			expect.objectContaining({ repeat: 'daily', count: 4, until: null }),
			undefined,
		);

		await choose(el, 'Ends', 'date');
		const date = labelled<HTMLInputElement>(el, 'End date', 'input');
		const start = new Date(el.startText).getTime();
		const day = new Date(start + 7 * 86_400_000);
		date.value = `${day.getFullYear()}-${String(day.getMonth() + 1).padStart(2, '0')}-${String(day.getDate()).padStart(2, '0')}`;
		date.dispatchEvent(new Event('change'));
		await submit(el);
		const last = save.mock.calls.at(-1) as unknown as [{ until: number; count: null }];
		expect(last[0].count).toBeNull();
		expect(new Date(last[0].until).getHours()).toBe(23);
		expect(last[0].until).toBeGreaterThanOrEqual(start + 7 * 86_400_000);

		await choose(el, 'Ends', 'never');
		await submit(el);
		expect(save).toHaveBeenLastCalledWith(
			expect.objectContaining({ count: null, until: null }),
			undefined,
		);
		el.remove();
	});

	it('refuses a bad count and an end before the first event', async () => {
		const save = vi.fn(() => true);
		const el = await dialog(null, { start: at(3, 2), save });
		labelled<HTMLInputElement>(el, 'Title', 'input').value = 'Review';
		labelled<HTMLInputElement>(el, 'Title', 'input').dispatchEvent(new Event('input'));
		await choose(el, 'Repeat', 'weekly');
		await choose(el, 'Ends', 'count');
		const count = labelled<HTMLInputElement>(el, 'Number of events', 'input');
		count.value = '0';
		count.dispatchEvent(new Event('input'));
		await submit(el);
		expect(field(el, '.issue').textContent).toContain('from 1 to 999');
		await choose(el, 'Ends', 'date');
		const date = labelled<HTMLInputElement>(el, 'End date', 'input');
		date.value = '2001-01-01';
		date.dispatchEvent(new Event('change'));
		await submit(el);
		expect(field(el, '.issue').textContent).toContain('on or after its first event');
		expect(save).not.toHaveBeenCalled();
		el.remove();
	});

	it('shows the saved end when editing a series', async () => {
		const counted = await dialog(view({ count: 7 }));
		expect(labelled<HTMLSelectElement>(counted, 'Ends', 'select').value).toBe('count');
		expect(labelled<HTMLInputElement>(counted, 'Number of events', 'input').value).toBe('7');
		counted.remove();
		const dated = await dialog(view({ until: at(3, 20) }));
		expect(labelled<HTMLSelectElement>(dated, 'Ends', 'select').value).toBe('date');
		dated.remove();
	});
});

describe('meeting dialog: one occurrence', () => {
	it('edits only the time, title and details of the occurrence that was opened', async () => {
		const save = vi.fn(() => true);
		const m = view({
			exceptions: [
				{
					originalStart: at(3, 3),
					cancelled: false,
					title: 'Moved',
					start: at(3, 3, 14),
					end: at(3, 3, 15),
				},
			],
		});
		const el = await dialog(m, { occurrence: at(3, 3), start: at(3, 3, 14), save });
		expect(el.shadowRoot!.textContent).toContain('Edit this event');
		expect(el.shadowRoot!.textContent).not.toContain('Repeat');
		expect(el.shadowRoot!.textContent).not.toContain('Attendees');
		expect(labelled<HTMLInputElement>(el, 'Title', 'input').value).toBe('Moved');
		labelled<HTMLInputElement>(el, 'Title', 'input').value = 'Moved again';
		labelled<HTMLInputElement>(el, 'Title', 'input').dispatchEvent(new Event('input'));
		await submit(el);
		const call = save.mock.calls[0] as unknown as [Record<string, unknown>, string, number];
		expect(call[0]).toMatchObject({ title: 'Moved again', start: at(3, 3, 14), end: at(3, 3, 15) });
		expect(call[0]).not.toHaveProperty('repeat');
		expect(call[1]).toBe('m1');
		expect(call[2]).toBe(at(3, 3));
		el.remove();
	});
});

describe('meeting details: edit this event or the series', () => {
	async function details(m: MeetingView, occurrenceStart = 0): Promise<TeamsMeetingDetails> {
		const el = document.createElement('teams-meeting-details') as TeamsMeetingDetails;
		Object.assign(el, { meeting: m, occurrenceStart, zone: NY });
		document.body.append(el);
		el.open = true;
		await el.updateComplete;
		return el;
	}
	const button = (el: Element, text: string): HTMLButtonElement =>
		Array.from(el.shadowRoot!.querySelectorAll('button')).find(
			(b) => b.textContent?.trim() === text,
		)!;
	const heard = (el: Element, name: string): unknown[] => {
		const out: unknown[] = [];
		el.addEventListener(name, (e) => out.push((e as CustomEvent).detail));
		return out;
	};

	it('asks which events an edit is for, defaulting to this event', async () => {
		const el = await details(view(), at(3, 4));
		const edits = heard(el, 'teams-meeting-edit');
		button(el, 'Edit').click();
		await el.updateComplete;
		expect(el.shadowRoot!.textContent).toContain('All events in the series');
		expect(edits).toEqual([]);
		field<HTMLFormElement>(el, 'form.scope').dispatchEvent(
			new Event('submit', { cancelable: true }),
		);
		expect(edits).toEqual([{ meetingId: 'm1', scope: 'occurrence', recurrenceId: at(3, 4) }]);
		el.remove();
	});

	it('edits the series when that choice is made, and can go back', async () => {
		const el = await details(view(), at(3, 4));
		const edits = heard(el, 'teams-meeting-edit');
		button(el, 'Edit').click();
		await el.updateComplete;
		button(el, 'Back').click();
		await el.updateComplete;
		expect(field(el, 'form.scope')).toBeNull();
		button(el, 'Edit').click();
		await el.updateComplete;
		const series = el.shadowRoot!.querySelectorAll<HTMLInputElement>('input[name=scope]')[1]!;
		series.dispatchEvent(new Event('change'));
		await el.updateComplete;
		field<HTMLFormElement>(el, 'form.scope').dispatchEvent(
			new Event('submit', { cancelable: true }),
		);
		expect(edits).toEqual([{ meetingId: 'm1', scope: 'series', recurrenceId: at(3, 4) }]);
		el.remove();
	});

	it('cancels one event or the series, and skips the question for a single meeting', async () => {
		const el = await details(view(), at(3, 4));
		const cancels = heard(el, 'teams-meeting-cancel');
		button(el, 'Cancel meeting').click();
		await el.updateComplete;
		field<HTMLFormElement>(el, 'form.scope').dispatchEvent(
			new Event('submit', { cancelable: true }),
		);
		expect(cancels).toEqual([{ meetingId: 'm1', scope: 'occurrence', recurrenceId: at(3, 4) }]);
		el.remove();

		const single = await details(view({ repeat: 'none' }));
		const edits = heard(single, 'teams-meeting-edit');
		button(single, 'Edit').click();
		expect(edits).toEqual([{ meetingId: 'm1', scope: 'series', recurrenceId: at(3, 2) }]);
		single.remove();
	});

	it('shows the end of the repeat and the occurrence overrides', async () => {
		const m = view({
			count: 3,
			exceptions: [
				{ originalStart: at(3, 3), cancelled: false, title: 'Moved', description: 'Room 4' },
			],
		});
		const el = await details(m, at(3, 3));
		expect(el.shadowRoot!.textContent).toContain('Repeats every day, 3 times');
		expect(el.shadowRoot!.textContent).toContain('Moved');
		expect(el.shadowRoot!.textContent).toContain('Room 4');
		el.remove();
	});
});

describe('<teams-app> recurring meeting changes', () => {
	let n = 0;
	async function mount(): Promise<TeamsApp> {
		const el = document.createElement('teams-app') as TeamsApp;
		el.setAttribute('workspace-id', `recur-${Date.now()}-${n++}`);
		el.setAttribute('user-name', 'Ada');
		el.setAttribute('user-id', 'ada');
		el.setAttribute('server-config', JSON.stringify({ mode: 'local', iceServers: [] }));
		document.body.append(el);
		await el.updateComplete;
		return el;
	}

	it('cancels one occurrence and edits another without touching the series', async () => {
		const el = await mount();
		const start = Math.ceil((Date.now() + 2 * 86_400_000) / HOUR) * HOUR;
		const meeting = el.client!.createMeeting({
			title: 'Daily',
			start,
			end: start + HOUR,
			repeat: 'daily',
			count: 5,
			zone: 'UTC',
		})!;
		const second = start + 86_400_000;
		const third = start + 2 * 86_400_000;
		el.shadowRoot!.querySelector('office-ui-app-rail')!.dispatchEvent(
			new CustomEvent('office-rail-select', { detail: { id: 'calendar' } }),
		);
		await el.updateComplete;
		const details = el.shadowRoot!.querySelector('teams-meeting-details')!;
		el.shadowRoot!.querySelector('office-ui-calendar-agenda')!.dispatchEvent(
			new CustomEvent('office-agenda-open', { detail: { meetingId: meeting.id, start: second } }),
		);
		await el.updateComplete;
		details.dispatchEvent(
			new CustomEvent('teams-meeting-cancel', {
				detail: { meetingId: meeting.id, scope: 'occurrence', recurrenceId: second },
			}),
		);
		await tick(10);
		expect(el.client!.getState().meetings[0]!.exceptions).toEqual([
			{ originalStart: second, cancelled: true },
		]);
		expect(el.client!.getState().meetings[0]!.repeat).toBe('daily');

		details.dispatchEvent(
			new CustomEvent('teams-meeting-edit', {
				detail: { meetingId: meeting.id, scope: 'occurrence', recurrenceId: third },
			}),
		);
		await el.updateComplete;
		const form = el.shadowRoot!.querySelector('teams-meeting-dialog') as TeamsMeetingDialog;
		expect(form.occurrence).toBe(third);
		expect(
			form.save!({ title: 'Third, renamed', start: third, end: third + HOUR }, meeting.id, third),
		).toBe(true);
		await tick(10);
		const saved = el.client!.getState().meetings[0]!;
		expect(saved.title).toBe('Daily');
		expect(saved.exceptions).toHaveLength(2);
		expect(saved.exceptions?.[1]).toMatchObject({ originalStart: third, title: 'Third, renamed' });
		el.remove();
	});
});
