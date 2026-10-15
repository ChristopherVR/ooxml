import type { MeetingView } from 'ooxml-core/teams';
import { defineTeamsMeetingDetails, type TeamsMeetingDetails } from './meeting-details';
import { defineTeamsMeetingDialog, type TeamsMeetingDialog } from './meeting-dialog';
import { parseRepeatValue, repeatOptions, repeatText, repeatValue } from './repeat-text';

const HOUR = 3_600_000;
const ZONE = Intl.DateTimeFormat().resolvedOptions().timeZone;
/** Local wall time, so the test does not depend on the machine's zone. */
const local = (month: number, day: number, hour = 9): number =>
	new Date(2026, month - 1, day, hour).getTime();

beforeAll(() => {
	defineTeamsMeetingDialog();
	defineTeamsMeetingDetails();
});
afterEach(() => document.body.replaceChildren());

async function dialog(start: number, save = vi.fn(() => true)) {
	const el = document.createElement('teams-meeting-dialog') as TeamsMeetingDialog;
	Object.assign(el, { meeting: null, start, save });
	document.body.append(el);
	el.open = true;
	await el.updateComplete;
	return { el, save };
}
const select = (el: TeamsMeetingDialog): HTMLSelectElement =>
	Array.from(el.shadowRoot!.querySelectorAll('label'))
		.find((l) => l.textContent?.trim().startsWith('Repeat'))!
		.querySelector('select')!;
const values = (el: TeamsMeetingDialog): string[] =>
	Array.from(select(el).options).map((o) => o.value);
const pick = async (el: TeamsMeetingDialog, value: string): Promise<void> => {
	const s = select(el);
	s.value = value;
	s.dispatchEvent(new Event('change'));
	await el.updateComplete;
};
const titled = async (el: TeamsMeetingDialog): Promise<void> => {
	const title = el.shadowRoot!.querySelector<HTMLInputElement>('input[maxlength]')!;
	title.value = 'Review';
	title.dispatchEvent(new Event('input'));
	await el.updateComplete;
};
const submit = async (el: TeamsMeetingDialog): Promise<void> => {
	el.shadowRoot!.querySelector('form')!.dispatchEvent(
		new Event('submit', { cancelable: true, bubbles: true }),
	);
	await el.updateComplete;
};

describe('repeat options', () => {
	it('names the date, the weekday rule and the year for the start', () => {
		const options = repeatOptions(local(10, 20), ZONE);
		expect(options.map((o) => o.label)).toEqual([
			'Does not repeat',
			'Every day',
			'Every week',
			'Every month on day 20',
			'Every month on the third Tuesday',
			'Every year on October 20',
		]);
	});

	it('offers both "fourth" and "last" when a date is both', () => {
		// 27 November 2026 is the fourth and the last Friday.
		expect(repeatOptions(local(11, 27), ZONE).map((o) => o.value)).toContain('monthly-4');
		expect(repeatOptions(local(11, 27), ZONE).map((o) => o.value)).toContain('monthly--1');
		// 30 October 2026 is only the last Friday.
		const only = repeatOptions(local(10, 30), ZONE).map((o) => o.value);
		expect(only).toContain('monthly--1');
		expect(only).not.toContain('monthly-5');
	});

	it('round-trips a value', () => {
		expect(parseRepeatValue('monthly--1')).toEqual({
			repeat: 'monthly',
			monthlyBy: 'weekday',
			ordinal: -1,
		});
		expect(repeatValue(parseRepeatValue('monthly-3'))).toBe('monthly-3');
		expect(repeatValue(parseRepeatValue('monthly-date'))).toBe('monthly-date');
		expect(repeatValue(parseRepeatValue('yearly'))).toBe('yearly');
		expect(parseRepeatValue('junk').repeat).toBe('none');
	});

	it('describes a stored repeat in words', () => {
		const base = { start: local(10, 30), zone: ZONE };
		expect(repeatText({ ...base, repeat: 'monthly' })).toBe('Repeats every month on day 30');
		expect(repeatText({ ...base, repeat: 'monthly', monthlyBy: 'weekday', ordinal: -1 })).toBe(
			'Repeats every month on the last Friday',
		);
		expect(repeatText({ ...base, repeat: 'yearly' })).toBe('Repeats every year on October 30');
		expect(repeatText({ ...base, repeat: 'weekly' })).toBe('Repeats every week');
	});
});

describe('meeting dialog: monthly and yearly', () => {
	it('lists the choices for the opened date', async () => {
		const { el } = await dialog(local(10, 20));
		expect(values(el)).toEqual(['none', 'daily', 'weekly', 'monthly-date', 'monthly-3', 'yearly']);
	});

	it('prints option labels without stray whitespace', async () => {
		const { el } = await dialog(local(10, 20));
		const labels = Array.from(select(el).options).map((o) => o.textContent);
		expect(labels).toContain('Every month on the third Tuesday');
		expect(labels).toContain('Every year on October 20');
	});

	it('saves a monthly date repeat', async () => {
		const { el, save } = await dialog(local(10, 20));
		await titled(el);
		await pick(el, 'monthly-date');
		await submit(el);
		expect(save).toHaveBeenLastCalledWith(
			expect.objectContaining({ repeat: 'monthly', monthlyBy: 'date' }),
			undefined,
		);
		expect(save.mock.calls.at(-1)).not.toHaveProperty('0.ordinal');
	});

	it('saves "the third Tuesday" with its ordinal', async () => {
		const { el, save } = await dialog(local(10, 20));
		await titled(el);
		await pick(el, 'monthly-3');
		await submit(el);
		expect(save).toHaveBeenLastCalledWith(
			expect.objectContaining({ repeat: 'monthly', monthlyBy: 'weekday', ordinal: 3 }),
			undefined,
		);
	});

	it('saves "the last Friday" and a yearly repeat', async () => {
		const { el, save } = await dialog(local(10, 30));
		await titled(el);
		await pick(el, 'monthly--1');
		await submit(el);
		expect(save).toHaveBeenLastCalledWith(
			expect.objectContaining({ monthlyBy: 'weekday', ordinal: -1 }),
			undefined,
		);
		await pick(el, 'yearly');
		await submit(el);
		const last = save.mock.calls.at(-1) as unknown as [Record<string, unknown>];
		expect(last[0]).toMatchObject({ repeat: 'yearly' });
		expect(last[0]).not.toHaveProperty('monthlyBy');
	});

	it('moves a weekday rule to an ordinal that fits when the start moves', async () => {
		const { el, save } = await dialog(local(10, 20));
		await titled(el);
		await pick(el, 'monthly-3');
		// Move to the second Tuesday (13 October): the third no longer fits.
		const start = el.shadowRoot!.querySelector<HTMLInputElement>('input[type=datetime-local]')!;
		const text = (d: Date): string =>
			`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}T09:00`;
		start.value = text(new Date(2026, 9, 13));
		start.dispatchEvent(new Event('change'));
		await el.updateComplete;
		expect(select(el).value).toBe('monthly-2');
		await submit(el);
		expect(save).toHaveBeenLastCalledWith(
			expect.objectContaining({ monthlyBy: 'weekday', ordinal: 2 }),
			undefined,
		);
	});

	it('shows an existing monthly weekday meeting as selected', async () => {
		const meeting: MeetingView = {
			id: 'm1',
			title: 'Review',
			description: '',
			organizerId: 'ada',
			organizerName: 'Ada',
			zone: ZONE,
			start: local(10, 20),
			end: local(10, 20) + HOUR,
			repeat: 'monthly',
			monthlyBy: 'weekday',
			ordinal: 3,
			attendees: [],
			createdAt: 0,
			cancelled: false,
			live: false,
			mine: true,
		};
		const el = document.createElement('teams-meeting-dialog') as TeamsMeetingDialog;
		Object.assign(el, { meeting });
		document.body.append(el);
		el.open = true;
		await el.updateComplete;
		expect(select(el).querySelector<HTMLOptionElement>('option[selected]')?.value).toBe(
			'monthly-3',
		);
	});
});

describe('meeting details: repeat sentence', () => {
	it('says how a monthly meeting repeats', async () => {
		const el = document.createElement('teams-meeting-details') as TeamsMeetingDetails;
		Object.assign(el, {
			zone: ZONE,
			occurrenceStart: local(10, 20),
			meeting: {
				id: 'm1',
				title: 'Review',
				description: '',
				organizerId: 'ada',
				organizerName: 'Ada',
				zone: ZONE,
				start: local(10, 20),
				end: local(10, 20) + HOUR,
				repeat: 'monthly',
				monthlyBy: 'weekday',
				ordinal: 3,
				count: 6,
				attendees: [],
				createdAt: 0,
				cancelled: false,
				live: false,
				mine: true,
			} satisfies MeetingView,
		});
		document.body.append(el);
		el.open = true;
		await el.updateComplete;
		expect(el.shadowRoot!.textContent).toContain(
			'Repeats every month on the third Tuesday, 6 times',
		);
	});
});
