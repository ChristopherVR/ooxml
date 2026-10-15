import { type MeetingView, zonedToUtc } from 'ooxml-core/teams';
import { registerOfficeUi } from '../index';
import type { OfficeUiCalendarView } from './calendar-view';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

const ZONE = 'America/New_York';
const NOW = zonedToUtc({ year: 2026, month: 10, day: 15, hour: 11 }, ZONE); // a Thursday
const at = (day: number, hour: number, minute = 0): number =>
	zonedToUtc({ year: 2026, month: 10, day, hour, minute }, ZONE);
const meeting = (
	over: Partial<MeetingView> & { id: string; start: number; end: number },
): MeetingView => ({
	title: over.id,
	description: '',
	organizerId: 'ada',
	organizerName: 'Ada',
	zone: ZONE,
	repeat: 'none',
	attendees: [],
	createdAt: 0,
	cancelled: false,
	live: false,
	mine: true,
	...over,
});

async function mount(meetings: MeetingView[]): Promise<OfficeUiCalendarView> {
	const el = document.createElement('office-ui-calendar-view') as OfficeUiCalendarView;
	el.zone = ZONE;
	el.now = NOW;
	el.meetings = meetings;
	document.body.append(el);
	await el.updateComplete;
	return el;
}
const qa = <T extends Element>(el: HTMLElement, s: string): T[] => [
	...el.shadowRoot!.querySelectorAll<T>(s),
];

describe('office-ui-calendar-view', () => {
	it('shows the work week of the anchor day with today marked', async () => {
		const el = await mount([]);
		expect(qa(el, '.head .dom').map((d) => d.textContent)).toEqual(['12', '13', '14', '15', '16']);
		expect(qa<HTMLElement>(el, '.head').map((h) => h.dataset.today)).toEqual([
			'false',
			'false',
			'false',
			'true',
			'false',
		]);
		expect(el.shadowRoot!.querySelector('.range')!.textContent).toMatch(/^October 12\D+16, 2026$/u);
		expect(qa(el, '.column')).toHaveLength(5);
		expect(el.shadowRoot!.querySelector('.now')).not.toBeNull();
	});

	it('switches between work week, week and day, and steps with previous and next', async () => {
		const el = await mount([]);
		const modes = qa<HTMLButtonElement>(el, '.modes button');
		modes[1]!.click();
		await el.updateComplete;
		expect(qa(el, '.column')).toHaveLength(7);
		modes[2]!.click();
		await el.updateComplete;
		expect(qa(el, '.column')).toHaveLength(1);
		expect(el.shadowRoot!.querySelector('.range')!.textContent).toBe('Thursday, October 15, 2026');
		el.shadowRoot!.querySelector<HTMLButtonElement>('button[aria-label="Next day"]')!.click();
		await el.updateComplete;
		expect(el.shadowRoot!.querySelector('.head .dom')!.textContent).toBe('16');
		el.shadowRoot!.querySelector<HTMLButtonElement>('.today')!.click();
		await el.updateComplete;
		expect(el.shadowRoot!.querySelector('.head .dom')!.textContent).toBe('15');
		modes[0]!.click();
		await el.updateComplete;
		el.shadowRoot!.querySelector<HTMLButtonElement>('button[aria-label="Next week"]')!.click();
		await el.updateComplete;
		expect(qa(el, '.head .dom').map((d) => d.textContent)).toEqual(['19', '20', '21', '22', '23']);
	});

	it('places meetings in their day and lets overlapping ones share the column', async () => {
		const el = await mount([
			meeting({ id: 'a', start: at(14, 9), end: at(14, 10) }),
			meeting({ id: 'b', start: at(14, 9, 30), end: at(14, 10, 30) }),
			meeting({ id: 'c', start: at(16, 13), end: at(16, 14), live: true }),
		]);
		const columns = qa<HTMLElement>(el, '.column');
		const wednesday = [...columns[2]!.querySelectorAll<HTMLElement>('.event')];
		expect(wednesday).toHaveLength(2);
		expect(wednesday.map((e) => e.style.getPropertyValue('--span'))).toEqual(['2', '2']);
		expect(wednesday.map((e) => e.style.getPropertyValue('--col'))).toEqual(['0', '1']);
		expect(wednesday[0]!.style.getPropertyValue('--top')).toBe(String((9 * 60 * 100) / 1440));
		const friday = columns[4]!.querySelector<HTMLElement>('.event')!;
		expect(friday.dataset.live).toBe('true');
		expect(friday.getAttribute('aria-label')).toContain('c, 1:00');
	});

	it('repeats a daily meeting in every visible day', async () => {
		const el = await mount([
			meeting({ id: 'daily', start: at(12, 8), end: at(12, 8, 30), repeat: 'daily' }),
		]);
		expect(qa(el, '.event')).toHaveLength(5);
	});

	it('emits open for a meeting and new for an empty slot or the button', async () => {
		const el = await mount([meeting({ id: 'a', start: at(14, 9), end: at(14, 10) })]);
		const got: { type: string; detail: unknown }[] = [];
		for (const type of ['office-calendar-open', 'office-calendar-new'])
			el.addEventListener(type, (e) => got.push({ type, detail: (e as CustomEvent).detail }));
		qa<HTMLButtonElement>(el, '.event')[0]!.click();
		el.shadowRoot!.querySelector<HTMLButtonElement>('.new')!.click();
		expect(got).toEqual([
			{ type: 'office-calendar-open', detail: { meetingId: 'a', start: at(14, 9) } },
			{ type: 'office-calendar-new', detail: { start: 0 } },
		]);
	});
});
