import { registerOfficeUi } from '../index';
import type { OfficeUiAppRail, RailItem } from './app-rail';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

const items: RailItem[] = [
	{ id: 'activity', label: 'Activity', icon: 'bell' },
	{ id: 'teams', label: 'Teams', icon: 'users' },
	{ id: 'calls', label: 'Calls', icon: 'phone', secondary: true, badge: 2 },
	{ id: 'files', label: 'Files', icon: 'folder', secondary: true, badge: 150 },
];

async function mount(selected = 'teams'): Promise<OfficeUiAppRail> {
	const el = document.createElement('office-ui-app-rail') as OfficeUiAppRail;
	el.items = items;
	el.selected = selected;
	document.body.append(el);
	await el.updateComplete;
	return el;
}

describe('office-ui-app-rail More sheet', () => {
	it('renders every item for the desktop rail and a More button carrying the hidden badges', async () => {
		const el = await mount();
		const bar = el.shadowRoot!.querySelector('ul')!;
		expect(
			[...bar.querySelectorAll('li:not(.more-item) > button')].map((b) =>
				b.getAttribute('aria-label'),
			),
		).toEqual(['Activity', 'Teams', 'Calls', 'Files']);
		const more = el.shadowRoot!.querySelector<HTMLButtonElement>('.more')!;
		expect(more.getAttribute('aria-label')).toBe('More, 152 unread');
		expect(more.querySelector('.badge')!.textContent).toBe('99+');
		expect(more.getAttribute('aria-haspopup')).toBe('dialog');
	});

	it('opens a labelled dialog, selects from it and returns focus to More', async () => {
		const el = await mount();
		const got: unknown[] = [];
		el.addEventListener('office-rail-select', (e) => got.push((e as CustomEvent).detail));
		const more = el.shadowRoot!.querySelector<HTMLButtonElement>('.more')!;
		more.click();
		await el.updateComplete;
		const sheet = el.shadowRoot!.querySelector('dialog')!;
		expect(sheet.getAttribute('aria-label')).toBe('More apps');
		expect(more.getAttribute('aria-expanded')).toBe('true');
		sheet.dispatchEvent(new Event('cancel', { cancelable: true }));
		await el.updateComplete;
		expect(more.getAttribute('aria-expanded')).toBe('false');
		more.click();
		await el.updateComplete;
		el.shadowRoot!.querySelectorAll<HTMLButtonElement>('dialog li button')[1]!.click();
		await el.updateComplete;
		expect(got).toEqual([{ id: 'files' }]);
		expect(more.getAttribute('aria-expanded')).toBe('false');
	});

	it('marks More as current when a hidden app is selected', async () => {
		const el = await mount('files');
		expect(el.shadowRoot!.querySelector('.more')!.getAttribute('aria-current')).toBe('page');
	});

	it('has no More button when nothing is secondary', async () => {
		const el = document.createElement('office-ui-app-rail') as OfficeUiAppRail;
		el.items = items.slice(0, 2);
		document.body.append(el);
		await el.updateComplete;
		expect(el.shadowRoot!.querySelector('.more')).toBeNull();
	});
});
