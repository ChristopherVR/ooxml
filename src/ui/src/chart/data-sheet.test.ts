import { registerOfficeUi } from '../index';
import type { OfficeChartDataChange, OfficeUiChartDataSheet } from './data-sheet';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

function mount(): OfficeUiChartDataSheet {
	const sheet = document.createElement('office-ui-chart-data-sheet') as OfficeUiChartDataSheet;
	sheet.data = {
		categories: ['Q1', 'Q2'],
		series: [
			{ name: 'North', values: [1, 2.5] },
			{ name: 'South', values: [3, null] },
		],
	};
	document.body.append(sheet);
	return sheet;
}

const cell = (sheet: HTMLElement, row: number, col: number) =>
	sheet.shadowRoot!.querySelector<HTMLInputElement>(`input[data-row="${row}"][data-col="${col}"]`)!;
const changes = (sheet: HTMLElement) => {
	const seen: OfficeChartDataChange[] = [];
	sheet.addEventListener('office-chart-data-change', (event) =>
		seen.push((event as CustomEvent<OfficeChartDataChange>).detail),
	);
	return seen;
};
const type = (input: HTMLInputElement, value: string) => {
	input.value = value;
	input.dispatchEvent(new Event('change', { bubbles: true }));
};
const key = (input: HTMLInputElement, name: string) =>
	input.dispatchEvent(
		new KeyboardEvent('keydown', { key: name, bubbles: true, composed: true, cancelable: true }),
	);

describe('office-ui-chart-data-sheet', () => {
	it('draws series across, categories down, with accessible names', () => {
		const sheet = mount();
		expect(cell(sheet, 0, 1).value).toBe('North');
		expect(cell(sheet, 2, 0).value).toBe('Q2');
		expect(cell(sheet, 2, 1).value).toBe('2.5');
		expect(cell(sheet, 2, 2).value).toBe('');
		expect(cell(sheet, 1, 2).getAttribute('aria-label')).toBe('South value 1');
		expect(sheet.shadowRoot!.querySelector('.summary')!.textContent).toBe(
			'2 series · 2 categories',
		);
	});

	it('emits one edit per committed cell and refuses text that is not a number', () => {
		const sheet = mount();
		const seen = changes(sheet);
		type(cell(sheet, 1, 1), '42');
		type(cell(sheet, 0, 2), 'West');
		type(cell(sheet, 1, 0), 'Jan');
		expect(seen.map((change) => change.edits[0])).toEqual([
			{ kind: 'value', series: 0, category: 0, value: 42 },
			{ kind: 'series-name', series: 1, name: 'West' },
			{ kind: 'category-label', category: 0, label: 'Jan' },
		]);
		expect(seen[2]!.data.series[1]!.name).toBe('West');
		type(cell(sheet, 2, 1), 'abc');
		expect(seen).toHaveLength(3);
		expect(cell(sheet, 2, 1).value).toBe('2.5');
		expect(sheet.shadowRoot!.querySelector('.status')!.textContent).toBe('Enter a number');
	});

	it('reads and writes numbers in its locale', () => {
		const sheet = mount();
		sheet.locale = 'de-DE';
		const seen = changes(sheet);
		expect(cell(sheet, 2, 1).value).toBe('2,5');
		type(cell(sheet, 1, 1), '1.234,5');
		expect(seen[0]!.edits[0]).toMatchObject({ value: 1234.5 });
	});

	it('adds and removes series and categories, keeping at least one of each', () => {
		const sheet = mount();
		const seen = changes(sheet);
		sheet.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="add-series"]')!.click();
		sheet.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="add-category"]')!.click();
		expect(sheet.data.series.map((entry) => entry.name)).toEqual(['North', 'South', 'Series 3']);
		expect(sheet.data.categories).toEqual(['Q1', 'Q2', 'Cat 3']);
		const removes = () => sheet.shadowRoot!.querySelectorAll<HTMLButtonElement>('button.remove');
		expect(removes()).toHaveLength(6);
		removes()[0]!.click();
		expect(sheet.data.series.map((entry) => entry.name)).toEqual(['South', 'Series 3']);
		sheet.data = { categories: ['A'], series: [{ name: 'S', values: [1] }] };
		expect(removes()).toHaveLength(0);
		expect(seen.map((change) => change.edits[0]!.kind)).toEqual([
			'add-series',
			'add-category',
			'remove-series',
		]);
	});

	it('moves with the keyboard, commits on Enter and keeps keys from the editor around it', () => {
		const sheet = mount();
		const seen = changes(sheet);
		const outside = vi.fn();
		document.body.addEventListener('keydown', outside);
		const first = cell(sheet, 1, 1);
		first.focus();
		key(first, '9');
		first.value = '9';
		key(first, 'Enter');
		expect(seen[0]!.edits[0]).toEqual({ kind: 'value', series: 0, category: 0, value: 9 });
		expect(sheet.shadowRoot!.activeElement).toBe(cell(sheet, 2, 1));
		key(cell(sheet, 2, 1), 'ArrowRight');
		expect(sheet.shadowRoot!.activeElement).toBe(cell(sheet, 2, 2));
		const typed = cell(sheet, 2, 2);
		typed.value = '77';
		key(typed, 'Escape');
		expect(typed.value).toBe('');
		expect(outside).not.toHaveBeenCalled();
		document.body.removeEventListener('keydown', outside);
	});

	it('pastes a tab-separated block from the focused cell, growing the sheet', () => {
		const sheet = mount();
		const seen = changes(sheet);
		const paste = new Event('paste', { bubbles: true, cancelable: true }) as ClipboardEvent;
		Object.defineProperty(paste, 'clipboardData', {
			value: { getData: () => '5\t6\t7\r\n8\t9\t10\r\n' },
		});
		cell(sheet, 2, 1).dispatchEvent(paste);
		expect(paste.defaultPrevented).toBe(true);
		expect(seen).toHaveLength(1);
		expect(sheet.data.categories).toHaveLength(3);
		expect(sheet.data.series.map((entry) => entry.values)).toEqual([
			[1, 5, 8],
			[3, 6, 9],
			[0, 7, 10],
		]);
	});

	it('marks the highlighted cell, and a disabled sheet edits nothing', () => {
		const sheet = mount();
		sheet.highlight = { series: 1, point: 0 };
		expect(cell(sheet, 1, 2).classList.contains('highlight')).toBe(true);
		sheet.highlight = { series: 0 };
		expect(cell(sheet, 0, 1).classList.contains('highlight')).toBe(true);
		sheet.disabled = true;
		const seen = changes(sheet);
		expect(sheet.shadowRoot!.querySelector('.actions')).toBeNull();
		expect(cell(sheet, 1, 1).disabled).toBe(true);
		type(cell(sheet, 1, 1), '5');
		expect(seen).toHaveLength(0);
	});

	it('takes translated labels', () => {
		const sheet = mount();
		sheet.labels = { heading: 'Daten', valueCell: '{name} Wert {number}' };
		expect(sheet.shadowRoot!.querySelector('.heading')!.textContent).toBe('Daten');
		expect(cell(sheet, 1, 1).getAttribute('aria-label')).toBe('North Wert 1');
	});
});
