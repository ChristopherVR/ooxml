import { registerOfficeUi } from '../index';
import type { OfficeChartValue } from './chart-picker';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Picker = HTMLElement & {
	types: { id: string; label: string }[];
	groupings: { id: string; label: string }[];
	groupedTypes: string[];
	drawPreview: (value: OfficeChartValue) => string;
	value: OfficeChartValue;
	changeOnly: boolean;
};

const svg = (value: OfficeChartValue) =>
	`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><title>${value.type}:${value.grouping}:${value.title}</title><rect width="10" height="10"/></svg>`;

function mount(): Picker {
	const picker = document.createElement('office-ui-chart-picker') as Picker;
	picker.types = [
		{ id: 'column', label: 'Column' },
		{ id: 'line', label: 'Line' },
	];
	picker.groupings = [
		{ id: 'clustered', label: 'Clustered' },
		{ id: 'stacked', label: 'Stacked' },
	];
	picker.groupedTypes = ['column'];
	picker.drawPreview = svg;
	document.body.append(picker);
	return picker;
}

const root = (picker: Picker) => picker.shadowRoot!;
const changes = (picker: Picker) => {
	const seen: OfficeChartValue[] = [];
	picker.addEventListener('office-chart-change', (event) =>
		seen.push((event as CustomEvent).detail),
	);
	return seen;
};

describe('office-ui-chart-picker', () => {
	it('draws a preview for the chosen type and a tile for each type', () => {
		const picker = mount();
		expect(root(picker).querySelector('.preview title')!.textContent).toBe('column:clustered:');
		expect(root(picker).querySelectorAll('[data-gallery-item]')).toHaveLength(2);
		expect(root(picker).querySelectorAll('office-ui-gallery .tile svg')).toHaveLength(2);
	});

	it('draws type tiles at the preview size, flagged as tiles, so the gallery scales them down', () => {
		const picker = mount();
		const calls: [string, number, boolean][] = [];
		picker.drawPreview = ((value: OfficeChartValue, size: { width: number }, tile: boolean) => {
			calls.push([value.type, size.width, tile]);
			return svg(value);
		}) as never;
		expect(calls).toContainEqual(['line', 360, true]);
		expect(calls).toContainEqual(['column', 360, false]);
	});

	it('offers grouping only for the types that have it', () => {
		const picker = mount();
		expect(root(picker).querySelector('[data-key="grouping"]')).not.toBeNull();
		root(picker).querySelector<HTMLElement>('[data-gallery-item="line"]')!.click();
		expect(picker.value.type).toBe('line');
		expect(root(picker).querySelector('[data-key="grouping"]')).toBeNull();
		expect(root(picker).querySelector('.preview title')!.textContent).toBe('line:clustered:');
	});

	it('reports each edit with the whole value', () => {
		const picker = mount();
		const seen = changes(picker);
		const title = root(picker).querySelector<HTMLInputElement>('[data-key="title"]')!;
		title.value = 'Sales';
		title.dispatchEvent(new Event('input'));
		const grouping = root(picker).querySelector<HTMLSelectElement>('[data-key="grouping"]')!;
		grouping.value = 'stacked';
		grouping.dispatchEvent(new Event('change'));
		expect(seen).toEqual([
			{ type: 'column', grouping: 'clustered', title: 'Sales' },
			{ type: 'column', grouping: 'stacked', title: 'Sales' },
		]);
	});

	it('shows the data editor only when data is set', () => {
		const picker = mount();
		expect(root(picker).querySelector('[data-key="data"]')).toBeNull();
		picker.value = { ...picker.value, data: 'a\tb' };
		const data = root(picker).querySelector<HTMLTextAreaElement>('[data-key="data"]')!;
		expect(data.value).toBe('a\tb');
		const seen = changes(picker);
		data.value = 'a\tb\n1\t2';
		data.dispatchEvent(new Event('input'));
		expect(seen.at(-1)!.data).toBe('a\tb\n1\t2');
	});

	it('hides the title and data in change-only mode', () => {
		const picker = mount();
		picker.value = { ...picker.value, data: 'a' };
		picker.changeOnly = true;
		expect(root(picker).querySelector('[data-key="title"]')).toBeNull();
		expect(root(picker).querySelector('[data-key="data"]')).toBeNull();
	});

	it('survives a preview that throws', () => {
		const picker = mount();
		picker.drawPreview = () => {
			throw new Error('bad chart');
		};
		expect(root(picker).querySelector('.preview')!.childElementCount).toBe(0);
	});
});
