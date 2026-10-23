import { registerOfficeUi } from '../index';
import type {
	OfficeSmartArtLayout,
	OfficeSmartArtPreview,
	OfficeSmartArtValue,
} from './smartart-picker';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Picker = HTMLElement & {
	categories: { id: string; label: string; description?: string }[];
	layouts: OfficeSmartArtLayout[];
	labels: Record<string, string>;
	drawPreview: OfficeSmartArtPreview | undefined;
	value: OfficeSmartArtValue;
	category: string;
	itemsEditor: boolean;
};

function mount(): Picker {
	const picker = document.createElement('office-ui-smartart-picker') as Picker;
	picker.categories = [
		{ id: 'list', label: 'List', description: 'Grouped blocks.' },
		{ id: 'cycle', label: 'Cycle', description: 'A continual process.' },
	];
	picker.layouts = [
		{ id: 'blocks', category: 'list', label: 'Blocks', family: 'list', defaultItems: ['A', 'B'] },
		{ id: 'stacked', category: 'list', label: 'Stacked', defaultItems: ['X', 'Y', 'Z'] },
		{ id: 'ring', category: 'cycle', label: 'Ring', family: 'cycle', defaultItems: ['1', '2'] },
	];
	document.body.append(picker);
	return picker;
}

const root = (picker: Picker) => picker.shadowRoot!;
const options = (picker: Picker) => [
	...root(picker).querySelectorAll<HTMLElement>('[role="option"]'),
];
const events = (picker: Picker, name: string) => {
	const seen: OfficeSmartArtValue[] = [];
	picker.addEventListener(name, (event) => seen.push((event as CustomEvent).detail));
	return seen;
};

describe('office-ui-smartart-picker', () => {
	it('shows the first category with its first layout chosen, and family tiles', () => {
		const picker = mount();
		expect(picker.category).toBe('list');
		expect(picker.value).toEqual({ layout: 'blocks', items: ['A', 'B'] });
		expect(options(picker).map((option) => option.dataset.layout)).toEqual(['blocks', 'stacked']);
		expect(options(picker)[0]!.getAttribute('aria-selected')).toBe('true');
		expect(root(picker).querySelectorAll('.tile svg')).toHaveLength(2);
		expect(root(picker).querySelector('.title')!.textContent).toBe('Blocks');
		expect(root(picker).querySelector('.description')!.textContent).toBe('Grouped blocks.');
	});

	it('switches category, selecting its first layout, and reports the change', () => {
		const picker = mount();
		const seen = events(picker, 'office-smartart-change');
		root(picker).querySelector<HTMLElement>('[data-category="cycle"]')!.click();
		expect(picker.category).toBe('cycle');
		expect(seen).toEqual([{ layout: 'ring', items: ['1', '2'] }]);
		expect(
			root(picker).querySelector('[data-category="cycle"]')!.getAttribute('aria-pressed'),
		).toBe('true');
	});

	it('picks with a click, activates with a double click or Enter, and moves with arrows', () => {
		const picker = mount();
		const changes = events(picker, 'office-smartart-change');
		const activated = events(picker, 'office-smartart-activate');
		options(picker)[1]!.click();
		expect(changes.at(-1)).toEqual({ layout: 'stacked', items: ['X', 'Y', 'Z'] });
		options(picker)[0]!.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		expect(activated).toEqual([{ layout: 'blocks', items: ['A', 'B'] }]);
		options(picker)[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight' }));
		expect(picker.value.layout).toBe('stacked');
		expect(root(picker).activeElement).toBe(options(picker)[1]);
		options(picker)[1]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
		expect(activated.at(-1)!.layout).toBe('stacked');
	});

	it('edits the node texts and keeps typed text across a layout change', () => {
		const picker = mount();
		picker.itemsEditor = true;
		const box = root(picker).querySelector<HTMLTextAreaElement>('textarea')!;
		expect(box.value).toBe('A\nB');
		box.value = 'One\n\n Two \n';
		box.dispatchEvent(new Event('input'));
		expect(picker.value.items).toEqual(['One', 'Two']);
		options(picker)[1]!.click();
		expect(picker.value).toEqual({ layout: 'stacked', items: ['One', 'Two'] });
	});

	it('draws host previews (markup or elements) once per layout, without scripts', () => {
		const picker = mount();
		const calls: [string, boolean][] = [];
		picker.drawPreview = (layout, _size, tile) => {
			calls.push([layout.id, tile]);
			if (tile)
				return `<svg xmlns="http://www.w3.org/2000/svg"><script>x</script><title>${layout.id}</title></svg>`;
			const live = document.createElement('div');
			live.className = 'live';
			return live;
		};
		expect(root(picker).querySelector('.tile title')!.textContent).toBe('blocks');
		expect(root(picker).querySelector('.tile script')).toBeNull();
		expect(root(picker).querySelector('.preview .live')).not.toBeNull();
		const before = calls.length;
		picker.labels = { layouts: 'Layouts' };
		expect(calls.length).toBe(before);
		expect(root(picker).querySelector('[role="listbox"]')!.getAttribute('aria-label')).toBe(
			'Layouts',
		);
	});

	it('restores a value in another category', () => {
		const picker = mount();
		picker.value = { layout: 'ring', items: ['Q'] };
		expect(picker.category).toBe('cycle');
		expect(options(picker).map((option) => option.dataset.layout)).toEqual(['ring']);
	});
});
