import { registerOfficeUi } from '../index';
import type { OfficeSmartArtTextNode, OfficeSmartArtTextPaneLabels } from './smartart-text-pane';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Pane = HTMLElement & {
	nodes: OfficeSmartArtTextNode[];
	selected: string | null;
	closable: boolean;
	labels: Partial<OfficeSmartArtTextPaneLabels>;
	focusSelected(): void;
};

function mount(): Pane {
	const pane = document.createElement('office-ui-smartart-text-pane') as Pane;
	pane.nodes = [
		{ id: 'a', text: 'Alpha', level: 0 },
		{ id: 'a1', text: 'Child', level: 1 },
		{ id: 'b', text: '', level: 0 },
	];
	document.body.append(pane);
	return pane;
}

const inputs = (pane: Pane) => [
	...pane.shadowRoot!.querySelectorAll<HTMLInputElement>('input[data-node]'),
];
const listen = (pane: Pane, name: string) => {
	const seen: unknown[] = [];
	pane.addEventListener(name, (event) => seen.push((event as CustomEvent).detail));
	return seen;
};
const key = (input: HTMLInputElement, init: KeyboardEventInit) =>
	input.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }));

describe('office-ui-smartart-text-pane', () => {
	it('shows the nodes as an outline indented by level, with a placeholder for empty text', () => {
		const pane = mount();
		const lines = [...pane.shadowRoot!.querySelectorAll<HTMLElement>('[role="treeitem"]')];
		expect(lines.map((line) => line.getAttribute('aria-level'))).toEqual(['1', '2', '1']);
		expect(lines[1]!.style.getPropertyValue('--level')).toBe('1');
		expect(inputs(pane).map((input) => input.value)).toEqual(['Alpha', 'Child', '']);
		expect(inputs(pane)[2]!.placeholder).toBe('[Text]');
		expect(pane.shadowRoot!.querySelector('[role="tree"]')!.getAttribute('aria-label')).toBe(
			'SmartArt text',
		);
	});

	it('turns typing and keys into requests for the host', () => {
		const pane = mount();
		const changes = listen(pane, 'office-smartart-text-change');
		const adds = listen(pane, 'office-smartart-node-add');
		const demotes = listen(pane, 'office-smartart-node-demote');
		const promotes = listen(pane, 'office-smartart-node-promote');
		const removes = listen(pane, 'office-smartart-node-remove');
		const [alpha, child, empty] = inputs(pane) as [
			HTMLInputElement,
			HTMLInputElement,
			HTMLInputElement,
		];
		alpha.value = 'Alpha 2';
		alpha.dispatchEvent(new Event('input', { bubbles: true }));
		expect(changes).toEqual([{ id: 'a', text: 'Alpha 2' }]);
		expect(key(alpha, { key: 'Enter' })).toBe(false);
		expect(adds).toEqual([{ id: 'a' }]);
		key(child, { key: 'Tab' });
		key(child, { key: 'Tab', shiftKey: true });
		expect(demotes).toEqual([{ id: 'a1' }]);
		expect(promotes).toEqual([{ id: 'a1' }]);
		// Backspace removes only an empty line.
		expect(key(child, { key: 'Backspace' })).toBe(true);
		key(empty, { key: 'Backspace' });
		expect(removes).toEqual([{ id: 'b' }]);
	});

	it('reports the focused node and moves focus with the arrow keys', () => {
		const pane = mount();
		const selects = listen(pane, 'office-smartart-node-select');
		const [alpha, child] = inputs(pane) as [HTMLInputElement, HTMLInputElement];
		alpha.focus();
		expect(selects).toEqual([{ id: 'a' }]);
		expect(pane.selected).toBe('a');
		key(alpha, { key: 'ArrowDown' });
		expect(pane.shadowRoot!.activeElement).toBe(child);
		expect(selects).toEqual([{ id: 'a' }, { id: 'a1' }]);
	});

	it('keeps the caret line and focuses the selected node after the host updates', () => {
		const pane = mount();
		const [alpha] = inputs(pane) as [HTMLInputElement];
		alpha.focus();
		key(alpha, { key: 'Enter' });
		pane.nodes = [
			{ id: 'a', text: 'Alpha', level: 0 },
			{ id: 'n', text: '', level: 0 },
			{ id: 'a1', text: 'Child', level: 1 },
			{ id: 'b', text: '', level: 0 },
		];
		pane.selected = 'n';
		expect(pane.shadowRoot!.activeElement).toBe(inputs(pane)[1]);
		expect(inputs(pane)[0]!.value).toBe('Alpha');
	});

	it('translates its strings, shows an empty state and closes on request', () => {
		const pane = mount();
		pane.labels = { title: 'Texte', placeholder: '[Texte]' };
		pane.closable = true;
		const closes = listen(pane, 'office-smartart-text-pane-close');
		expect(pane.shadowRoot!.querySelector('.title')!.textContent).toBe('Texte');
		expect(inputs(pane)[2]!.placeholder).toBe('[Texte]');
		pane.shadowRoot!.querySelector<HTMLButtonElement>('.close')!.click();
		expect(closes).toHaveLength(1);
		pane.nodes = [];
		expect(pane.shadowRoot!.querySelector('.empty')!.textContent).toContain('no shapes');
	});
});
