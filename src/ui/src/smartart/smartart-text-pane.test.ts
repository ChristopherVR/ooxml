import { registerOfficeUi } from '../index';
import type { OfficeSmartArtTextNode, OfficeSmartArtTextPaneLabels } from './smartart-text-pane';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Pane = HTMLElement & {
	nodes: OfficeSmartArtTextNode[];
	selected: string | null;
	closable: boolean;
	disabled: boolean;
	tools: boolean;
	appendDisabled: boolean;
	notes: string[];
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

	it('shows markers and the tools, and turns them into requests', () => {
		const pane = mount();
		pane.tools = true;
		pane.notes = ['Hint one'];
		pane.nodes = [
			{ id: 'a', text: 'Alpha', level: 0, marker: '1', label: 'Item 1', canMoveUp: false },
			{ id: 'a1', text: 'Child', level: 1, marker: '•', canAddChild: false, style: { bold: true } },
			{ id: 'b', text: '', level: 0, marker: '2', canRemove: false },
		];
		const root = pane.shadowRoot!;
		expect([...root.querySelectorAll('.marker')].map((m) => m.textContent)).toEqual([
			'1',
			'•',
			'2',
		]);
		expect(inputs(pane)[0]!.getAttribute('aria-label')).toBe('Item 1');
		expect(inputs(pane)[0]!.dataset.testid).toBe('smartart-node-text');
		expect(root.querySelector('.note')!.textContent).toBe('Hint one');
		const lines = [...root.querySelectorAll<HTMLElement>('.line')];
		const tool = (line: number, name: string) =>
			lines[line]!.querySelector<HTMLButtonElement>(`.${name}`);
		expect(tool(0, 'move-up')!.disabled).toBe(true);
		expect(tool(0, 'move-up')!.getAttribute('aria-label')).toBe('Move up: Item 1');
		expect(tool(1, 'add-child')).toBeNull();
		expect(tool(1, 'bold')!.getAttribute('aria-pressed')).toBe('true');
		expect(tool(2, 'remove')!.disabled).toBe(true);
		expect(tool(2, 'remove')!.title).toBe('The layout needs this shape');

		const moves = listen(pane, 'office-smartart-node-move');
		const children = listen(pane, 'office-smartart-node-add-child');
		const removes = listen(pane, 'office-smartart-node-remove');
		const styles = listen(pane, 'office-smartart-node-style');
		const appends = listen(pane, 'office-smartart-node-append');
		tool(0, 'move-down')!.click();
		tool(0, 'add-child')!.click();
		tool(1, 'remove')!.click();
		tool(1, 'bold')!.click();
		const fill = lines[0]!.querySelector<HTMLInputElement>('.fill input')!;
		fill.value = '#ff0000';
		fill.dispatchEvent(new Event('change', { bubbles: true }));
		root.querySelector<HTMLButtonElement>('.append')!.click();
		expect(moves).toEqual([{ id: 'a', direction: 1 }]);
		expect(children).toEqual([{ id: 'a' }]);
		expect(removes).toEqual([{ id: 'a1' }]);
		expect(styles).toEqual([
			{ id: 'a1', style: { bold: false } },
			{ id: 'a', style: { fillColor: '#ff0000' } },
		]);
		expect(appends).toHaveLength(1);

		// A line the layout needs ignores Backspace and Delete; Delete removes other empty lines.
		key(inputs(pane)[2]!, { key: 'Delete' });
		expect(removes).toEqual([{ id: 'a1' }]);
		pane.nodes = [{ id: 'c', text: '', level: 0 }];
		key(inputs(pane)[0]!, { key: 'Delete' });
		expect(removes).toEqual([{ id: 'a1' }, { id: 'c' }]);
	});

	it('is read-only while disabled', () => {
		const pane = mount();
		pane.tools = true;
		pane.disabled = true;
		const adds = listen(pane, 'office-smartart-node-add');
		const [alpha] = inputs(pane) as [HTMLInputElement];
		expect(alpha.disabled).toBe(true);
		key(alpha, { key: 'Enter' });
		expect(adds).toEqual([]);
		expect(pane.shadowRoot!.querySelector<HTMLButtonElement>('.append')!.disabled).toBe(true);
		expect(
			[...pane.shadowRoot!.querySelectorAll<HTMLButtonElement>('.tool')].every((b) => b.disabled),
		).toBe(true);
		pane.disabled = false;
		pane.appendDisabled = true;
		expect(pane.shadowRoot!.querySelector<HTMLButtonElement>('.append')!.disabled).toBe(true);
	});
});
