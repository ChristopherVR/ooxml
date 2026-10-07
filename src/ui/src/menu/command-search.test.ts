import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { registerOfficeUi, type OfficeSearchCommand } from '../index';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Search = HTMLElement & { commands: OfficeSearchCommand[]; isOpen: boolean };
function setup() {
	const el = document.createElement('office-ui-command-search') as Search;
	document.body.append(el);
	el.commands = [
		{ id: 'bold', label: 'Bold', disabled: true, title: 'Needs text formatting' },
		{ id: 'grid', label: 'Grid', keywords: 'show lines', description: 'View › Show' },
		{ id: 'ruler', label: 'Ruler', keywords: 'show measure' },
	];
	const input = el.shadowRoot!.querySelector('input')!;
	const type = (value: string) => {
		input.value = value;
		input.dispatchEvent(new Event('input'));
	};
	const key = (k: string) =>
		input.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));
	const options = () => [...el.shadowRoot!.querySelectorAll('li')].map((li) => li.textContent);
	const seen: string[] = [];
	el.addEventListener('office-command', (e) => seen.push((e as CustomEvent).detail.command));
	return { el, input, type, key, options, seen };
}

describe('office-ui-command-search', () => {
	it('is a labelled combobox with the Office placeholder', () => {
		const { input } = setup();
		expect(input.getAttribute('role')).toBe('combobox');
		expect(input.placeholder).toBe('Tell me what you want to do');
		expect(input.getAttribute('aria-label')).toBe('Tell me what you want to do');
	});

	it('matches labels and keywords and runs the active enabled match', () => {
		const { el, input, type, key, options, seen } = setup();
		type('show');
		expect(el.isOpen).toBe(true);
		expect(options()).toEqual(['GridView › Show', 'Ruler']);
		expect(input.getAttribute('aria-activedescendant')).toMatch(/-0$/);
		key('ArrowDown');
		key('Enter');
		expect(seen).toEqual(['ruler']);
		expect(el.isOpen).toBe(false);
		expect(input.value).toBe('');
	});

	it('shows disabled matches with their reason but never runs them', () => {
		const { type, key, options, seen, el } = setup();
		type('bo');
		expect(options()).toEqual(['BoldNeeds text formatting']);
		key('Enter');
		el.shadowRoot!.querySelector('li')!.click();
		expect(seen).toEqual([]);
		type('zzz');
		expect(options()).toEqual(['No matching commands']);
		key('Escape');
		expect(el.isOpen).toBe(false);
	});

	it('ranks usable commands and label prefixes first', () => {
		const { el, type, options } = setup();
		el.commands = [
			{ id: 'a', label: 'Bold text', disabled: true, title: 'Off' },
			{ id: 'b', label: 'Text bold' },
			{ id: 'c', label: 'Bold' },
		];
		type('bold');
		expect(options()).toEqual(['Bold', 'Text bold', 'Bold textOff']);
	});
});
