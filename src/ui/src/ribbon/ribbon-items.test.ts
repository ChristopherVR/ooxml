// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { registerOfficeUi } from '../index';
import type { OfficeUiRibbon } from './ribbon-tabs';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

function tabRow(): OfficeUiRibbon {
	const el = document.createElement('office-ui-ribbon') as OfficeUiRibbon;
	el.innerHTML = '<div slot="actions"><button>Share</button></div><div id="body">Commands</div>';
	document.body.append(el);
	el.items = [
		{ id: 'home', label: 'Home', title: 'Home tab' },
		{ id: 'view', label: 'View' },
		{ id: 'shape', label: 'Shape Format', contextual: true },
	];
	return el;
}

const tabs = (el: OfficeUiRibbon) => [
	...el.renderRoot.querySelectorAll<HTMLButtonElement>('[role="tab"]'),
];

describe('office-ui-ribbon with items', () => {
	it('draws the tabs from the list and leaves its children alone', () => {
		const el = tabRow();
		expect(tabs(el).map((tab) => tab.textContent)).toEqual(['Home', 'View', 'Shape Format']);
		expect(tabs(el)[0]!.title).toBe('Home tab');
		expect(tabs(el)[1]!.hasAttribute('title')).toBe(false);
		expect(tabs(el)[2]!.hasAttribute('data-contextual')).toBe(true);
		const body = el.querySelector<HTMLElement>('#body')!;
		expect(body.hidden).toBe(false);
		expect(body.hasAttribute('role')).toBe(false);
	});

	it('selects only an id from the list and lets the host own the selection', () => {
		const el = tabRow();
		expect(el.selected).toBe('');
		el.selected = 'view';
		expect(tabs(el)[1]!.getAttribute('aria-selected')).toBe('true');
		el.selected = 'missing';
		expect(el.selected).toBe('');
		const picked: string[] = [];
		el.addEventListener('office-ribbon-select', (event) => {
			picked.push((event as CustomEvent<{ tab: string }>).detail.tab);
			event.preventDefault();
		});
		el.selected = 'home';
		tabs(el)[2]!.click();
		// The selected tab is reported too: the host may peek a collapsed ribbon's commands.
		tabs(el)[0]!.click();
		expect(picked).toEqual(['shape', 'home']);
		expect(el.selected).toBe('home');
	});

	it('moves focus with the arrow keys even before the host applies the selection', () => {
		const el = tabRow();
		el.selected = 'home';
		el.addEventListener('office-ribbon-select', (event) => event.preventDefault());
		tabs(el)[0]!.focus();
		tabs(el)[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
		expect(el.shadowRoot!.activeElement).toBe(tabs(el)[1]);
	});

	it('updates the tabs when the list changes', () => {
		const el = tabRow();
		el.items = [{ id: 'home', label: 'Start' }];
		expect(tabs(el).map((tab) => tab.textContent)).toEqual(['Start']);
	});
});
