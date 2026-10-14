// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';
import { applyRibbonVisibility } from './ribbon-visibility';

afterEach(() => document.body.replaceChildren());

describe('Word Home split commands', () => {
	it('keeps the default action separate from repeated dropdown selections', () => {
		const ribbon = createRibbon();
		const actions: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) =>
			actions.push((event as CustomEvent).detail),
		);
		ribbon.querySelector<HTMLButtonElement>('[aria-label="Underline"]')!.click();
		expect(actions.pop()).toEqual({ type: 'format', key: 'underline' });
		const menu = ribbon.querySelector<HTMLSelectElement>('[aria-label="Underline style"]')!;
		for (let count = 0; count < 2; count++) {
			menu.value = 'double';
			menu.dispatchEvent(new Event('change'));
			expect(actions.pop()).toEqual({ type: 'underline', value: 'double' });
			expect(menu.selectedIndex).toBe(-1);
		}
	});

	it('hides a menu independently and restores it when customization changes', () => {
		const ribbon = createRibbon();
		const menu = ribbon.querySelector('[aria-label="Underline style"]')!;
		const main = ribbon.querySelector('[aria-label="Underline"]')!;
		applyRibbonVisibility(ribbon, true, ['underline-style']);
		expect(menu.hasAttribute('data-dve-hidden')).toBe(true);
		expect(main.hasAttribute('data-dve-hidden')).toBe(false);
		applyRibbonVisibility(ribbon, true, []);
		expect(menu.hasAttribute('data-dve-hidden')).toBe(false);
	});

	it('routes paste and list choices and hides their menus with the main action', () => {
		const ribbon = createRibbon();
		const actions: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) =>
			actions.push((event as CustomEvent).detail),
		);
		for (const [label, value, action] of [
			['Paste options', 'text', { type: 'clipboard', key: 'paste', plainText: true }],
			['Bullets options', 'remove', { type: 'list', key: 'remove' }],
		] as const) {
			const menu = ribbon.querySelector<HTMLSelectElement>(`[aria-label="${label}"]`)!;
			menu.value = value;
			menu.dispatchEvent(new Event('change'));
			expect(actions.pop()).toEqual(action);
		}
		applyRibbonVisibility(ribbon, true, ['paste', 'underline']);
		for (const label of ['Paste options', 'Underline style']) {
			const menu = ribbon.querySelector(`[aria-label="${label}"]`)!;
			expect(menu.closest('.ribbon-command-split')!.hasAttribute('data-dve-hidden')).toBe(true);
		}
	});
});
