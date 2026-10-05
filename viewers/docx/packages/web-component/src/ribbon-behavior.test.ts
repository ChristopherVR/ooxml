// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createRibbon, setRibbonLocale } from './ribbon';
import { ribbonFile, ribbonTab, ribbonTabs } from './test-support';

const mount = () => {
	const ribbon = createRibbon();
	document.body.append(ribbon);
	return ribbon;
};

afterEach(() => (document.body.innerHTML = ''));

describe('ribbon collapse', () => {
	it('uses the shared ribbon collapse, named in the display language', async () => {
		const ribbon = mount();
		await ribbonTab(ribbon, 'home');
		expect(ribbon.hasAttribute('collapsible')).toBe(true);
		const button = ribbon.shadowRoot!.querySelector<HTMLButtonElement>('.collapse')!;
		expect(button.getAttribute('aria-label')).toBe('Collapse the ribbon');
		button.click();
		expect(ribbon.hasAttribute('collapsed')).toBe(true);
		setRibbonLocale(ribbon, 'fr');
		await ribbonTab(ribbon, 'home');
		expect(button.getAttribute('aria-label')).toBe('Réduire le ruban');
	});
});

/** The badges the key tips put on the tab row, which lives in the shared ribbon's shadow root. */
const badges = async (ribbon: HTMLElement) =>
	[...(await ribbonTabs(ribbon))]
		.flatMap((tab) => [...tab.querySelectorAll('.dve-keytip')])
		.map((badge) => badge.textContent);

describe('tab KeyTips', () => {
	it('shows a badge per tab and opens the tab for its letter', async () => {
		const ribbon = mount();
		await ribbonTab(ribbon, 'home');
		ribbon.dispatchEvent(new CustomEvent('dve-keytips'));
		expect(await badges(ribbon)).toEqual(['H', 'N', 'P', 'S', 'R', 'W']);
		expect((await ribbonFile(ribbon)).querySelector('.dve-keytip')?.textContent).toBe('F');
		ribbon.dispatchEvent(new KeyboardEvent('keydown', { key: 'n' }));
		expect((await ribbonTab(ribbon, 'insert')).getAttribute('aria-selected')).toBe('true');
		expect(await badges(ribbon)).toEqual([]);
	});

	it('hides the badges on an unmatched key or Escape without changing tabs', async () => {
		const ribbon = mount();
		await ribbonTab(ribbon, 'home');
		ribbon.dispatchEvent(new CustomEvent('dve-keytips'));
		ribbon.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(await badges(ribbon)).toEqual([]);
		expect((await ribbonTab(ribbon, 'home')).getAttribute('aria-selected')).toBe('true');
	});
});
