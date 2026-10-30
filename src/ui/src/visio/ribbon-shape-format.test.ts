// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';
import { renderContextualTabs } from './viewer-contextual-tabs';
import type { ViewerState } from './controller';

const stateWith = (count: number) =>
	({
		selectedShapes: Array.from({ length: count }, (_, i) => ({ id: String(i) })),
	}) as never as ViewerState;

describe('Shape Format contextual tab', () => {
	it('is a contextual tab, hidden until shapes are selected', () => {
		const ribbon = createRibbon(document);
		document.body.append(ribbon);
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="shape-format"]')!;
		expect(panel.dataset.contextual).toBeDefined();
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		const groups = [...panel.querySelectorAll('office-ui-ribbon-group')].map((g) =>
			g.getAttribute('label'),
		);
		expect(groups).toEqual(['Shape Styles', 'WordArt Styles', 'Arrange', 'Size']);
		ribbon.remove();
	});

	it('shows with a selection and mirrors the Home state onto its twins', () => {
		const ribbon = createRibbon(document);
		document.body.append(ribbon);
		const setTabHidden = (tab: string, _reason: string, hidden: boolean) =>
			ribbon
				.querySelector(`[data-ribbon-tab="${tab}"]`)!
				.toggleAttribute('data-tab-hidden', hidden);
		Object.assign(ribbon, { setTabHidden });
		const panel = ribbon.querySelector<HTMLElement>('[data-ribbon-tab="shape-format"]')!;
		const home = ribbon.querySelector<HTMLElement>('#home-panel [command="bring-forward"]')!;
		const twin = panel.querySelector<HTMLElement>('[command="bring-forward"]')!;
		home.setAttribute('disabled', '');
		renderContextualTabs(ribbon, stateWith(1));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(false);
		expect(twin.hasAttribute('disabled')).toBe(true);
		home.removeAttribute('disabled');
		renderContextualTabs(ribbon, stateWith(2));
		expect(twin.hasAttribute('disabled')).toBe(false);
		renderContextualTabs(ribbon, stateWith(0));
		expect(panel.hasAttribute('data-tab-hidden')).toBe(true);
		ribbon.remove();
	});

	it('lists the unsupported WordArt commands disabled with a reason', () => {
		const ribbon = createRibbon(document);
		for (const id of ['text-fill', 'text-outline', 'text-effects']) {
			const el = ribbon.querySelector<HTMLElement>(
				`[data-ribbon-tab="shape-format"] [data-menu="${id}"]`,
			)!;
			expect(el.hasAttribute('disabled')).toBe(true);
			expect(el.title).toContain('not available yet');
		}
	});
});
