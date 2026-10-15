// @vitest-environment jsdom
import { registerRibbonTabs, resetRibbonTabs, whatIfCommands } from 'ooxml-core/xlsx/ui';
import { afterEach, describe, expect, it } from 'vitest';
import { flush, shellFixture } from '../../test-support/shell';
import { createRibbon } from '../ribbon';
import { closeRibbonPopover } from '../popover';
import { dataTab } from './formulas-data';

afterEach(() => {
	closeRibbonPopover();
	resetRibbonTabs();
	document.body.replaceChildren();
});

function setup() {
	const fixture = shellFixture();
	fixture.core.commands.registerAll(whatIfCommands());
	registerRibbonTabs([dataTab()]);
	const ribbon = createRibbon(fixture.core.ctx, {
		openBackstage: () => undefined,
		isHidden: () => false,
	});
	fixture.core.ctx.root.append(ribbon.element);
	ribbon.refresh();
	const group = ribbon.element.querySelector<HTMLElement>(
		'#xve-panel-data .ribbon-group[aria-label="Forecast"]',
	)!;
	return { ...fixture, group };
}

describe('Data tab Forecast group', () => {
	it('has a What-If Analysis menu with the three tools and a disabled Forecast Sheet', async () => {
		const { group, core } = setup();
		await flush();
		const menu = group.querySelector<HTMLButtonElement>('.ribbon-menu-button')!;
		expect(menu.getAttribute('aria-label')).toBe('What-If Analysis');
		expect(menu.getAttribute('aria-haspopup')).toBe('true');
		expect(menu.dataset.menuCommands).toBe('data.scenario-manager data.goal-seek data.data-table');
		expect(menu.disabled).toBe(false);
		menu.click();
		await flush();
		const items = [
			...core.ctx.root.querySelectorAll<HTMLElement>('.ribbon-menu-list [role="menuitem"]'),
		].map((item) => item.textContent?.trim() ?? '');
		expect(items.join('|')).toContain('Scenario Manager...');
		expect(items.join('|')).toContain('Goal Seek...');
		expect(items.join('|')).toContain('Data Table...');
		const forecast = group.querySelector<HTMLButtonElement>(
			'[data-command="data.forecast-sheet"]',
		)!;
		expect(forecast.disabled).toBe(true);
		expect(forecast.title).toContain('FORECAST.ETS');
		expect(forecast.getAttribute('aria-description')).toContain('not available');
		expect(core.commands.isEnabled('data.forecast-sheet')).toBe(false);
	});

	it('every command in the group is registered', () => {
		const { core } = setup();
		for (const id of [
			'data.scenario-manager',
			'data.goal-seek',
			'data.data-table',
			'data.forecast-sheet',
		])
			expect(core.commands.get(id)).toBeDefined();
	});
});
