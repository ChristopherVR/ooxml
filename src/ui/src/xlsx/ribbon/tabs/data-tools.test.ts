// @vitest-environment jsdom
import { dataCommands, registerRibbonTabs, resetRibbonTabs } from 'ooxml-core/xlsx/ui';
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
	fixture.core.commands.registerAll(dataCommands());
	registerRibbonTabs([dataTab()]);
	const ribbon = createRibbon(fixture.core.ctx, {
		openBackstage: () => undefined,
		isHidden: () => false,
	});
	fixture.core.ctx.root.append(ribbon.element);
	ribbon.refresh();
	return { ...fixture, ribbon };
}

describe('Data tab tools (Flash Fill, Consolidate, Subtotal, Advanced Filter)', () => {
	it('renders the four buttons in Sort & Filter, Data Tools and Outline groups', async () => {
		const { ribbon } = setup();
		await flush();
		const panel = ribbon.element.querySelector<HTMLElement>('#xve-panel-data')!;
		const byCommand = (groupLabel: string, command: string) =>
			panel.querySelector<HTMLButtonElement>(
				`.ribbon-group[aria-label="${groupLabel}"] [data-command="${command}"]`,
			);
		expect(byCommand('Sort & Filter', 'data.advanced-filter')?.disabled).toBe(false);
		expect(byCommand('Data Tools', 'data.flash-fill')?.disabled).toBe(false);
		expect(byCommand('Data Tools', 'data.consolidate')?.disabled).toBe(false);
		expect(byCommand('Outline', 'data.subtotal')?.disabled).toBe(false);
	});
});
