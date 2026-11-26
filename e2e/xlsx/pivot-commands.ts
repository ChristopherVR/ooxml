import type { Page } from '@playwright/test';
import { clickRibbonControl, editor, ribbon } from './helpers';

type AnalyzeCommand =
	| 'refresh'
	| 'settings'
	| 'field-settings'
	| 'group-selection'
	| 'group'
	| 'ungroup'
	| 'delete';

export async function clickPivotAnalyzeCommand(page: Page, command: AnalyzeCommand) {
	const bar = ribbon(page);
	const tab = bar.getByRole('tab', { name: 'PivotTable Analyze', exact: true });
	if ((await tab.getAttribute('aria-selected')) !== 'true') await tab.click();
	if (command === 'delete') {
		await clickRibbonControl(page, bar.locator('[data-menu-commands~="pivot.delete"]'));
		await editor(page).getByRole('menuitem', { name: 'Delete PivotTable', exact: true }).click();
	} else {
		await clickRibbonControl(page, bar.locator(`[data-command="pivot.${command}"]`));
	}
}
