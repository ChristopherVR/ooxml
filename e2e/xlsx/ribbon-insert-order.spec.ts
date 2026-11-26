import { expect, test } from '@playwright/test';
import { clippedRibbonLabels } from '../ribbon-labels';
import { expectContextualTabContrast } from './contextual-tab-contrast';
import { openFixture } from './fixture-files';
import { clickRibbonControl, editor, FRAMEWORKS, goToCell, ribbon } from './helpers';

for (const framework of FRAMEWORKS) {
	test(`${framework}: Insert groups fit and commands use the correct tabs`, async ({ page }) => {
		await page.setViewportSize({ width: 1600, height: 900 });
		await openFixture(page, framework, 'excel-pivot-rows.xlsx');
		const band = ribbon(page);
		const tab = (name: string) => band.getByRole('tab', { name, exact: true });
		const panel = (id: string) => band.locator(`.ribbon-panel[data-tab="${id}"]`);
		await tab('Insert').click();
		const insert = panel('insert');
		expect(
			await insert
				.locator(':scope > .ribbon-group')
				.evaluateAll((groups) => groups.map((group) => group.getAttribute('data-label'))),
		).toEqual([
			'Tables',
			'Illustrations',
			'Add-ins',
			'Charts',
			'Sparklines',
			'Filters',
			'Links',
			'Forms',
			'Comments',
			'Text',
			'Symbols',
		]);
		await expect.poll(() => insert.evaluate(clippedRibbonLabels)).toEqual([]);
		const art = insert.locator('.ribbon-group[data-label="Illustrations"]');
		for (const name of ['Pictures', 'Shapes', 'Icons', '3D Models', 'Screenshot']) {
			const button = art.getByRole('button', { name, exact: true });
			await expect(button).toBeVisible();
			await expect(button.locator('svg path').first()).toHaveAttribute('d', /\S/u);
			expect((await button.boundingBox())?.height).toBe(76);
		}
		await page.setViewportSize({ width: 1280, height: 900 });
		await expect.poll(() => insert.evaluate(clippedRibbonLabels)).toEqual([]);
		await expect(
			insert.locator('[aria-label="PivotTable"], [data-command="insert.pivot-table"]'),
		).toHaveCount(1);
		await expect(insert.locator('[data-label="Sparklines"] button[data-command]')).toHaveCount(3);
		await expect(
			insert.locator(
				'[data-command^="pivot."], [data-command^="sparkline."], [data-command="insert.note"]',
			),
		).toHaveCount(0);

		await goToCell(page, 'A2');
		await expectContextualTabContrast(tab('PivotTable Analyze'));
		await expectContextualTabContrast(tab('PivotTable Design'));
		await tab('PivotTable Analyze').click();
		const analyze = panel('pivot-analyze');
		for (const id of [
			'settings',
			'filter-field',
			'refresh',
			'group',
			'group-selection',
			'ungroup',
			'field-settings',
		])
			await expect(analyze.locator(`[data-command="pivot.${id}"]`)).toHaveCount(1);
		const clear = band.locator('.ribbon-group[data-label="Actions"] button[aria-label="Clear"]');
		await clickRibbonControl(page, clear);
		await expect(editor(page).getByRole('menuitem', { name: 'Delete PivotTable' })).toBeVisible();
		await page.keyboard.press('Escape');

		await tab('Review').click();
		await clickRibbonControl(
			page,
			band.locator('.ribbon-group[data-label="Notes"] button.ribbon-menu-button'),
		);
		await expect(
			editor(page).getByRole('menuitem', { name: 'New Note', exact: true }),
		).toBeVisible();
	});
}
