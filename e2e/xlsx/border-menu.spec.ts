import { expect, test } from '@playwright/test';
import { clickRibbonControl, FRAMEWORKS, grid, newWorkbook, ribbon } from './helpers';

for (const framework of FRAMEWORKS) {
	test(`${framework}: the Borders menu shows a distinct icon for each preset`, async ({ page }) => {
		await newWorkbook(page, framework);
		await clickRibbonControl(page, ribbon(page).getByRole('button', { name: 'Borders options' }));

		const menu = page.getByRole('menu', { name: 'Borders', exact: true });
		const topBorder = menu.getByRole('menuitem', { name: 'Top Border', exact: true });
		await expect(topBorder).toBeVisible();
		const items = menu.getByRole('menuitem');
		const icons = items.locator('svg path[d]');
		const itemCount = 15;
		await expect(items).toHaveCount(itemCount);
		await expect(icons).toHaveCount(itemCount);
		const paths = await icons.evaluateAll((nodes) => nodes.map((node) => node.getAttribute('d')));
		expect(new Set(paths).size, 'distinct rendered border menu icons').toBe(itemCount);

		await topBorder.click();
		await expect(grid(page).locator('.xg-b')).toHaveCSS('border-top-style', 'solid');
	});
}
