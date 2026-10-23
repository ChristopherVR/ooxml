/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
/**
 * Insert > Chart > Insert Chart...: the shared `pptx-ui-insert-chart-dialog` (the
 * `office-ui-chart-picker` Excel and Visio use, with previews drawn by the slide's chart
 * renderer) opens from the end of the Chart menu in every demo, and Insert adds the chart with
 * the chosen type, grouping and title. Framework-neutral: every binding mounts the same element.
 */
import { expect, test } from '@playwright/test';

import { loadDeck, ribbonTab, slideElements } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

test('Insert Chart... opens the shared picker and inserts the chosen chart', async ({ page }) => {
	await loadDeck(page);
	await ribbonTab(page, 'Insert').click();
	const insert = page.locator('pptx-ui-ribbon-insert');
	const before = await slideElements(page).count();
	await insert.getByRole('button', { name: 'Chart', exact: true }).click();
	await insert.locator('[data-insert-item="__chartDialog__"]').click();

	const dialog = page.getByRole('dialog', { name: 'Insert Chart' });
	await expect(dialog).toBeVisible();
	const picker = page.locator('pptx-ui-insert-chart-dialog office-ui-chart-picker');
	// Every type tile and the large preview are drawn by the chart renderer.
	await expect(picker.locator('office-ui-gallery .tile svg').first()).toBeVisible();
	expect(await picker.locator('office-ui-gallery .tile svg').count()).toBeGreaterThan(6);
	await expect(picker.locator('.preview svg')).toBeVisible();

	await picker.locator('[data-gallery-item="bar"]').click();
	await picker.locator('select[data-key="grouping"]').selectOption('stacked');
	await picker.locator('input[data-key="title"]').fill('Quarterly sales');
	// The footer is slotted into the shared dialog element, not the shadow box getByRole finds.
	await page
		.locator('pptx-ui-insert-chart-dialog')
		.getByRole('button', { name: 'Insert', exact: true })
		.click();

	await expect(dialog).toBeHidden();
	await expect(slideElements(page)).toHaveCount(before + 1);
	await expect(page.locator('[data-pptx-viewport]').getByText('Quarterly sales')).toBeVisible();
});

test('Escape closes the Insert Chart dialog without inserting', async ({ page }) => {
	await loadDeck(page);
	await ribbonTab(page, 'Insert').click();
	const insert = page.locator('pptx-ui-ribbon-insert');
	const before = await slideElements(page).count();
	await insert.getByRole('button', { name: 'Chart', exact: true }).click();
	await insert.locator('[data-insert-item="__chartDialog__"]').click();
	const dialog = page.getByRole('dialog', { name: 'Insert Chart' });
	await expect(dialog).toBeVisible();
	await page.keyboard.press('Escape');
	await expect(dialog).toBeHidden();
	await expect(slideElements(page)).toHaveCount(before);
	// It opens again after a close.
	await insert.getByRole('button', { name: 'Chart', exact: true }).click();
	await insert.locator('[data-insert-item="__chartDialog__"]').click();
	await expect(dialog).toBeVisible();
});
