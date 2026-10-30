import { expect, test } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

test('a selected chart shows Chart Design, whose Edit Data opens the chart editor', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg.paper')).toContainText('Release workflow');
	await page.locator('#file').setInputFiles({
		name: 'Contextual.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Contextual target'),
	});
	await expect(viewer.locator('svg.paper')).toContainText('Contextual target');
	const tab = (name: string) => viewer.getByRole('tab', { name, exact: true });

	// A plain shape gets Shape Format and no chart tab.
	await viewer.locator('svg.paper [data-shape-id]').first().click({ force: true });
	await expect(tab('Shape Format')).toBeVisible();
	await expect(tab('Chart Design')).toBeHidden();

	await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
	await viewer.locator('[command="chart"] button').click();
	await viewer.locator('[command="chart-dialog-insert"] button').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Inserted chart/);
	const chart = viewer.locator('svg.paper [data-shape-id]:has([data-embedded-chart])').first();
	await chart.click({ force: true });
	await expect(tab('Chart Design')).toBeVisible();
	await expect(tab('Shape Format')).toBeHidden();
	await tab('Chart Design').click();
	const panel = viewer.locator('[data-ribbon-tab="chart-design"]');
	await panel.locator('[command="chart-tab-data"] button').click();
	const dialog = viewer.locator('.chart-edit-dialog');
	await expect(dialog).toHaveAttribute('heading', 'Edit Data');
	await dialog.locator('[command="chart-edit-cancel"] button').click();
	expect(errors).toEqual([]);
});
