import { expect, test } from '@playwright/test';
import { fixture, shape, rectangle } from '../../src/core/visio/test-fixtures';
import { openDemo } from './demo-page';

test('Review finds an unlabeled graphic and saves accessible native title and description', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const bytes = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', rectangle)}</Shapes>` }],
	});
	await page.locator('#file').setInputFiles({
		name: 'accessibility.vsdx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from(bytes),
	});
	await expect(page.locator('#file-name')).toHaveText('accessibility.vsdx');
	const viewer = page.locator('visio-viewer');
	await viewer.getByRole('tab', { name: 'Review', exact: true }).click();
	await viewer.getByRole('button', { name: 'Check Accessibility', exact: true }).click();
	const report = viewer.locator('.accessibility-dialog');
	await expect(report).toContainText('1 graphic needs alternative text');
	await report.locator('div[aria-live] > button').click();
	const dialog = viewer.locator('.alt-text-dialog');
	await dialog.getByRole('textbox', { name: 'Title', exact: true }).fill('Service status');
	await dialog
		.getByRole('textbox', { name: 'Description', exact: true })
		.fill('Green box\nAll services operational.');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	const graphic = viewer.locator('svg.paper [data-shape-id="1"]');
	await expect(graphic).toHaveAttribute('aria-label', 'Service status');
	await expect(graphic).toHaveAttribute('aria-description', 'Green box\nAll services operational.');
	await viewer.locator('.qat [data-command="undo"]').click();
	await expect(graphic).not.toHaveAttribute('aria-label', 'Service status');
	await viewer.locator('.qat [data-command="redo"]').click();
	await expect(graphic).toHaveAttribute('aria-label', 'Service status');
	await viewer.getByRole('button', { name: 'Check Accessibility', exact: true }).click();
	await expect(report).toContainText('No missing text alternatives found.');
	await report
		.locator('office-ui-button[command="accessibility-dialog-close"]')
		.getByRole('button', { name: 'Close', exact: true })
		.click();
	await viewer.getByRole('button', { name: 'Page Alt Text', exact: true }).click();
	await dialog.getByRole('textbox', { name: 'Title', exact: true }).fill('Process overview');
	await dialog
		.getByRole('textbox', { name: 'Description', exact: true })
		.fill('Requests arrive\nthen proceed to approval.');
	await dialog.screenshot({ path: test.info().outputPath('page-alt-text.png') });
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(viewer.locator('svg.paper')).toHaveAttribute('aria-label', 'Process overview');
	await viewer.locator('.qat [data-command="undo"]').click();
	await expect(viewer.locator('svg.paper')).not.toHaveAttribute('aria-label', 'Process overview');
});
