import { expect, test } from '@playwright/test';
import { parseVsdx } from 'ooxml-core/visio';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

test('the Text dialog Tabs tab writes tab stops that lay the text out', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await expect(viewer.locator('svg.paper')).toContainText('Release workflow');
	await page.locator('#file').setInputFiles({
		name: 'Tabs.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('ab\tcd'),
	});
	await expect(page.locator('#file-name')).toHaveText('Tabs.vsdx');
	const paper = viewer.locator('svg.paper');
	await expect(paper).toContainText('abcd');
	const status = viewer.locator('[data-status]');
	await paper.locator('[data-shape-id]').first().click();
	const starts = () =>
		paper
			.locator('[data-shape-id]')
			.first()
			.locator('text > tspan')
			.evaluateAll((rows) => rows.map((row) => Number(row.getAttribute('x'))));
	const before = await starts();

	await viewer.locator('office-ui-ribbon-group[launcher="font-dialog"] .launcher').click();
	const dialog = viewer.locator('.text-dialog');
	await dialog.locator('[data-text-tab="tabs"]').click();
	const field = dialog.locator('[name="text-tabStops"]');
	await expect(field).toHaveValue('');
	await field.fill('72 right');
	await dialog.locator('[label="OK"] button').click();
	await expect(status).toHaveText(/Updated text formatting/);

	// The part after the tab now ends on the 1 in stop of the text block.
	const after = await starts();
	expect(after).toHaveLength(2);
	expect(after).not.toEqual(before);

	const bytes = await viewer.evaluate((node) =>
		Array.from((node as VisioViewerElement).exportVsdx().bytes),
	);
	const tabs = (await parseVsdx(new Uint8Array(bytes))).pages[0]!.shapes[0]!.text.paragraphs![0]!
		.tabs;
	expect(tabs?.find((stop) => stop.position > 0)).toEqual({ position: 1, align: 'right' });

	// The dialog shows the saved stops again.
	await viewer.locator('office-ui-ribbon-group[launcher="font-dialog"] .launcher').click();
	await dialog.locator('[data-text-tab="tabs"]').click();
	await expect(field).toHaveValue('72 right');
	await page.screenshot({ path: test.info().outputPath('text-tab-stops.png') });
	expect(errors).toEqual([]);
});
