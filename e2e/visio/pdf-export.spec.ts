import { test, expect } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { fileBackstage } from './ribbon';
import { openDemo } from './demo-page';

test('File > Export > PDF honours the page range and shows the per-page report', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'report.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Report'),
	});
	await fileBackstage(viewer, 'export');
	const scope = viewer.locator('[data-pdf-scope]');
	await expect(scope).toHaveValue('all');
	await expect(viewer.locator('[data-pdf-range]')).toBeHidden();
	await scope.selectOption('range');
	await expect(viewer.locator('[data-pdf-range]')).toBeVisible();
	await viewer.locator('[data-pdf-from]').fill('1');
	await viewer.locator('[data-pdf-to]').fill('1');

	const downloading = page.waitForEvent('download');
	await viewer.locator('[data-backstage-action="export-pdf"]').first().click();
	const download = await downloading;
	const chunks: Buffer[] = [];
	for await (const chunk of (await download.createReadStream())!) chunks.push(Buffer.from(chunk));
	const pdf = Buffer.concat(chunks).toString('latin1');
	expect(pdf.startsWith('%PDF-1.4')).toBe(true);
	expect(pdf).toContain('/Count 1');

	// The Export page stays open and lists how each page was drawn.
	const report = viewer.locator('[data-pdf-report]');
	await expect(report).toBeVisible();
	await expect(report).toContainText('PDF created: 1 page(s)');
	await expect(report.locator('li').first()).toContainText('drawn as vectors');
});
