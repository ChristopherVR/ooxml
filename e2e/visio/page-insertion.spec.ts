import { test, expect } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import { downloadCopy } from './ribbon.js';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: inserts a page through the worker, restores history and reopens the saved package`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1440, height: 900 });
		await page.goto(framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`);
		const viewer = page.locator('visio-viewer');
		const tabs = viewer.locator('.page-tabs');
		await expect(tabs.getByRole('button', { name: 'Insert Page', exact: true })).toBeDisabled();
		await page
			.locator('#file')
			.setInputFiles({
				name: 'pages.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer: await createVsdxFixture('Original page shape'),
			});
		await expect(page.locator('#file-name')).toHaveText('pages.vsdx');
		await expect(viewer.locator('svg text')).toContainText('Original page shape');
		await tabs.getByRole('button', { name: 'Insert Page', exact: true }).click();
		await expect(tabs.getByRole('tab')).toHaveCount(2);
		await expect(tabs.getByRole('tab', { name: 'Page-2', exact: true })).toHaveAttribute(
			'aria-selected',
			'true',
		);
		await expect(viewer.locator('[data-shape-id]')).toHaveCount(0);
		await viewer.locator('[command="undo"]').getByRole('button').click();
		await expect(tabs.getByRole('tab')).toHaveCount(1);
		await expect(viewer.locator('svg text')).toContainText('Original page shape');
		await viewer.locator('[command="redo"]').getByRole('button').click();
		await expect(tabs.getByRole('tab')).toHaveCount(2);
		const downloadButton = await downloadCopy(viewer);
		const pending = page.waitForEvent('download');
		await downloadButton.click();
		const stream = await (await pending).createReadStream();
		const chunks: Buffer[] = [];
		for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
		await page
			.locator('#file')
			.setInputFiles({
				name: 'reopened-pages.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer: Buffer.concat(chunks),
			});
		await expect(page.locator('#file-name')).toHaveText('reopened-pages.vsdx');
		await expect(tabs.getByRole('tab')).toHaveCount(2);
		await expect(tabs.getByRole('tab', { name: 'Page-2', exact: true })).toHaveCount(1);
	});
}
