import { test, expect } from '@playwright/test';
import { downloadCopy, loadSampleTemplate, saveCommand, taskPane } from './ribbon.js';
import { createVsdxFixture } from './fixture.mjs';

test('edits literal text locally, undoes/redoes and downloads a reopenable VSDX copy', async ({
	page,
	baseURL,
}) => {
	const external: string[] = [];
	page.on('request', (request) => {
		if (!request.url().startsWith(baseURL!) && !request.url().startsWith('data:'))
			external.push(request.url());
	});
	await page.goto('/demo/?sample=1');
	await expect(saveCommand(page.locator('visio-viewer'))).toBeDisabled();
	await page.locator('#file').setInputFiles({
		name: 'editable.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Before edit'),
	});
	await expect(page.locator('#file-name')).toHaveText('editable.vsdx');
	await page.locator('visio-viewer [data-shape-id="1"]').click();
	await page.locator('visio-viewer .edit-controls summary').click();
	const input = page.getByLabel('Selected shape text', { exact: true });
	await expect(input).toHaveValue('Before edit');
	await input.fill('<script>Literal edited text</script>');
	await page.getByRole('button', { name: 'Apply text', exact: true }).click();
	await expect(page.locator('visio-viewer svg text')).toContainText(
		'<script>Literal edited text</script>',
	);
	await expect(page.locator('#edit-label')).toHaveText('EDITED COPY');
	await expect(page.locator('visio-viewer script')).toHaveCount(0);
	await page
		.locator('visio-viewer .edit-controls')
		.getByRole('button', { name: 'Undo', exact: true })
		.click();
	await expect(page.locator('visio-viewer svg text')).toContainText('Before edit');
	await expect(page.locator('#edit-label')).toHaveText('ORIGINAL');
	await page
		.locator('visio-viewer .edit-controls')
		.getByRole('button', { name: 'Redo', exact: true })
		.click();
	await expect(page.locator('visio-viewer svg text')).toContainText('Literal edited text');
	const saveAs = await downloadCopy(page.locator('visio-viewer'));
	const downloadEvent = page.waitForEvent('download');
	await saveAs.click();
	const download = await downloadEvent;
	expect(download.suggestedFilename()).toBe('editable-edited-copy.vsdx');
	const stream = await download.createReadStream();
	const chunks: Buffer[] = [];
	for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
	await page.locator('#file').setInputFiles({
		name: 'reopened.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: Buffer.concat(chunks),
	});
	await expect(page.locator('#file-name')).toHaveText('reopened.vsdx');
	await expect(page.locator('visio-viewer svg text')).toContainText('Literal edited text');
	expect(external).toEqual([]);
});

test('mobile editor preserves literal drafts, keyboard cancellation and touch target sizes', async ({
	page,
}) => {
	await page.setViewportSize({ width: 390, height: 844 });
	await page.goto('/demo/?sample=1');
	await page.locator('#file').setInputFiles({
		name: 'mobile.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Original'),
	});
	await expect(page.locator('#file-name')).toHaveText('mobile.vsdx');
	await page.locator('visio-viewer [data-shape-id="1"]').click();
	await taskPane(page.locator('visio-viewer'), 'Inspector');
	await page.locator('visio-viewer .edit-controls summary').click();
	const input = page.getByLabel('Selected shape text', { exact: true });
	await input.fill('draft + - 0');
	await input.press('Escape');
	await expect(input).toHaveValue('Original');
	await expect(input).toBeFocused();
	for (const action of ['apply', 'cancel', 'undo', 'redo']) {
		const box = await page.locator(`visio-viewer [data-edit="${action}"]`).boundingBox();
		expect(box?.height).toBeGreaterThanOrEqual(44);
	}
	expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
	await loadSampleTemplate(page.locator('visio-viewer'));
	await expect(input).toHaveValue('');
	await expect(input).toBeDisabled();
});
