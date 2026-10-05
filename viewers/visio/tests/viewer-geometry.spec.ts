import { test, expect } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';

test('geometry controls create, resize, move existing shapes, delete and undo through the isolated worker', async ({
	page,
}) => {
	await page.goto('/demo/?sample=1');
	await page.locator('#file').setInputFiles({
		name: 'geometry.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Existing shape'),
	});
	await expect(page.locator('#file-name')).toHaveText('geometry.vsdx');
	await page.locator('visio-viewer .edit-controls summary').click();
	await page.getByLabel('New rectangle ID').fill('42');
	await page.getByLabel('Pin X (inches)').fill('2');
	await page.getByLabel('Pin Y (inches)').fill('3');
	await page.getByLabel('Width (inches)', { exact: true }).fill('1.5');
	await page.getByLabel('Height (inches)', { exact: true }).fill('1');
	await page.getByRole('button', { name: 'Create rectangle', exact: true }).click();
	const rectangle = page.locator('visio-viewer [data-shape-id="42"]');
	await expect(rectangle).toHaveCount(1);
	await rectangle.click();
	await page.getByLabel('Width (inches)', { exact: true }).fill('2.5');
	await page.getByLabel('Height (inches)', { exact: true }).fill('2');
	await page.getByRole('button', { name: 'Resize selected', exact: true }).click();
	await expect(page.locator('visio-viewer [data-geometry-error]')).toBeHidden();
	await expect(page.getByLabel('Width (inches)', { exact: true })).toHaveValue('');
	await page.getByRole('button', { name: 'Delete selected', exact: true }).click();
	await expect(rectangle).toHaveCount(0);
	await page
		.locator('visio-viewer .edit-controls')
		.getByRole('button', { name: 'Undo', exact: true })
		.click();
	await expect(rectangle).toHaveCount(1);
	const existing = page.locator('visio-viewer [data-shape-id="1"]');
	await existing.click();
	const before = await existing.getAttribute('transform');
	await page.getByLabel('Pin X (inches)').fill('5');
	await page.getByLabel('Pin Y (inches)').fill('6');
	await page.getByRole('button', { name: 'Move selected', exact: true }).click();
	await expect(existing).not.toHaveAttribute('transform', before!);
	await expect(page.locator('visio-viewer [data-geometry-error]')).toBeHidden();
	await page
		.locator('visio-viewer .edit-controls')
		.getByRole('button', { name: 'Undo', exact: true })
		.click();
	await expect(existing).toHaveAttribute('transform', before!);
});
