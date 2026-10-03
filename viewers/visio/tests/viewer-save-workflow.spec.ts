import { test, expect, type Page } from '@playwright/test';
import { taskPane } from './ribbon.js';
import JSZip from 'jszip';
import { createVsdxFixture } from './fixture.mjs';

async function downloadCopy(page: Page): Promise<Buffer> {
	const next = page.waitForEvent('download');
	await page.getByRole('button', { name: 'Download VSDX copy' }).click();
	const download = await next,
		stream = await download.createReadStream();
	const chunks: Buffer[] = [];
	for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
	return Buffer.concat(chunks);
}
async function open(page: Page, bytes: Buffer, name: string) {
	await page
		.locator('#file')
		.setInputFiles({ name, mimeType: 'application/vnd.ms-visio.drawing', buffer: bytes });
	await expect(page.locator('#file-name')).toHaveText(name);
}

test('geometry preserves an unapplied text draft and downloads/reloads actual edits with unknown parts', async ({
	page,
}) => {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Original text'));
	const preserved = Buffer.from([1, 9, 0, 255, 42]);
	zip.file('unknown/preserved.bin', preserved);
	const original = await zip.generateAsync({ type: 'nodebuffer' });
	await page.goto('/demo/?sample=1');
	await open(page, original, 'roundtrip.vsdx');
	await page.locator('visio-viewer [data-shape-id="1"]').click();
	await page.locator('visio-viewer .edit-controls summary').click();
	const text = page.getByLabel('Selected shape text', { exact: true });
	await text.fill('Unapplied text draft');
	await page.getByLabel('Pin X (inches)').fill('5');
	await page.getByLabel('Pin Y (inches)').fill('6');
	await page.getByRole('button', { name: 'Move selected', exact: true }).click();
	await expect(page.getByLabel('Pin X (inches)')).toHaveValue('');
	await expect(text).toHaveValue('Unapplied text draft');
	await expect(page.locator('visio-viewer svg text')).toContainText('Original text');
	await page
		.locator('visio-viewer .edit-controls')
		.getByRole('button', { name: 'Undo', exact: true })
		.click();
	await expect(page.locator('#edit-label')).toHaveText('ORIGINAL');
	await expect(text).toHaveValue('Unapplied text draft');
	await page
		.locator('visio-viewer .edit-controls')
		.getByRole('button', { name: 'Redo', exact: true })
		.click();
	await expect(page.locator('#edit-label')).toHaveText('EDITED COPY');
	await expect(text).toHaveValue('Unapplied text draft');
	await page.getByLabel('Width (inches)', { exact: true }).fill('4');
	await page.getByRole('button', { name: 'Apply text', exact: true }).click();
	await expect(page.getByLabel('Width (inches)', { exact: true })).toHaveValue('4');
	await expect(page.locator('visio-viewer [data-shape-id="1"] text')).toContainText(
		'Unapplied text draft',
	);
	await page.getByLabel('New rectangle ID').fill('42');
	await page.getByLabel('Rectangle text (optional)').fill('<literal rectangle>');
	await page.getByLabel('Pin X (inches)').fill('2');
	await page.getByLabel('Pin Y (inches)').fill('3');
	await page.getByLabel('Width (inches)', { exact: true }).fill('1');
	await page.getByLabel('Height (inches)', { exact: true }).fill('2');
	await page.getByRole('button', { name: 'Create rectangle', exact: true }).click();
	await expect(page.locator('visio-viewer [data-shape-id="42"]')).toHaveCount(1);
	const copy = await downloadCopy(page),
		reopened = await JSZip.loadAsync(copy);
	expect(await reopened.file('unknown/preserved.bin')!.async('nodebuffer')).toEqual(preserved);
	const xml = await reopened.file('visio/pages/page1.xml')!.async('string');
	expect(xml).toMatch(/N="PinX" V="5"/);
	expect(xml).toMatch(/N="PinY" V="6"/);
	await open(page, copy, 'reopened-geometry.vsdx');
	await expect(page.locator('visio-viewer [data-shape-id="42"]')).toHaveCount(1);
	await expect(page.locator('visio-viewer [data-shape-id="42"] text')).toContainText(
		'<literal rectangle>',
	);
	await expect(page.locator('visio-viewer [data-shape-id="1"] text')).toContainText(
		'Unapplied text draft',
	);
});

test('protected geometry refusal keeps draft, history and byte-exact original copy', async ({
	page,
}) => {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Protected shape'));
	const part = zip.file('visio/pages/page1.xml')!;
	zip.file(
		'visio/pages/page1.xml',
		(await part.async('string')).replace(
			'<Cell N="Width"',
			'<Cell N="LockMoveX" V="1"/><Cell N="Width"',
		),
	);
	const original = await zip.generateAsync({ type: 'nodebuffer' });
	await page.goto('/demo/?sample=1');
	await open(page, original, 'protected.vsdx');
	await page.locator('visio-viewer [data-shape-id="1"]').click();
	await page.locator('visio-viewer .edit-controls summary').click();
	const before = await page.locator('visio-viewer [data-shape-id="1"]').getAttribute('transform');
	await page.getByLabel('Selected shape text', { exact: true }).fill('Keep unapplied text');
	await page.getByLabel('Pin X (inches)').fill('5');
	await page.getByLabel('Pin Y (inches)').fill('6');
	await page.getByRole('button', { name: 'Move selected', exact: true }).click();
	await expect(page.locator('visio-viewer [data-geometry-error]')).toContainText(
		'EDIT_PROTECTED_CELL',
	);
	await expect(page.getByLabel('Pin X (inches)')).toHaveValue('5');
	await expect(page.getByLabel('Selected shape text', { exact: true })).toHaveValue(
		'Keep unapplied text',
	);
	await expect(page.locator('#edit-label')).toHaveText('ORIGINAL');
	await expect(
		page.locator('visio-viewer .edit-controls').getByRole('button', { name: 'Undo', exact: true }),
	).toBeDisabled();
	await expect(page.locator('visio-viewer [data-shape-id="1"]')).toHaveAttribute(
		'transform',
		before!,
	);
	expect(await downloadCopy(page)).toEqual(original);
	await page.getByRole('button', { name: 'Cancel', exact: true }).click();
	await expect(page.getByLabel('Pin X (inches)')).toHaveValue('');
	await expect(page.getByLabel('Selected shape text', { exact: true })).toHaveValue(
		'Protected shape',
	);
});

test('mobile geometry-only drafts can be cancelled without editing the document', async ({
	page,
}) => {
	await page.setViewportSize({ width: 390, height: 844 });
	await page.goto('/demo/?sample=1');
	await open(page, await createVsdxFixture(), 'mobile-draft.vsdx');
	await taskPane(page.locator('visio-viewer'), 'Inspector');
	await page.locator('visio-viewer .edit-controls summary').click();
	await page.getByLabel('Pin X (inches)').fill('4');
	await expect(page.getByRole('button', { name: 'Cancel', exact: true })).toBeEnabled();
	await page.getByRole('button', { name: 'Cancel', exact: true }).click();
	await expect(page.getByLabel('Pin X (inches)')).toHaveValue('');
	await expect(page.locator('#edit-label')).toHaveText('ORIGINAL');
});
