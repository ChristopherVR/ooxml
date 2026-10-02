import { test, expect } from '@playwright/test';

test('sample supports page navigation, zoom, selection and compatibility notes', async ({
	page,
}) => {
	await page.goto('/demo/');
	await expect(page.locator('visio-viewer svg')).toBeVisible();
	await expect(page.getByText('Sample workflow', { exact: true })).toBeVisible();
	await page.locator('visio-viewer select').selectOption('1');
	await expect(page.locator('visio-viewer svg')).toHaveAttribute('aria-label', 'Architecture');
	await page.locator('visio-viewer button').filter({ hasText: '100%' }).click();
	await expect(page.locator('visio-viewer output')).toHaveText('100%');
	await page.locator('visio-viewer [data-shape-id="a1"]').click();
	await expect(page.locator('#selection')).toContainText('Your framework');
	await expect(page.locator('#notes')).toContainText('Text metrics');
});
test('rejected inputs show an error and retain prior diagram', async ({ page }) => {
	await page.goto('/demo/');
	await page.locator('#file').setInputFiles({
		name: 'broken.vsdx',
		mimeType: 'application/octet-stream',
		buffer: Buffer.from('not a zip'),
	});
	await expect(page.locator('#error')).toBeVisible();
	await expect(page.locator('visio-viewer svg')).toHaveAttribute('aria-label', 'Release workflow');
	await page.getByRole('button', { name: 'Load sample' }).click();
	await expect(page.locator('#error')).toBeHidden();
});
test('mobile layout keeps open control, canvas and notes accessible', async ({ page }) => {
	await page.setViewportSize({ width: 390, height: 844 });
	await page.goto('/demo/');
	await expect(page.getByRole('button', { name: 'Open .vsdx' })).toBeVisible();
	await expect(page.locator('visio-viewer svg')).toBeVisible();
	expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
	await page.screenshot({ path: 'test-results/mobile-workspace.png', fullPage: true });
});
test('does not make remote requests while using sample', async ({ page }) => {
	const external: string[] = [];
	page.on('request', (request) => {
		if (!request.url().startsWith('http://127.0.0.1:4173') && !request.url().startsWith('data:'))
			external.push(request.url());
	});
	await page.goto('/demo/');
	await expect(page.locator('visio-viewer svg')).toBeVisible();
	expect(external).toEqual([]);
	await page.screenshot({ path: 'test-results/desktop-workspace.png', fullPage: true });
});

test('opens actual synthetic VSDX locally and safely renders literal document text', async ({
	page,
}) => {
	const { createVsdxFixture } = await import('./fixture.mjs');
	await page.goto('/demo/');
	await page.locator('#file').setInputFiles({
		name: 'sample.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('<script>literal text</script>'),
	});
	await expect(page.locator('#file-name')).toHaveText('sample.vsdx');
	await expect(page.locator('visio-viewer svg')).toHaveAttribute('aria-label', 'Imported page');
	await expect(page.locator('visio-viewer svg text')).toContainText(
		'<script>literal text</script>',
	);
	expect(await page.locator('visio-viewer svg script').count()).toBe(0);
	await expect(page.locator('#notes')).toContainText('cached values');
});

test('back and forward navigation leave a working viewer', async ({ page }) => {
	await page.goto('/demo/');
	await expect(page.locator('visio-viewer svg')).toBeVisible();
	await page.locator('a.brand').click();
	await expect(page).toHaveURL(/\/$/);
	await page.goBack();
	await expect(page.locator('visio-viewer svg')).toBeVisible();
	await page.locator('visio-viewer select').selectOption('1');
	await expect(page.locator('visio-viewer svg')).toHaveAttribute('aria-label', 'Architecture');
});

test('mobile primary controls have usable touch targets', async ({ page }) => {
	await page.setViewportSize({ width: 390, height: 844 });
	await page.goto('/demo/');
	for (const locator of [
		page.getByRole('button', { name: 'Open .vsdx' }),
		page.locator('visio-viewer button[data-action="fit"]'),
	]) {
		const box = await locator.boundingBox();
		expect(box?.height).toBeGreaterThanOrEqual(44);
		expect(box?.width).toBeGreaterThanOrEqual(44);
	}
});

test('explicit SVG download contains the current page and leaves selection unchanged', async ({
	page,
}) => {
	await page.goto('/demo/');
	await page.locator('visio-viewer select').selectOption('1');
	await page.locator('visio-viewer [data-shape-id="a1"]').click();
	const selection = await page.locator('#selection').textContent();
	const downloadEvent = page.waitForEvent('download');
	await page.getByRole('button', { name: 'Export SVG', exact: true }).click();
	const download = await downloadEvent;
	expect(download.suggestedFilename()).toBe('visio-page-2.svg');
	const stream = await download.createReadStream();
	const chunks: Buffer[] = [];
	for await (const chunk of stream!) chunks.push(Buffer.from(chunk));
	const svg = Buffer.concat(chunks).toString('utf8');
	expect(svg).toContain('Architecture');
	expect(svg).toContain('<metadata>');
	expect(svg).not.toContain('blob:');
	expect(svg).not.toContain('data-shape-id');
	expect(svg).not.toContain('tabindex');
	await expect(page.locator('#selection')).toHaveText(selection!);
	await expect(page.locator('visio-viewer svg')).toHaveAttribute('aria-label', 'Architecture');
	await expect(page.locator('#export-status')).toContainText('approximate snapshot');
});

test('isolated import renders the bounded embedded EMF subset and safely omits other records', async ({
	page,
}) => {
	const { createMetafileFixture } = await import('./metafile-fixture.mjs');
	const external: string[] = [];
	page.on('request', (request) => {
		if (!request.url().startsWith('http://127.0.0.1:4173') && !request.url().startsWith('data:'))
			external.push(request.url());
	});
	await page.goto('/demo/');
	await page.locator('#file').setInputFiles({
		name: 'vector.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createMetafileFixture(),
	});
	await expect(page.locator('visio-viewer [data-shape-id="emf"] svg path').first()).toBeVisible();
	await expect(page.locator('#notes')).toContainText('bounded EMF primitive subset');
	expect(external).toEqual([]);
	await page.locator('#file').setInputFiles({
		name: 'unsupported.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createMetafileFixture(true),
	});
	await expect(page.locator('#file-name')).toHaveText('unsupported.vsdx');
	await expect(page.locator('visio-viewer [data-shape-id="emf"] svg')).toHaveCount(0);
	await expect(page.locator('#notes')).toContainText('conversion subset');
	await page.getByRole('button', { name: 'Load sample' }).click();
	await expect(page.locator('visio-viewer svg')).toHaveAttribute('aria-label', 'Release workflow');
});
