import { expect, test } from '@playwright/test';
import { demoReady, openLanding, pageErrors } from './helpers';

test('landing page offers browse, new workbook and the sample', async ({ page }) => {
	const errors = pageErrors(page);
	await openLanding(page);
	await expect(page).toHaveTitle(/XLSX Viewer/u);
	await expect(page.locator('#dropzone')).toBeVisible();
	await expect(page.locator('.landing-badge')).toHaveText('X');
	await expect(page.locator('#browse')).toHaveText('Browse files');
	await expect(page.locator('#blank')).toHaveText('or create a New Workbook');
	await expect(page.locator('#sample')).toHaveText('Open the sample workbook');
	await expect(page.locator('#landing-file')).toHaveAttribute('accept', '.xlsx,.xlsm,.xls,.csv');
	await expect(page.locator('#landing-error')).toBeHidden();
	await expect(page.locator('#workspace')).toBeHidden();
	await expect(page.locator('#build-stamp')).toHaveText(/xlsx-viewer demo · vanilla/u);
	expect(errors).toEqual([]);
});

test('the theme toggle follows and writes vitepress-theme-appearance', async ({ page }) => {
	await page.addInitScript(() => localStorage.setItem('vitepress-theme-appearance', 'dark'));
	await page.goto('/');
	await demoReady(page);
	await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
	await page.locator('#theme-toggle').click();
	await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
	expect(await page.evaluate(() => localStorage.getItem('vitepress-theme-appearance'))).toBe(
		'light',
	);
});

test('?locale= selects the interface language and the picker updates the URL', async ({ page }) => {
	await openLanding(page, 'vanilla', 'locale=de-DE');
	await expect(page.locator('#locale-select')).toHaveValue('de');
	await page.locator('#locale-select').selectOption('zh-CN');
	await expect(page).toHaveURL(/locale=zh-CN/u);
});

test('an unreadable file returns to the landing page with an error', async ({ page }) => {
	await openLanding(page);
	await page.locator('#landing-file').setInputFiles({
		name: 'broken.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		// Binary noise: text bytes would open as CSV, as the core sniffs content, not the name.
		buffer: Buffer.from([0x00, 0x01, 0x02, 0x03, 0xff, 0x00, 0x7f, 0x00]),
	});
	await expect(page.locator('#landing-error')).toBeVisible();
	await expect(page.locator('#landing')).toBeVisible();
});
