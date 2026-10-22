import { expect, test } from '@playwright/test';

for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'])
	test(`${framework}: a theme scope inside an application shadow root reaches native blocks`, async ({
		page,
	}) => {
		await page.emulateMedia({ colorScheme: 'light' });
		await page.goto(`/native-blocks.html?framework=${framework}&shadow`);
		await expect(page.locator('output')).toHaveText('view:1');
		const toolbar = page.locator('[data-office-block="toolbar"]');
		await expect(toolbar).toHaveCSS('color', 'rgb(249, 250, 251)');
		const canvas = page.locator('[part="canvas"]');
		await expect(canvas).toHaveCSS('color', 'rgb(249, 250, 251)');
		await page.locator('#app').evaluate((scope) => {
			scope.style.setProperty('--office-font', 'Courier New, monospace');
			scope.dataset.officeTheme = 'light';
		});
		await expect(toolbar).toHaveCSS('color', 'rgb(31, 41, 55)');
		await expect(canvas).toHaveCSS('color', 'rgb(31, 41, 55)');
		await expect(toolbar).toHaveCSS('font-family', '"Courier New", monospace');
		await expect(canvas).toHaveCSS('font-family', '"Courier New", monospace');
	});
