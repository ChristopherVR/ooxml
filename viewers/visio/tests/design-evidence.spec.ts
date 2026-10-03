import { test, expect } from '@playwright/test';

for (const theme of ['light', 'dark'] as const) {
	for (const size of [
		{ name: 'desktop', width: 1440, height: 1000 },
		{ name: 'mobile', width: 390, height: 844 },
	]) {
		test(`${size.name} ${theme} design evidence for workspace and documentation`, async ({
			page,
		}, testInfo) => {
			await page.setViewportSize({ width: size.width, height: size.height });
			await page.addInitScript((value) => localStorage.setItem('visio-docs-theme', value), theme);
			for (const route of [
				{ name: 'workspace', url: '/demo/?sample=1' },
				{ name: 'landing', url: '/' },
				{ name: 'guide', url: '/docs/index.html' },
			]) {
				await page.goto(route.url);
				await expect(page.locator('html')).toHaveAttribute('data-theme', theme);
				if (route.name === 'workspace')
					await expect(page.locator('visio-viewer .viewport > svg')).toBeVisible();
				await page.evaluate(() => document.fonts.ready);
				expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(
					size.width,
				);
				await page.screenshot({
					path: testInfo.outputPath(`${route.name}-${size.name}-${theme}.png`),
					fullPage: true,
				});
				if (route.name === 'landing') {
					await page.locator('#load-demo').click();
					const embedded = page.frameLocator('#live-viewer');
					await expect(embedded.locator('visio-viewer .viewport > svg')).toBeVisible();
					await expect(embedded.locator('html')).toHaveAttribute('data-theme', theme);
					await page
						.locator('#live-demo')
						.screenshot({ path: testInfo.outputPath(`live-demo-${size.name}-${theme}.png`) });
				}
			}
		});
	}
}

test('workspace appearance persists across navigation and embed mode is compact', async ({
	page,
}) => {
	await page.addInitScript(() => {
		if (!localStorage.getItem('vitepress-theme-appearance'))
			localStorage.setItem('vitepress-theme-appearance', 'dark');
	});
	await page.goto('/demo/?sample=1');
	await expect(page.locator('visio-viewer .viewport > svg')).toBeVisible();
	await page.getByRole('button', { name: 'Switch to light theme', exact: true }).click();
	await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
	await page.reload();
	await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
	await page.goto('/demo/?embed=1');
	await expect(page.locator('html')).toHaveAttribute('data-embedded', '');
	await expect(page.locator('.workspace-footer')).toBeHidden();
	await expect(page.getByRole('button', { name: 'Open .vsdx' })).toBeVisible();
});
