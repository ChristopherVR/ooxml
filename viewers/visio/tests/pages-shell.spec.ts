import { test, expect } from '@playwright/test';

for (const width of [1440, 390]) {
	for (const theme of ['light', 'dark']) {
		test(`Pages navigation and search work at ${width}px in ${theme}`, async ({ page }) => {
			await page.setViewportSize({ width, height: 844 });
			await page.addInitScript(
				(value) => localStorage.setItem('vitepress-theme-appearance', value),
				theme,
			);
			await page.goto('/');
			await expect(page.locator('html')).toHaveAttribute('data-theme', theme);
			const search = page.locator('.site-search-toggle');
			await search.click();
			await expect(page.locator('.site-search')).toBeVisible();
			await expect(page.locator('#site-search-input')).toBeFocused();
			await page.locator('#site-search-input').fill('zzzzunmatched');
			await expect(page.locator('.site-search-status')).toHaveText('No matching topics.');
			await page.locator('#site-search-input').fill('framework');
			await expect(page.locator('.site-search-results a').first()).toBeVisible();
			await page.keyboard.press('Escape');
			await expect(search).toBeFocused();
			await page.keyboard.press('Control+k');
			await expect(page.locator('.site-search')).toBeVisible();
			await page.getByRole('button', { name: 'Close search' }).click();
			await expect(page.locator('.site-search')).toBeHidden();
			if (width === 390) await page.locator('.menu-toggle').click();
			await page.getByRole('link', { name: 'Developer Guide', exact: true }).click();
			await expect(page.locator('.doc-content h1')).toHaveText('Getting started');
			if (width === 390) await page.locator('.menu-toggle').click();
			await page.locator('.nav-resources summary').click();
			await page.locator('.nav-resources').getByRole('link', { name: 'Architecture' }).click();
			await expect(page.locator('.doc-content h1')).toContainText('Architecture');
			expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(
				width,
			);
			await page.screenshot({ path: test.info().outputPath(`pages-${width}-${theme}.png`) });
		});
	}
}
