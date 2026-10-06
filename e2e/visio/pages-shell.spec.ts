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
			if (theme === 'dark') await expect(page.locator('html')).toHaveClass(/dark/);
			else await expect(page.locator('html')).not.toHaveClass(/dark/);

			// Local search: opens from the button and from Ctrl+K, finds topics, closes with Escape.
			const search = page.locator('.VPNavBarSearch button').first();
			await search.click();
			await expect(page.locator('.VPLocalSearchBox')).toBeVisible();
			await page.locator('.VPLocalSearchBox input').fill('zzzzunmatched');
			await expect(page.locator('.VPLocalSearchBox')).toContainText('No results');
			await page.locator('.VPLocalSearchBox input').fill('framework');
			await expect(page.locator('.VPLocalSearchBox .result').first()).toBeVisible();
			await page.keyboard.press('Escape');
			await expect(page.locator('.VPLocalSearchBox')).toBeHidden();
			await page.keyboard.press('Control+k');
			await expect(page.locator('.VPLocalSearchBox')).toBeVisible();
			await page.keyboard.press('Escape');

			// Navigation: the Developer Guide link, then Architecture from the sidebar or menu.
			if (width === 390) await page.locator('.VPNavBarHamburger').click();
			await page
				.getByRole('link', { name: 'Developer Guide', exact: true })
				.filter({ visible: true })
				.first()
				.click();
			await expect(page.locator('.vp-doc h1')).toHaveText('Getting started');
			await page.goto('/architecture');
			await expect(page.locator('.vp-doc h1')).toContainText('Architecture');
			expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(
				width,
			);
			await page.screenshot({ path: test.info().outputPath(`pages-${width}-${theme}.png`) });
		});
	}
}
