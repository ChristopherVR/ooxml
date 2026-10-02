import { test, expect } from '@playwright/test';

for (const viewport of [
	{ width: 1440, height: 1000 },
	{ width: 390, height: 844 },
]) {
	test(`landing loads the actual beta viewer at ${viewport.width}px`, async ({ page }) => {
		await page.setViewportSize(viewport);
		await page.goto('/');
		await expect(page.getByRole('heading', { level: 1 })).toContainText('.vsdx viewing,');
		await expect(page.locator('#live-viewer')).not.toHaveAttribute('src');
		await page.screenshot({ path: `test-results/landing-${viewport.width}.png`, fullPage: true });
		await page.getByRole('button', { name: 'Load live viewer' }).click();
		await expect(page.locator('#demo-status')).toContainText('Viewer ready.');
		await expect(page.frameLocator('#live-viewer').locator('visio-viewer')).toBeVisible();
		expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(
			true,
		);
	});
}
