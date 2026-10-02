import { test, expect } from '@playwright/test';

test('shared Office theme changes update the embedded viewer without replacing its diagram', async ({
	page,
}) => {
	await page.goto('/');
	await page.evaluate(() => {
		localStorage.setItem('vitepress-theme-appearance', 'light');
		localStorage.setItem('visio-docs-theme', 'dark');
	});
	await page.reload();
	await page.getByRole('button', { name: 'Load live viewer' }).click();
	const frame = page.frameLocator('#live-viewer');
	await expect(frame.locator('html')).toHaveAttribute('data-theme', 'light');
	await frame.locator('visio-viewer select').selectOption('1');
	await frame.locator('visio-viewer [data-shape-id="a1"]').click();
	const svg = frame.locator('visio-viewer svg').first();
	await expect(svg).toHaveAttribute('aria-label', 'Architecture');
	await page.getByRole('button', { name: 'Switch to dark theme', exact: true }).click();
	await expect(frame.locator('html')).toHaveAttribute('data-theme', 'dark');
	const darkSurface = await frame
		.locator('visio-viewer .status')
		.evaluate((el) => getComputedStyle(el).backgroundColor);
	expect(darkSurface).toBe('rgb(23, 26, 30)');
	await expect(svg).toHaveAttribute('aria-label', 'Architecture');
	await expect(frame.locator('visio-viewer [data-shape-id="a1"]')).toHaveAttribute(
		'data-selected',
		'true',
	);
	await frame.getByRole('button', { name: 'Switch to light theme', exact: true }).click();
	await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
	await expect(svg).toHaveAttribute('aria-label', 'Architecture');
	const lightSurface = await frame
		.locator('visio-viewer .status')
		.evaluate((el) => getComputedStyle(el).backgroundColor);
	expect(lightSurface).not.toBe(darkSurface);
});

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
