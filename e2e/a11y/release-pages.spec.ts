/**
 * The release posts (`releases/<version>/`, built into suite-dist by build:suite): they follow
 * the device's colour scheme until the reader presses the appearance switch, the switch is a
 * named toggle that remembers the choice in `ooxml-suite-theme`, and the text tokens reach 4.5:1
 * in both themes.
 */
import { expect, test, type Page } from '@playwright/test';

const POSTS = ['2.0', '2.5'] as const;
const KEY = 'ooxml-suite-theme';

/** Contrast of text against the nearest painted background, as the browser resolved them. */
async function contrastOf(page: Page, selector: string): Promise<number> {
	return page
		.locator(selector)
		.first()
		.evaluate((element) => {
			const rgb = (value: string) =>
				(/rgba?\(([^)]+)\)/u.exec(value)?.[1] ?? '0, 0, 0')
					.split(',')
					.slice(0, 3)
					.map((part) => Number(part.trim()) / 255);
			const luminance = (value: string) => {
				const [r = 0, g = 0, b = 0] = rgb(value).map((c) =>
					c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4,
				);
				return 0.2126 * r + 0.7152 * g + 0.0722 * b;
			};
			let paint: Element | null = element;
			let background = 'rgb(255, 255, 255)';
			while (paint) {
				const fill = getComputedStyle(paint).backgroundColor;
				if (fill !== 'rgba(0, 0, 0, 0)' && fill !== 'transparent') {
					background = fill;
					break;
				}
				paint = paint.parentElement;
			}
			const [a, b] = [luminance(getComputedStyle(element).color), luminance(background)];
			return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
		});
}

for (const post of POSTS) {
	for (const scheme of ['light', 'dark'] as const) {
		test(`${post} follows a ${scheme} device and keeps text readable`, async ({ browser }) => {
			const context = await browser.newContext({ colorScheme: scheme });
			const page = await context.newPage();
			await page.goto(`/releases/${post}/`);
			await expect(page.locator('html')).toHaveAttribute('data-theme', scheme);
			const toggle = page.getByRole('button', { name: 'Dark mode' });
			await expect(toggle).toHaveAttribute('aria-pressed', String(scheme === 'dark'));
			for (const selector of [
				'.post-hero h1',
				'.post-meta',
				'.post-article p',
				'.post-article a[href]',
				'.post-toc a',
			])
				expect(await contrastOf(page, selector), `${selector} in ${scheme}`).toBeGreaterThanOrEqual(
					4.5,
				);
			await context.close();
		});
	}

	test(`${post}: the switch toggles the theme and remembers the choice`, async ({ browser }) => {
		const context = await browser.newContext({ colorScheme: 'light' });
		const page = await context.newPage();
		await page.goto(`/releases/${post}/`);
		const toggle = page.getByRole('button', { name: 'Dark mode' });
		await toggle.focus();
		await page.keyboard.press('Enter');
		await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
		await expect(toggle).toHaveAttribute('aria-pressed', 'true');
		expect(
			await page.evaluate((key) => JSON.parse(localStorage.getItem(key) ?? '{}').mode, KEY),
		).toBe('dark');
		await page.reload();
		await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
		await page.getByRole('button', { name: 'Dark mode' }).click();
		await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');
		await context.close();
	});
}

test('the 2.5 post links its migration guide and plays its recordings from posters', async ({
	page,
}) => {
	await page.goto('/releases/2.5/');
	await expect(page.locator('.release-intro a[href="migration.html"]')).toBeVisible();
	const scene = page.locator('[data-showcase="new-features"]');
	await expect(scene.getByRole('button', { name: /Mail merge/u })).toBeVisible();
	await expect(scene.locator('img[data-animation$=".gif"]')).toHaveCount(8);
	for (const img of await scene.locator('img[data-animation]').all())
		expect(await img.getAttribute('alt')).toBeTruthy();
	await page.goto('/releases/2.5/migration.html');
	await expect(
		page.getByRole('heading', { name: 'Word view modes have six values' }),
	).toBeVisible();
});
