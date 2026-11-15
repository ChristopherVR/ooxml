/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The opt-in cards look (`data-office-look="cards"` next to `data-office-theme`
 * on the page root): the thumbnail rail, canvas, notes and side pane float as
 * rounded, shadowed cards 12px apart on a workspace background, the side pane
 * is 320px and the Format pane headings are 36px tall. Light and dark use the
 * card shadows of the shared token contract. No pixel baselines: the computed
 * styles are asserted and the screenshots attached for review.
 *
 * Run: bunx playwright test cards-look
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { inspector, loadDeck, selectElement, slideElements } from './support/deck';
import { formatSection } from './support/format-pane';

test.use({ viewport: { width: 1440, height: 900 } });

const chrome = (page: Page, region: string): Locator =>
	page.locator(`[data-pptx-chrome="${region}"]`).first();

async function computed(target: Locator, property: string): Promise<string> {
	return target.evaluate((el, name) => getComputedStyle(el).getPropertyValue(name), property);
}

async function setLook(page: Page, theme: 'light' | 'dark', look: 'cards' | null): Promise<void> {
	await page.evaluate(
		([mode, value]) => {
			const root = document.documentElement;
			root.setAttribute('data-office-theme', mode!);
			if (value) root.setAttribute('data-office-look', value);
			else root.removeAttribute('data-office-look');
		},
		[theme, look] as const,
	);
}

for (const theme of ['light', 'dark'] as const) {
	test(`the cards look floats the editor chrome (${theme})`, async ({ page }, testInfo) => {
		await loadDeck(page);
		await setLook(page, theme, null);
		const slides = chrome(page, 'slides');
		const flatRadius = await computed(slides, 'border-top-left-radius');

		await setLook(page, theme, 'cards');
		for (const region of ['slides', 'inspector', 'notes'])
			await expect
				.poll(() => computed(chrome(page, region), 'border-top-left-radius'))
				.toBe('14px');
		await expect
			.poll(() => computed(page.locator('[data-pptx-viewport]').first(), 'border-top-left-radius'))
			.toBe('14px');
		expect(flatRadius).not.toBe('14px');
		const shadow = await computed(slides, 'box-shadow');
		expect(shadow).toContain(theme === 'dark' ? 'rgba(0, 0, 0, 0.4)' : 'rgba(60, 64, 67, 0.2)');
		expect(await computed(chrome(page, 'body'), 'column-gap')).toBe('12px');

		// 12px between the rail and the canvas, whatever splitter sits between them.
		const rail = (await slides.boundingBox())!;
		const canvas = (await page.locator('[data-pptx-viewport]').first().boundingBox())!;
		expect(Math.round(canvas.x - (rail.x + rail.width))).toBe(12);
		expect(Math.round(rail.x)).toBe(12);
		expect((await inspector(page).boundingBox())!.width).toBeGreaterThanOrEqual(320);

		await selectElement(page, slideElements(page).filter({ hasText: 'Product Overview' }).first());
		const section = formatSection(page, 'font').or(formatSection(page, 'fill')).first();
		await section.waitFor();
		const heading = section.locator('summary[part="heading"]');
		await expect.poll(() => computed(heading, 'min-height')).toBe('36px');
		expect(await computed(heading, 'font-size')).toBe('13px');
		await testInfo.attach(`cards-${theme}`, {
			body: await page.screenshot(),
			contentType: 'image/png',
		});

		// Switching the look back keeps the editor and its selection.
		await setLook(page, theme, null);
		await expect.poll(() => computed(slides, 'border-top-left-radius')).toBe(flatRadius);
		await expect(section).toBeVisible();
	});
}
