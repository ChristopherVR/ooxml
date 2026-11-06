/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Static slide stages and the public `SlideThumbnail` building block.
 *
 * 1. The slides pane draws its thumbnails without `data-element-id`, so every
 *    marked node on the page belongs to the editing canvas. Svelte and Vanilla
 *    used to leave the markers on their rail, and the rail sits ahead of the
 *    canvas in the DOM, so a `[data-element-id]` lookup found a thumbnail.
 * 2. A host that keeps its own strip (the custom-shell demo with `hostStrip=1`,
 *    one `SlideThumbnail` per slide fed from the slides pane's props) gets one
 *    thumbnail per slide at the deck's aspect ratio, without element ids, and
 *    each one follows an edit on the canvas before anything is saved.
 *
 * Run: bunx playwright test slide-thumbnail-strip
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { loadDeck, SAMPLE_DECK, slideElements, thumbnail } from './support/deck';
import { hostOwnedSessionUrl } from './support/host-owned-session';

test.use({ viewport: { width: 1440, height: 1000 } });

/** Every `data-element-id` value on the page that sits outside the editing canvas. */
function idsOffCanvas(page: Page): Promise<string[]> {
	return page.evaluate(() =>
		[...document.querySelectorAll('[data-element-id]')]
			.filter((node) => !node.closest('[data-pptx-viewport]'))
			.map((node) => node.getAttribute('data-element-id') ?? ''),
	);
}

const strip = (page: Page): Locator =>
	page.getByRole('navigation', { name: 'Host slide strip', exact: true });
const stripItems = (page: Page): Locator =>
	strip(page).getByRole('button', { name: /^Show slide \d+$/u });
const stripItem = (page: Page, slideNumber: number): Locator =>
	strip(page).getByRole('button', { name: `Show slide ${slideNumber}`, exact: true });

/** Open the custom-shell demo with the host-owned strip in place of the slides pane. */
async function openHostStrip(page: Page, room: string): Promise<void> {
	await page.goto(hostOwnedSessionUrl(room, { headless: '1', hostStrip: '1', sample: '1' }));
	await expect(page.getByLabel('Host session state', { exact: true })).toContainText(
		'Host: connected',
		{ timeout: 30_000 },
	);
	await expect(slideElements(page).first()).toBeVisible({ timeout: 30_000 });
	await expect(page.locator('[data-host-custom-shell]')).toBeVisible();
	await expect(
		page.getByRole('status', { name: 'Collaboration status', exact: true }),
	).toContainText('Editable', { timeout: 30_000 });
}

/** Replace the text of the canvas element showing `original`, then leave the editor. */
async function replaceText(page: Page, original: string, replacement: string): Promise<void> {
	await slideElements(page).filter({ hasText: original }).dblclick();
	const editor = page.locator('[data-inline-editor]').first();
	await expect(editor).toBeVisible();
	await editor.press('ControlOrMeta+A');
	await page.keyboard.type(replacement);
	await page.getByLabel('Host session state', { exact: true }).click();
}

test.describe('slide thumbnails', () => {
	test('the slides pane draws its thumbnails without element ids', async ({ page }) => {
		await loadDeck(page, SAMPLE_DECK);
		await expect(thumbnail(page, 7)).toBeVisible();
		await expect(slideElements(page).first()).toBeVisible();
		// The canvas keeps them (static stages drop theirs after their first paint).
		await expect.poll(() => idsOffCanvas(page)).toStrictEqual([]);
		expect(await page.locator('[data-pptx-viewport] [data-element-id]').count()).toBeGreaterThan(0);
	});

	test('a host-owned strip draws every slide from the live model', async ({ page }) => {
		test.setTimeout(90_000);
		await openHostStrip(page, `host-strip-${Date.now()}-${Math.random().toString(36).slice(2)}`);
		await expect(page.getByRole('navigation', { name: 'Slides', exact: true })).toHaveCount(0);

		// One thumbnail per slide, each as tall as its 160px width at the deck's aspect ratio.
		await expect(stripItems(page)).toHaveCount(7);
		const stage = await page
			.locator('[data-pptx-viewport] [aria-roledescription="slide"]')
			.first()
			.boundingBox();
		expect(stage).not.toBeNull();
		const expectedHeight = (160 * stage!.height) / stage!.width;
		for (const item of await stripItems(page).all()) {
			const box = await item.boundingBox();
			expect(box?.width).toBeCloseTo(160, 0);
			expect(Math.abs((box?.height ?? 0) - expectedHeight)).toBeLessThan(1.5);
		}
		await expect(stripItem(page, 1)).toContainText('Product Overview');

		// Static: no element ids anywhere but the editing canvas.
		await expect.poll(() => idsOffCanvas(page)).toStrictEqual([]);

		// The strip is the host's navigation.
		await stripItem(page, 2).click();
		await expect(stripItem(page, 2)).toHaveAttribute('aria-current', 'true');
		await expect(
			page.locator('[data-pptx-viewport] [data-element-id^="ppt/slides/slide2.xml"]').first(),
		).toBeVisible();
		await stripItem(page, 1).click();
		await expect(stripItem(page, 1)).toHaveAttribute('aria-current', 'true');

		// Live: an unsaved canvas edit reaches the slide's thumbnail.
		await replaceText(page, 'Product Overview', 'Strip follows edits');
		await expect(stripItem(page, 1)).toContainText('Strip follows edits');
		await expect(stripItem(page, 1)).not.toContainText('Product Overview');
		await expect.poll(() => idsOffCanvas(page)).toStrictEqual([]);
	});
});
