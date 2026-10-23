/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { loadDeck, ribbonTab, slideElements, thumbnail } from './support/deck';
import { slidePosition } from './support/keyboard';

/**
 * Round 5 ribbon parity: the categorised Shapes gallery (Insert and Home), Slide Templates in the
 * New Slide menu, and Design > Themes / Variants / Designer. One framework-neutral spec, run on
 * every binding; it reads only `data-ribbon-*`, `data-insert-item`, `data-value` and
 * `data-gallery-item` hooks plus accessible names.
 */
test.use({ viewport: { width: 1440, height: 900 } });

const RECENT_KEY = 'pptx-ui.shapes.recent';
const CATEGORIES = [
	'Lines',
	'Rectangles',
	'Basic Shapes',
	'Block Arrows',
	'Equation Shapes',
	'Flowchart',
	'Stars and Banners',
	'Callouts',
	'Action Buttons',
];

async function open(page: Page, tab: string): Promise<void> {
	await loadDeck(page);
	await page.evaluate((key) => localStorage.removeItem(key), RECENT_KEY);
	await ribbonTab(page, tab).click();
	await page.waitForTimeout(200);
}

const headings = (menu: Locator) => menu.locator('.heading').allTextContents();

test.describe('Insert > Shapes', () => {
	test('is categorised like PowerPoint, with ScreenTips and Freeform in Lines', async ({
		page,
	}) => {
		await open(page, 'Insert');
		const menu = page.locator('[data-ribbon-control="insert.illustrations.shapes"]');
		await menu.getByRole('button', { name: 'Shapes', exact: true }).click();
		expect(await headings(menu)).toEqual(CATEGORIES);
		expect(await menu.locator('[data-insert-item]').count()).toBeGreaterThan(180);
		await expect(menu.locator('[data-insert-item="star5"]')).toHaveAttribute(
			'title',
			'Star: 5 Points',
		);
		await expect(menu.locator('[data-insert-item="wedgeRectCallout"]')).toHaveAttribute(
			'title',
			'Speech Bubble: Rectangle',
		);
		await expect(menu.locator('[data-insert-item="freeform:freeformShape"]')).toBeEnabled();
		await expect(menu.locator('[data-insert-item="freeform:curve"]')).toBeEnabled();
		// Scribble has no drawing tool yet: listed, disabled, with the reason as its ScreenTip.
		const scribble = menu.locator('[data-insert-item="freeform:scribble"]');
		await expect(scribble).toBeDisabled();
		await expect(scribble).toHaveAttribute('title', /Scribble[\s\S]*Not available/u);
	});

	test('inserts any preset and remembers it under Recently Used Shapes', async ({ page }) => {
		await open(page, 'Insert');
		const menu = page.locator('[data-ribbon-control="insert.illustrations.shapes"]');
		const trigger = menu.getByRole('button', { name: 'Shapes', exact: true });
		const before = await slideElements(page).count();
		await trigger.click();
		await menu.locator('[data-insert-item="wedgeRectCallout"]').click();
		await expect(slideElements(page)).toHaveCount(before + 1);
		await trigger.click();
		expect((await headings(menu))[0]).toBe('Recently Used Shapes');
		const recent = menu.locator('[data-insert-item="recent:wedgeRectCallout"]');
		await expect(recent).toHaveAttribute('title', 'Speech Bubble: Rectangle');
		await recent.click();
		await expect(slideElements(page)).toHaveCount(before + 2);
		expect(await page.evaluate((key) => localStorage.getItem(key), RECENT_KEY)).toBe(
			'["wedgeRectCallout"]',
		);
	});

	test('keeps working when storage is blocked', async ({ page }) => {
		await open(page, 'Insert');
		await page.evaluate(() => {
			Storage.prototype.getItem = () => {
				throw new Error('blocked');
			};
			Storage.prototype.setItem = () => {
				throw new Error('blocked');
			};
		});
		const menu = page.locator('[data-ribbon-control="insert.illustrations.shapes"]');
		const before = await slideElements(page).count();
		await menu.getByRole('button', { name: 'Shapes', exact: true }).click();
		await menu.locator('[data-insert-item="heart"]').click();
		await expect(slideElements(page)).toHaveCount(before + 1);
	});
});

test.describe('Home > Shapes', () => {
	test('lists the same categories and inserts the picked preset', async ({ page }) => {
		await open(page, 'Home');
		const menu = page.locator('[data-ribbon-control="home.drawing.shapes"]');
		await menu.getByRole('button', { name: 'Shapes', exact: true }).click();
		expect(await headings(menu)).toEqual(CATEGORIES);
		const star = menu.locator('[data-value="star5"]');
		await expect(star).toHaveAttribute('title', 'Star: 5 Points');
		// Freeform tools are drawn from Insert; Home lists them disabled and says so.
		const curve = menu.locator('[data-value="freeform:curve"]');
		await expect(curve).toBeDisabled();
		await expect(curve).toHaveAttribute('title', /Insert > Shapes/u);
		const before = await slideElements(page).count();
		await menu.locator('[data-value="mathPlus"]').click();
		await expect(slideElements(page)).toHaveCount(before + 1);
		await menu.getByRole('button', { name: 'Shapes', exact: true }).click();
		expect((await headings(menu))[0]).toBe('Recently Used Shapes');
	});
});

test.describe('Home > Slides', () => {
	test('has no Slide Templates button; the dialog opens from the New Slide menu', async ({
		page,
	}) => {
		await open(page, 'Home');
		await expect(
			page.locator(
				'[data-ribbon-group="home.slides"] > [data-ribbon-control="home.slides.slideTemplates"]',
			),
		).toHaveCount(0);
		await page
			.locator('[data-ribbon-control="home.slides.newSlide"] [data-pptx-chrome="split-caret"]')
			.click();
		const entry = page.getByRole('button', { name: 'Slide Templates...' });
		await expect(entry).toBeVisible();
		await entry.click();
		await expect(page.getByRole('dialog', { name: /Slide Templates/iu })).toBeVisible();
	});
});

test.describe('Insert > Slides', () => {
	test('the New Slide menu ends with Slide Templates, which opens the dialog', async ({ page }) => {
		await open(page, 'Insert');
		const menu = page.locator('[data-ribbon-control="insert.slides.newSlide"]');
		await menu.locator('.trigger').last().click();
		const items = menu.locator('[data-insert-item]');
		await expect(items.last()).toHaveText('Slide Templates...');
		await items.last().click();
		await expect(page.getByRole('dialog', { name: /Slide Templates/iu })).toBeVisible();
	});
});

test.describe('Design', () => {
	test('inline Themes gallery re-themes the deck and marks the current theme', async ({ page }) => {
		await open(page, 'Design');
		const themes = page.locator('[data-ribbon-control="design.themes.gallery"]');
		const tile = (id: string) => themes.locator(`[data-gallery-item="${id}"]`).first();
		await expect(tile('office')).toHaveAttribute('aria-pressed', 'true');
		await expect(tile('facet')).toHaveAttribute('title', 'Theme: Facet');
		await tile('facet').click();
		await expect(tile('facet')).toHaveAttribute('aria-pressed', 'true');
		await expect(tile('office')).toHaveAttribute('aria-pressed', 'false');
	});

	test('a theme pick keeps the active slide and every tile in place', async ({ page }) => {
		await open(page, 'Design');
		await thumbnail(page, 2).click();
		await expect.poll(() => slidePosition(page)).toMatch(/^Slide 2 of/u);
		const themes = page.locator('[data-ribbon-control="design.themes.gallery"]');
		const ids = () =>
			themes
				.locator('[data-gallery-item]')
				.evaluateAll((tiles) => tiles.map((tile) => tile.getAttribute('data-gallery-item')));
		const before = await ids();
		await themes.locator('[data-gallery-item="facet"]').first().click();
		await expect(themes.locator('[data-gallery-item="facet"]').first()).toHaveAttribute(
			'aria-pressed',
			'true',
		);
		// Re-theming used to reload the deck in React, jumping back to slide 1.
		expect(await slidePosition(page)).toMatch(/^Slide 2 of/u);
		expect(await ids()).toEqual(before);
	});

	test('Variants offers four tiles and Colors, Fonts, Effects and Background Styles', async ({
		page,
	}) => {
		await open(page, 'Design');
		const variants = page.locator('[data-ribbon-control="design.variants.gallery"]');
		await expect(variants.locator('[data-gallery-item]')).toHaveCount(4);
		const pressed = () =>
			variants
				.locator('[data-gallery-item][aria-pressed="true"]')
				.getAttribute('data-gallery-item');
		expect(await pressed()).toBe('original');
		await variants.locator('[data-gallery-item="inverted"]').first().click();
		await expect.poll(pressed).toBe('inverted');
		for (const id of ['colors', 'fonts']) {
			await expect(page.locator(`[data-ribbon-control="design.variants.${id}"]`)).toBeVisible();
		}
		// Effects and Background Styles are galleries too (design-effects-backgrounds.spec.ts).
		for (const id of ['effects', 'backgroundStyles']) {
			const control = page.locator(`[data-ribbon-control="design.variants.${id}"] button`).first();
			await expect(control).toBeEnabled();
		}
	});

	test('Designer is listed but disabled, with an honest tooltip', async ({ page }) => {
		await open(page, 'Design');
		const designer = page.locator('[data-ribbon-control="design.designer.designIdeas"]');
		await expect(designer).toBeVisible();
		const button = designer
			.locator('button')
			.first()
			.or(designer.and(page.locator('button')));
		await expect(button).toBeDisabled();
		await expect(designer).toHaveAttribute('title', /not available in this viewer/iu);
	});
});
