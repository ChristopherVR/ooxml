/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { savePptxViaBackstage } from './save-pptx';
import { loadDeck, ribbonTab, slideElements, thumbnail } from './support/deck';
import { downloadBytes } from './support/exports';
import { slidePosition } from './support/keyboard';
import { readZipPartText } from './support/pptx-xml';

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

/** The sample deck's theme, and the Facet theme a Themes pick writes. */
const OFFICE_THEME = { name: 'Office Theme', accent1: '4472C4', major: 'Calibri Light' };
const FACET_THEME = { name: 'Facet', accent1: '90C226', major: 'Trebuchet MS' };

/**
 * Save the deck through File > Save and read its theme part's name, accent 1 and heading face.
 * The File backstage closes after the save; the Design tab is put back on screen.
 */
async function savedTheme(page: Page): Promise<typeof OFFICE_THEME> {
	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const design = ribbonTab(page, 'Design');
	if ((await design.getAttribute('aria-selected')) !== 'true') {
		await design.click();
	}
	const xml = await readZipPartText(bytes, 'ppt/theme/theme1.xml');
	const read = (pattern: RegExp) => pattern.exec(xml)?.[1] ?? 'missing';
	return {
		name: read(/<a:theme\b[^>]*\bname="([^"]*)"/u),
		accent1: read(/<a:accent1>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/u).toUpperCase(),
		major: read(/<a:majorFont>\s*<a:latin typeface="([^"]*)"/u),
	};
}

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

	test('Undo takes a theme pick back on the tiles and in the saved file, Redo reapplies it', async ({
		page,
	}) => {
		await open(page, 'Design');
		const themes = page.locator('[data-ribbon-control="design.themes.gallery"]');
		const tile = (id: string) => themes.locator(`[data-gallery-item="${id}"]`).first();
		const undo = page.getByRole('button', { name: 'Undo', exact: true }).first();
		const redo = page.getByRole('button', { name: 'Redo', exact: true }).first();
		await expect(tile('office')).toHaveAttribute('aria-pressed', 'true');
		await tile('facet').click();
		await expect(tile('facet')).toHaveAttribute('aria-pressed', 'true');
		await expect(undo).toBeEnabled();
		await undo.click();
		// The undo stack held only the slides: the viewer's theme and the theme part a save writes
		// kept the undone pick.
		await expect(tile('office')).toHaveAttribute('aria-pressed', 'true');
		await expect(tile('facet')).toHaveAttribute('aria-pressed', 'false');
		await expect(undo).toBeDisabled();
		expect(await savedTheme(page)).toStrictEqual(OFFICE_THEME);
		await redo.click();
		await expect(tile('facet')).toHaveAttribute('aria-pressed', 'true');
		expect(await savedTheme(page)).toStrictEqual(FACET_THEME);
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
