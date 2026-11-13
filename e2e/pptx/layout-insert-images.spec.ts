/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/** Layout artwork remains real image content after inserting or changing a slide's layout. */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import {
	buildUnusedLayoutImagesDeck,
	UNUSED_LAYOUT_PATH,
	USED_LAYOUT_PATH,
} from '../../src/core/pptx/__tests__/fixtures/unused-layout-images';
import { savePptxViaBackstage } from './save-pptx';
import {
	clearAutosaveSnapshots,
	loadDeck,
	openRibbonTab,
	resetTabSession,
	thumbnail,
	viewport,
} from './support/deck';
import { selectSlide } from './support/render-ready';

test.use({ viewport: { width: 1440, height: 900 } });

const MENU_TIMEOUT_MS = 15_000;
const USED_DESCRIPTION = 'Used Layout Logo';
const UNUSED_DESCRIPTION = 'Unused Layout Logo';
const UNUSED_LAYOUT_NAME = 'Unused Logo Layout';

function layoutEntries(page: Page): Locator {
	return page.locator('[data-testid="layout-gallery-menu"]:visible button');
}

async function closeMenu(page: Page): Promise<void> {
	await page.keyboard.press('Escape');
	await page.mouse.move(0, 0);
}

async function openLayoutMenu(page: Page): Promise<void> {
	await openRibbonTab(page, 'Home');
	const trigger = page.getByRole('button', { name: /^layout$/iu }).first();
	await trigger.hover();
	await trigger.click().catch(() => {
		// A hover-open menu can cover its trigger before the click lands.
	});
	await expect(layoutEntries(page).first()).toBeVisible({ timeout: MENU_TIMEOUT_MS });
}

/** Upload generated package bytes through the same input used to open a real deck. */
async function loadFixture(page: Page): Promise<void> {
	await resetTabSession(page);
	await page.goto('/');
	await clearAutosaveSnapshots(page);
	await page.locator('#file-input').setInputFiles({
		name: 'unused-layout-images.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: Buffer.from(await buildUnusedLayoutImagesDeck()),
	});
	await expect(page.getByRole('button', { name: /^Go to slide \d+$/u })).toHaveCount(1);
	await expectLayoutPicture(page, USED_LAYOUT_PATH, USED_DESCRIPTION, 1);
}

/** Only the editable model element can satisfy this locator, never a static gallery or rail. */
function layoutPicture(page: Page, layoutPath: string, description: string): Locator {
	const stem = layoutPath.slice(layoutPath.lastIndexOf('/') + 1).replace(/\.xml$/u, '');
	return viewport(page).locator(
		`[data-element-id^="layout-layout-pic-${stem}-"][aria-label="${description}"]`,
	);
}

async function expectLayoutPicture(
	page: Page,
	layoutPath: string,
	description: string,
	slideNumber: number,
): Promise<string> {
	await expect(thumbnail(page, slideNumber)).toHaveAttribute('aria-current', /^(true|page)$/u);
	const picture = layoutPicture(page, layoutPath, description);
	await expect(picture).toHaveCount(1);
	await expect(picture).toBeVisible();
	const image = picture.locator('img');
	await expect(image).toHaveCount(1);
	await expect(image).toBeVisible();
	await expect(image).toHaveAttribute('src', /\S/u);
	await expect
		.poll(() =>
			image.evaluate(
				(node) => node instanceof HTMLImageElement && node.complete && node.naturalWidth > 0,
			),
		)
		.toBe(true);
	await expect(picture.getByText('Image', { exact: true })).toBeHidden();
	const source = (await image.getAttribute('src'))!;
	// Static thumbnails deliberately strip model ids, so compare their decoded image source.
	await expect
		.poll(() =>
			thumbnail(page, slideNumber)
				.locator('img')
				.evaluateAll(
					(nodes, expected) =>
						nodes.some(
							(node) =>
								node instanceof HTMLImageElement &&
								node.getAttribute('src') === expected &&
								node.complete &&
								node.naturalWidth > 0,
						),
					source,
				),
		)
		.toBe(true);
	return source;
}

async function saveAndReopen(page: Page, path: string): Promise<void> {
	await (await savePptxViaBackstage(page)).saveAs(path);
	await loadDeck(page, path);
}

test('New Slide from an unused layout renders its picture and preserves it on reopen', async ({
	page,
}, testInfo) => {
	await loadFixture(page);
	const usedSource = await expectLayoutPicture(page, USED_LAYOUT_PATH, USED_DESCRIPTION, 1);
	await openRibbonTab(page, 'Home');
	await page
		.locator('[data-ribbon-control="home.slides.newSlide"] [data-pptx-chrome="split-caret"]')
		.click();
	// Gallery previews can hydrate media; cold direct-API access is covered separately in core.
	await layoutEntries(page).filter({ hasText: UNUSED_LAYOUT_NAME }).first().click();
	await closeMenu(page);
	await expect(page.getByRole('button', { name: /^Go to slide \d+$/u })).toHaveCount(2);
	const unusedSource = await expectLayoutPicture(page, UNUSED_LAYOUT_PATH, UNUSED_DESCRIPTION, 2);
	expect(unusedSource).not.toBe(usedSource);
	await expect(layoutPicture(page, USED_LAYOUT_PATH, USED_DESCRIPTION)).toHaveCount(0);

	await selectSlide(page, 1);
	await expectLayoutPicture(page, USED_LAYOUT_PATH, USED_DESCRIPTION, 1);
	await expect(layoutPicture(page, UNUSED_LAYOUT_PATH, UNUSED_DESCRIPTION)).toHaveCount(0);
	await selectSlide(page, 2);
	await expectLayoutPicture(page, UNUSED_LAYOUT_PATH, UNUSED_DESCRIPTION, 2);

	await saveAndReopen(page, testInfo.outputPath('inserted-layout-images.pptx'));
	await expect(page.getByRole('button', { name: /^Go to slide \d+$/u })).toHaveCount(2);
	await selectSlide(page, 2);
	await expectLayoutPicture(page, UNUSED_LAYOUT_PATH, UNUSED_DESCRIPTION, 2);
	await selectSlide(page, 1);
	await expectLayoutPicture(page, USED_LAYOUT_PATH, USED_DESCRIPTION, 1);
});

test('changing a slide to an unused layout renders its picture and preserves it on reopen', async ({
	page,
}, testInfo) => {
	await loadFixture(page);
	const usedSource = await expectLayoutPicture(page, USED_LAYOUT_PATH, USED_DESCRIPTION, 1);
	await openLayoutMenu(page);
	await layoutEntries(page).filter({ hasText: UNUSED_LAYOUT_NAME }).first().click();
	await closeMenu(page);
	await expect(page.getByRole('button', { name: /^Go to slide \d+$/u })).toHaveCount(1);
	const unusedSource = await expectLayoutPicture(page, UNUSED_LAYOUT_PATH, UNUSED_DESCRIPTION, 1);
	expect(unusedSource).not.toBe(usedSource);
	await expect(layoutPicture(page, USED_LAYOUT_PATH, USED_DESCRIPTION)).toHaveCount(0);

	await saveAndReopen(page, testInfo.outputPath('changed-layout-images.pptx'));
	await expect(page.getByRole('button', { name: /^Go to slide \d+$/u })).toHaveCount(1);
	await expectLayoutPicture(page, UNUSED_LAYOUT_PATH, UNUSED_DESCRIPTION, 1);
	await expect(layoutPicture(page, USED_LAYOUT_PATH, USED_DESCRIPTION)).toHaveCount(0);
});
