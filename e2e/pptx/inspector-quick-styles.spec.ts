/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The format inspector's Quick Styles are the shared gallery in every binding: a compact
 * `pptx-ui-ribbon-gallery` panel built by `buildQuickStylesGallery`, with the 18 presets in rows
 * of six that share the panel's width (so touch-size tiles never run past a narrow pane). A pick
 * must restyle the shape and save valid DrawingML: the shadow presets used to keep a CSS
 * `rgba()` colour, which the writer copied into `a:srgbClr/@val` as it was. Each binding keeps
 * its own caption, so the spec finds the gallery by its descriptor and picks through the DOM.
 *
 * Run: bunx playwright test inspector-quick-styles
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { elementWithText, inspector, loadDeck, selectElement } from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });

/** The Quick Styles gallery's tiles and how they sit in the panel, or null when there is none. */
function quickStylesState(page: Page) {
	return inspector(page).evaluate((root) => {
		const gallery = [
			...root.querySelectorAll<HTMLElement & { descriptor?: { id: string } }>(
				'pptx-ui-ribbon-gallery',
			),
		].find((candidate) => candidate.descriptor?.id === 'quickStyles');
		if (!gallery) {
			return null;
		}
		const panel = gallery.getBoundingClientRect();
		const tiles = [...gallery.querySelectorAll<HTMLButtonElement>('[data-gallery-item]')];
		const boxes = tiles.map((tile) => tile.getBoundingClientRect());
		return {
			panel: gallery.getAttribute('mode') === 'panel' && gallery.hasAttribute('compact'),
			ids: tiles.map((tile) => tile.dataset.galleryItem),
			rows: new Set(boxes.map((box) => Math.round(box.top))).size,
			inside: boxes.every((box) => box.left >= panel.left - 1 && box.right <= panel.right + 1),
			smallest: Math.round(Math.min(...boxes.flatMap((box) => [box.width, box.height]))),
		};
	});
}

/** The selected shape's `p:spPr` in the saved first slide. */
async function savedShapeProperties(page: Page, text: string): Promise<string> {
	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const shape = xml.split('<p:sp>').find((part) => part.includes(text)) ?? '';
	return /<p:spPr[\s>][\s\S]*?<\/p:spPr>/u.exec(shape)?.[0] ?? '';
}

test('the inspector Quick Styles offer every preset and save a valid shadow', async ({ page }) => {
	await loadDeck(page);
	await selectElement(page, elementWithText(page, 'Product Overview'));
	await expect(inspector(page)).toBeVisible();
	await expect.poll(() => quickStylesState(page)).toMatchObject({ panel: true, rows: 3 });
	const state = await quickStylesState(page);
	expect(state!.ids).toHaveLength(18);
	expect(state!.ids.slice(0, 2)).toStrictEqual(['blueFill', 'orangeFill']);
	expect(state!.inside).toBeTruthy();

	await inspector(page).evaluate((root) => {
		const gallery = [
			...root.querySelectorAll<HTMLElement & { descriptor?: { id: string } }>(
				'pptx-ui-ribbon-gallery',
			),
		].find((candidate) => candidate.descriptor?.id === 'quickStyles');
		gallery!.querySelector<HTMLButtonElement>('[data-gallery-item="blueShadow"]')!.click();
	});
	const spPr = await savedShapeProperties(page, 'Product Overview');
	expect(spPr).toMatch(/<a:solidFill><a:srgbClr val="4472C4"/u);
	expect(spPr).toMatch(
		/<a:outerShdw [^>]*><a:srgbClr val="000000"><a:alpha val="35000"(\/>|><\/a:alpha>)<\/a:srgbClr>/u,
	);
	// Every colour value is six hex digits (ST_HexColorRGB).
	for (const [, value] of spPr.matchAll(/<a:srgbClr val="([^"]*)"/gu)) {
		expect(value).toMatch(/^[0-9A-F]{6}$/iu);
	}
});

test('touch-size Quick Styles tiles stay inside a narrow inspector', async ({ page }) => {
	// Below 1024px the gallery gives its tiles the touch size; six of them must still fit the
	// side inspector, which keeps its desktop width there.
	await page.setViewportSize({ width: 900, height: 800 });
	await loadDeck(page);
	await selectElement(page, elementWithText(page, 'Product Overview'));
	await expect(inspector(page)).toBeVisible();
	await expect.poll(() => quickStylesState(page)).toMatchObject({ panel: true, rows: 3 });
	const state = await quickStylesState(page);
	expect(state!.inside).toBeTruthy();
	expect(state!.smallest).toBeGreaterThanOrEqual(24);
});
