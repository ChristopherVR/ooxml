/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The format inspector's Text Warp picker is the shared gallery in every binding: a compact
 * `pptx-ui-ribbon-gallery` panel built by `buildTextWarpGallery`, offering all 39 OOXML presets
 * with the current one pressed. Picking a tile must mark it and save that real preset as
 * `a:prstTxWarp`. Each binding keeps its own section header (a collapsible button, `<details>`),
 * so the spec finds the gallery by its descriptor and picks through the DOM.
 *
 * Run: bunx playwright test inspector-text-warp
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import JSZip from 'jszip';

import { savePptxViaBackstage } from './save-pptx';
import { elementWithText, inspector, loadDeck, selectElement } from './support/deck';
import { downloadBytes } from './support/exports';

test.use({ viewport: { width: 1440, height: 900 } });

/** The Text Warp gallery's tile count and pressed tiles, or null when the inspector has none. */
function textWarpState(page: Page) {
	return inspector(page).evaluate((root) => {
		const gallery = [
			...root.querySelectorAll<HTMLElement & { descriptor?: { id: string } }>(
				'pptx-ui-ribbon-gallery',
			),
		].find((candidate) => candidate.descriptor?.id === 'textWarp');
		if (!gallery) {
			return null;
		}
		const tiles = [...gallery.querySelectorAll<HTMLButtonElement>('[data-gallery-item]')];
		return {
			panel: gallery.getAttribute('mode') === 'panel' && gallery.hasAttribute('compact'),
			count: tiles.length,
			pressed: tiles
				.filter((tile) => tile.getAttribute('aria-pressed') === 'true')
				.map((tile) => tile.dataset.galleryItem),
		};
	});
}

test('the inspector Text Warp gallery offers every preset and saves the pick', async ({ page }) => {
	await loadDeck(page);
	await selectElement(page, elementWithText(page, 'Product Overview'));
	await expect(inspector(page)).toBeVisible();
	await expect
		.poll(() => textWarpState(page))
		.toStrictEqual({ panel: true, count: 39, pressed: ['textNoShape'] });

	await inspector(page).evaluate((root) => {
		const gallery = [
			...root.querySelectorAll<HTMLElement & { descriptor?: { id: string } }>(
				'pptx-ui-ribbon-gallery',
			),
		].find((candidate) => candidate.descriptor?.id === 'textWarp');
		gallery!.querySelector<HTMLButtonElement>('[data-gallery-item="textArchUp"]')!.click();
	});
	await expect
		.poll(() => textWarpState(page))
		.toStrictEqual({ panel: true, count: 39, pressed: ['textArchUp'] });

	const bytes = await downloadBytes(await savePptxViaBackstage(page));
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	expect(xml).toMatch(/<a:prstTxWarp prst="textArchUp"/u);
});
