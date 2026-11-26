/**
 * A PNG export is the slide at twice its own size, whatever the window.
 *
 * React and Angular sized the export from the on-screen stage's box, which the
 * viewer zooms with a CSS transform, so the same slide exported at 1400x788 in
 * a 1200 px window and 2232x1256 in a 1600 px one, while Vue, Svelte and
 * vanilla exported the 1280x720 slide at 2560x1440 every time.
 */
import { expect, test } from '@playwright/test';

import { loadDeck, SAMPLE_DECK } from './support/deck';
import { downloadBytes, downloadViaCard, openBackstageExport, PNG_CARD } from './support/exports';

/** Width and height from a PNG's IHDR chunk. */
function pngSize(bytes: Uint8Array): { width: number; height: number } {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	return { width: view.getUint32(16), height: view.getUint32(20) };
}

test('a PNG export keeps the slide resolution in any window size', async ({ browser }, info) => {
	const sizes = [];
	for (const viewport of [
		{ width: 1200, height: 800 },
		{ width: 1600, height: 1000 },
	]) {
		const page = await browser.newPage({ viewport, baseURL: info.project.use.baseURL });
		try {
			await loadDeck(page, SAMPLE_DECK);
			await openBackstageExport(page);
			sizes.push(pngSize(await downloadBytes(await downloadViaCard(page, PNG_CARD))));
		} finally {
			await page.close();
		}
	}
	// sample-deck.pptx is a 16:9 deck of 1280x720 CSS px; the export scale is 2.
	expect(sizes).toEqual([
		{ width: 2560, height: 1440 },
		{ width: 2560, height: 1440 },
	]);
});
