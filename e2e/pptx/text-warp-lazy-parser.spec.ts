/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * `opentype.js` (the font-file parser behind WordArt envelope glyph outlines)
 * is loaded on demand, not with the viewer: a deck with no envelope WordArt
 * never requests it. `text-warp-envelope-parity.spec.ts` covers the outlines
 * an envelope deck still renders.
 *
 * Run: bunx playwright test text-warp-lazy-parser
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { loadDeck, slideStage } from './support/deck';

/**
 * Whether `url` loads the parser library: its package files (dev servers), a
 * pre-bundled dependency (`opentype__js.js`) or a built chunk named after it.
 * The renderer's own `text-warp-opentype-parser` module is not it.
 */
function isParserRequest(url: string): boolean {
	const { pathname } = new URL(url);
	const file = pathname.slice(pathname.lastIndexOf('/') + 1);
	return /\/opentype\.js[@/]/u.test(pathname) || /^opentype(?:_+js|[.-])/u.test(file);
}

function recordParserRequests(page: Page): string[] {
	const urls: string[] = [];
	page.on('request', (request) => {
		if (isParserRequest(request.url())) {
			urls.push(request.url());
		}
	});
	return urls;
}

test('a deck without envelope WordArt never requests the outline font parser', async ({ page }) => {
	const requests = recordParserRequests(page);
	await loadDeck(page);
	await expect(slideStage(page).locator('[data-pptx-element="true"]').first()).toBeVisible();
	await page.evaluate(() => document.fonts.ready);
	await page.waitForTimeout(500);
	expect(requests).toStrictEqual([]);
});
