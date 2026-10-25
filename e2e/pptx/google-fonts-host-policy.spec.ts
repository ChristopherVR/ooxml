/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The host `googleFonts` option, driven through every binding by the demos'
 * `?googleFonts=0` / `?googleFontsBaseUrl=<url>` query options (see
 * `demos/pptx/shared/demo-root-options.ts`). The deck references "ADLaM
 * Display", a family that is neither installed nor embedded, so by default
 * the viewer requests it from Google Fonts.
 */
import { expect, test } from '@playwright/test';
import type { Page, Request } from '@playwright/test';

import { fixture, loadDeck } from './support/deck';

const DECK = fixture('adlam-webfont.pptx');
const MIRROR_ORIGIN = 'https://fonts.mirror.test';
const MIRROR = `${MIRROR_ORIGIN}/css2`;
const MIRROR_FONT_FILE = `${MIRROR_ORIGIN}/files/adlam-display.ttf`;
const FONT_CSS = '@font-face { font-family: "ADLaM Display"; src: local("Arial"); }';

function isGoogleFontsRequest(request: Request): boolean {
	const { hostname } = new URL(request.url());
	return hostname === 'fonts.googleapis.com' || hostname === 'fonts.gstatic.com';
}

/** Record (and answer offline) every request that reaches a Google Fonts host. */
async function recordGoogleRequests(page: Page): Promise<string[]> {
	const urls: string[] = [];
	await page.route(
		(url) => url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com',
		(route) => {
			urls.push(route.request().url());
			return route.fulfill({ contentType: 'text/css', body: FONT_CSS });
		},
	);
	page.on('request', (request) => {
		if (isGoogleFontsRequest(request) && !urls.includes(request.url())) {
			urls.push(request.url());
		}
	});
	return urls;
}

/** Give the asynchronous font resolution time to inject (or not inject) its link. */
async function settleFontResolution(page: Page): Promise<void> {
	await page.evaluate(() => document.fonts.ready);
	await page.waitForTimeout(500);
}

test.describe('googleFonts host option', () => {
	test('requests a missing family from Google Fonts by default', async ({ page }) => {
		const google = await recordGoogleRequests(page);
		await loadDeck(page, DECK);
		await expect(page.locator('head link[href^="https://fonts.googleapis.com/css2?"]')).toHaveCount(
			1,
		);
		await expect.poll(() => google.length).toBeGreaterThan(0);
	});

	test('googleFonts=false sends no request to Google', async ({ page }) => {
		const google = await recordGoogleRequests(page);
		await loadDeck(page, DECK, '/?googleFonts=0');
		await expect(page.locator('[data-pptx-element="true"]').first()).toBeVisible();
		await settleFontResolution(page);
		await expect(page.locator('head link[href*="fonts.googleapis.com"]')).toHaveCount(0);
		expect(google).toEqual([]);
	});

	test('a baseUrl mirror replaces Google for every font request', async ({ page }) => {
		const google = await recordGoogleRequests(page);
		const mirrored: string[] = [];
		// The glyph-outline fetch reads the stylesheet with fetch() and then fetches the font file
		// it names, so the mirror answers with CORS headers and a font file on its own origin.
		const cors = { 'Access-Control-Allow-Origin': '*' };
		await page.route(`${MIRROR_ORIGIN}/**`, (route) => {
			const url = route.request().url();
			mirrored.push(url);
			return url.startsWith(MIRROR)
				? route.fulfill({
						contentType: 'text/css',
						headers: cors,
						body: `@font-face { font-family: "ADLaM Display"; src: url(${MIRROR_FONT_FILE}) format("truetype"); }`,
					})
				: route.fulfill({ status: 404, headers: cors, body: '' });
		});
		await loadDeck(page, DECK, `/?googleFontsBaseUrl=${encodeURIComponent(MIRROR)}`);
		const link = page.locator(`head link[href^="${MIRROR}?"]`);
		await expect(link).toHaveCount(1);
		await expect(link).toHaveAttribute('href', /family=ADLaM%20Display/u);
		await expect.poll(() => mirrored.length).toBeGreaterThan(0);
		// The outline fetch followed the mirror's CSS to the mirror's font file.
		await expect.poll(() => mirrored.includes(MIRROR_FONT_FILE)).toBe(true);
		await settleFontResolution(page);
		expect(google).toEqual([]);
	});
});
