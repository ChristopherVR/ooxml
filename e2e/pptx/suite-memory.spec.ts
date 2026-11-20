/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Closing a PowerPoint document in the Office suite frees its viewer: decks with pictures, charts
 * and gradients are opened and closed repeatedly, and a heap snapshot after each close finds no
 * slide viewport left. Runs against a built suite (`bun run build:suite`, then
 * `bun run preview:suite`); set `PPTX_SUITE_URL` (for example http://localhost:8163) to enable
 * it, otherwise it is skipped.
 */
import { test, expect } from '@playwright/test';

import { heapReport } from '../../scripts/heap-snapshot.mjs';
import { fixture } from './support/deck';

const suite = process.env.PPTX_SUITE_URL;
const DECKS = ['bar-picture-fill-hill.pptx', 'chart-gallery.pptx'];
const CYCLES = 2;
const VIEWPORT = /^<\w+[^>]*data-pptx-viewport/;

test.skip(!suite, 'set PPTX_SUITE_URL to a served suite-dist');
test.use({ viewport: { width: 1440, height: 900 }, serviceWorkers: 'block' });

test('PowerPoint documents opened and closed repeatedly are collected', async ({ page }) => {
	test.setTimeout(300_000);
	await page.addInitScript(() => {
		(globalThis as { ooxmlSkipStart?: boolean }).ooxmlSkipStart = true;
	});
	await page.goto(`${suite}/?suite#/`);
	await page.locator('#create-row button[data-create]').first().waitFor();
	const session = await page.context().newCDPSession(page);
	const start = await heapReport(session, VIEWPORT);
	let references = 0;
	for (const deck of DECKS)
		for (let cycle = 1; cycle <= CYCLES; cycle++) {
			await page.locator('#file-input').setInputFiles(fixture(deck));
			await page.waitForFunction(
				() =>
					document.body.dataset.view === 'editor' &&
					document.querySelector('#status')?.textContent === 'Ready',
				null,
				{ timeout: 120_000 },
			);
			await expect(page.locator('[data-pptx-viewport]').first()).toBeVisible();
			await page.waitForTimeout(1500);
			references += await page
				.locator(
					'[data-pptx-viewport] [clip-path*="url("], [data-pptx-viewport] [fill*="url("], [data-pptx-viewport] [filter*="url("], [data-pptx-viewport] [mask*="url("]',
				)
				.count();
			const id = await page.evaluate(() => location.hash.split('/').pop());
			await page.locator(`[data-close="${id}"]`).click();
			await expect(page.locator('.document-host')).toHaveCount(0);
			await page.evaluate(() => (location.hash = '#/'));
			const report = await heapReport(session, VIEWPORT);
			test.info().annotations.push({
				type: 'heap',
				description: `${deck} close ${cycle}: ${report.count} closed viewports alive, ${report.detachedBytes} detached bytes, heap ${report.heapBytes - start.heapBytes} bytes over the start`,
			});
			expect.soft(report.count, `${deck}, close ${cycle}`).toBe(0);
		}
	expect(references).toBeGreaterThan(0);
});
