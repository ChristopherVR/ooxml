/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Issue #141: a slide that advances by itself ("After 0s") must play its own entry transition.
 *
 * PowerPoint counts the "After" time from the end of the slide's transition. Four bindings armed
 * the timer when the slide was entered, so a 0s timing fired straight away: the next slide's
 * transition replaced the one that was still running and the slide appeared in a single frame (a
 * "pop"). React already waited for the transition.
 *
 * The deck is `issue-132-gradient-fill.pptx` with slide 3 (a slow push) set to advance after 0s.
 * The test measures how long slide 3 stays up before slide 4's text takes over the show stage.
 *
 * Run: bunx playwright test issue-141-timed-advance-after-transition
 */
import { readFileSync, writeFileSync } from 'node:fs';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { fixture, resetTabSession } from './support/deck';

const SHOW_STAGE = '[data-pptx-presenting]';
const SLIDE_3_TEXT = 'CONTENTS';
const SLIDE_4_TEXT = 'Review of work summary';
/** The push's own length (`spd="slow"`), the shortest the hold can be. */
const MIN_HOLD_MS = 800;

async function timedDeck(path: string): Promise<void> {
	const zip = await JSZip.loadAsync(readFileSync(fixture('issue-132-gradient-fill.pptx')));
	const slide3 = zip.file('ppt/slides/slide3.xml');
	expect(slide3, 'slide 3 must exist').not.toBeNull();
	const xml = await slide3!.async('string');
	expect(xml).toContain('advTm="3000"');
	zip.file('ppt/slides/slide3.xml', xml.replace('advTm="3000"', 'advTm="0"'));
	writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }));
}

test('a slide timed After 0s plays its entry transition before moving on', async ({
	page,
}, testInfo) => {
	test.setTimeout(120_000);
	const deck = testInfo.outputPath('timed.pptx');
	await timedDeck(deck);

	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(deck);
	await page.locator('[aria-label="Go to slide 5"]').first().waitFor({ timeout: 60_000 });
	await page.waitForTimeout(1200);
	await page.locator('[aria-label="Go to slide 2"]').first().click();
	await page.waitForTimeout(700);
	await page.keyboard.press('Shift+F5');
	await expect(page.locator(SHOW_STAGE).first()).toBeVisible();
	// Let the show settle on slide 2 (its own 3s timer has not fired yet).
	await page.waitForTimeout(400);

	// Poll inside the page: timestamps taken here are not skewed by the round trips.
	const handle = await page.evaluateHandle(
		([stage, three, four]) => {
			const seen: { three?: number; four?: number } = {};
			const start = performance.now();
			const timer = setInterval(() => {
				const text = (document.querySelector(stage) as HTMLElement | null)?.textContent ?? '';
				const now = performance.now() - start;
				if (seen.three === undefined && text.includes(three)) seen.three = now;
				if (seen.three !== undefined && seen.four === undefined && text.includes(four)) {
					seen.four = now;
				}
			}, 20);
			return { seen, stop: () => clearInterval(timer) };
		},
		[SHOW_STAGE, SLIDE_3_TEXT, SLIDE_4_TEXT] as const,
	);

	await page.keyboard.press('ArrowRight');
	await expect
		.poll(async () => handle.evaluate((h) => h.seen.four ?? null), { timeout: 15_000 })
		.not.toBeNull();
	const seen = await handle.evaluate((h) => {
		h.stop();
		return h.seen;
	});

	expect(seen.three, 'slide 3 must reach the stage').toBeDefined();
	const heldMs = seen.four! - seen.three!;
	expect(
		heldMs,
		`slide 3 held for ${Math.round(heldMs)}ms before slide 4 took over; its push is ~1000ms`,
	).toBeGreaterThanOrEqual(MIN_HOLD_MS);
});
