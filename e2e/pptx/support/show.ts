/**
 * Slide-show helpers that wait for a state instead of a wall-clock sleep.
 *
 * The show specs used to load a deck, sleep 1.2 s, pick a slide, sleep 0.9 s,
 * press Present and sleep 2 s "for fullscreen and the opening build". Those
 * sleeps were most of each test's time on a fast machine and still too short
 * on a slow one. Each one stood for a condition this module waits for instead:
 * the deck is parsed, the slide is selected, the show stage is up and nothing
 * on it is still moving.
 */
import { expect, type Page } from '@playwright/test';

import { resetTabSession } from './deck';
import { nextPaint } from './render-ready';

/** The running show's own stage (the editor canvas stays mounted behind it). */
export const SHOW_STAGE = '[data-pptx-presenting]';

/** The transition overlay every binding marks while a transition plays. */
export const TRANSITION_OVERLAY = '[data-pptx-transition-overlay]';

/** Decks with embedded media take a while to parse on CI. */
const LOAD_TIMEOUT_MS = 60_000;

/** Frames in a row the show must hold still before it counts as settled. */
const STILL_FRAMES = 3;

/**
 * Open `file` from the landing dropzone and wait until its last slide's
 * thumbnail exists, which is when the deck has been parsed.
 */
export async function openDeckFile(page: Page, file: string, slideCount: number): Promise<void> {
	// Forget any restored session first, or the deck reopens and the landing
	// dropzone (the only place #file-input exists) never mounts.
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(file);
	await page
		.locator(`[aria-label="Go to slide ${slideCount}"]`)
		.first()
		.waitFor({ timeout: LOAD_TIMEOUT_MS });
}

/**
 * Select a slide in the rail, retrying the click until the rail reports it
 * current: a click that lands while the rail is still re-rendering after the
 * load is dropped by some bindings, which is what the old sleep papered over.
 */
export async function goToSlide(page: Page, slideNumber: number): Promise<void> {
	const thumb = page.locator(`[aria-label="Go to slide ${slideNumber}"]`).first();
	await expect(async () => {
		await thumb.click();
		await expect(thumb).toHaveAttribute('aria-current', /^(true|page)$/u, { timeout: 1_000 });
	}).toPass({ timeout: 15_000 });
	await nextPaint(page);
}

/**
 * Wait until the show is up and still: its stage exists, no transition overlay
 * is mounted, no finite animation is running, and the stage has kept the same
 * box for a few frames (entering full screen resizes it).
 *
 * Infinite animations (a looping build, a spinner) never finish, so they do not
 * count; paused ones (a spec that froze a transition on purpose) do not either.
 */
export async function waitForShowSettled(page: Page, timeout = 30_000): Promise<void> {
	await page.locator(SHOW_STAGE).first().waitFor({ timeout });
	await page.waitForFunction(
		([stageSelector, overlaySelector, stillFrames]) => {
			const scope = window as unknown as { __pptxShowSettle?: { key: string; frames: number } };
			const state = (scope.__pptxShowSettle ??= { key: '', frames: 0 });
			const moving = document.getAnimations().some((animation) => {
				if (animation.playState !== 'running') {
					return false;
				}
				const end = animation.effect?.getComputedTiming().endTime;
				return typeof end === 'number' && Number.isFinite(end);
			});
			const overlay = document.querySelector(overlaySelector) !== null;
			const key = [...document.querySelectorAll(stageSelector)]
				.map((node) => {
					const box = node.getBoundingClientRect();
					return `${box.x},${box.y},${box.width},${box.height}`;
				})
				.join('|');
			if (moving || overlay || key === '' || key !== state.key) {
				state.key = key;
				state.frames = 0;
				return false;
			}
			state.frames += 1;
			if (state.frames < stillFrames) {
				return false;
			}
			delete scope.__pptxShowSettle;
			return true;
		},
		[SHOW_STAGE, TRANSITION_OVERLAY, STILL_FRAMES] as const,
		{ polling: 'raf', timeout },
	);
}

/**
 * 1-based slide the show is on, read off the `ppt/slides/slideN.xml-…` element
 * ids of the largest painted stage (0 when it shows none). Slide counters
 * differ per binding; element ids do not.
 */
export async function shownSlideNumber(page: Page): Promise<number> {
	return page.evaluate(() => {
		const stages = [...document.querySelectorAll('[aria-roledescription="slide"]')]
			.filter((node) => node.getBoundingClientRect().width > 300)
			.sort((a, b) => b.getBoundingClientRect().width - a.getBoundingClientRect().width);
		for (const stage of stages) {
			const tally = new Map<number, number>();
			for (const node of stage.querySelectorAll('[data-element-id]')) {
				const match = (node.getAttribute('data-element-id') ?? '').match(/slide(\d+)\.xml/u);
				if (match) {
					tally.set(Number(match[1]), (tally.get(Number(match[1])) ?? 0) + 1);
				}
			}
			const top = [...tally.entries()].sort((a, b) => b[1] - a[1])[0];
			if (top) {
				return top[0];
			}
		}
		return 0;
	});
}

/**
 * Press `key` in a running show and wait until it shows slide `expected` with
 * its transition finished. The slide's own builds may still be playing: a spec
 * that needs them done calls {@link waitForShowSettled} as well.
 */
export async function advanceShow(page: Page, key: string, expected: number): Promise<void> {
	await page.keyboard.press(key);
	await expect
		.poll(() => shownSlideNumber(page), { message: `${key} shows slide ${expected}` })
		.toBe(expected);
	await expect(page.locator(TRANSITION_OVERLAY)).toHaveCount(0, { timeout: 15_000 });
	await nextPaint(page);
}

/**
 * Wait until the show stage is up on slide `slideNumber` with no transition
 * playing, without waiting for its builds. For decks whose slides advance on
 * their own timings, where the show never holds still.
 */
export async function waitForShowOn(page: Page, slideNumber: number): Promise<void> {
	await page.locator(SHOW_STAGE).first().waitFor({ timeout: 30_000 });
	await expect
		.poll(() => shownSlideNumber(page), { message: `the show is on slide ${slideNumber}` })
		.toBe(slideNumber);
	await expect(page.locator(TRANSITION_OVERLAY)).toHaveCount(0, { timeout: 15_000 });
	await nextPaint(page);
}

/** Press Present (or Slide Show) and wait for the show to settle. */
export async function startPresenting(page: Page): Promise<void> {
	await page
		.getByRole('button', { name: /^present$|slide show/iu })
		.first()
		.click();
	await waitForShowSettled(page);
}

/** Leave the show and wait until its stage is gone. */
export async function stopPresenting(page: Page): Promise<void> {
	await page.keyboard.press('Escape');
	await expect(page.locator(SHOW_STAGE)).toHaveCount(0, { timeout: 15_000 });
	await nextPaint(page);
}

/** Load a deck, park on a slide and start the show there, settled. */
export async function startShowOnSlide(
	page: Page,
	file: string,
	slideNumber: number,
	slideCount: number,
): Promise<void> {
	await openDeckFile(page, file, slideCount);
	await goToSlide(page, slideNumber);
	await startPresenting(page);
}
