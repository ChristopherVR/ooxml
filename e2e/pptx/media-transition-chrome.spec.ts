/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Issue #147: the mystery play icon inside a transition.
 *
 * `solution-explorer.pptx` slide 2 is a wheel painted over a full-bleed
 * background VIDEO (`!!Background`, a `p:pic` carrying `a:videoFile`), and
 * slide 3 arrives on a morph. The reporter caught a play triangle drifting
 * through that transition - faint, centred, and gone by the time the morph
 * landed.
 *
 * It came from the transition overlay, which renders the outgoing slide as a
 * STILL and deliberately passes no media map (a ghost must not mount a second
 * decoder for a video the live stage is already playing). The media renderer
 * therefore took its unplayable-media fallback: poster frame, plus the play
 * badge that goes with it on an authoring canvas. The badge is chrome, not
 * slide content, so a still must not paint it - nor the typed "Media"
 * placeholder box the other bindings would paint in the same spot.
 *
 * The rule now lives in the shared `mediaFallbackVisual`, and every binding
 * stamps `data-pptx-media-chrome` on whatever it paints for it, which is what
 * this spec looks for. The morph is frozen and scrubbed rather than sampled
 * live: the deck asks for `spd="slow"` (1s), far too tight to catch reliably.
 *
 * Run: bunx playwright test media-transition-chrome
 */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { test, expect } from '@playwright/test';
import type { Page } from '@playwright/test';

import { nextPaint } from './support/render-ready';
import {
	SHOW_STAGE,
	TRANSITION_OVERLAY,
	goToSlide,
	openDeckFile,
	startPresenting,
	stopPresenting,
} from './support/show';

const deck = resolve(fileURLToPath(new URL('./fixtures/solution-explorer.pptx', import.meta.url)));

/** The slide whose background is a video, and the morph it leaves on. */
const VIDEO_SLIDE = 2;

/** The background video itself, addressed by its core element id. */
const BACKGROUND_VIDEO = 'ppt/slides/slide2.xml-pic-0';

/** Neutral marker every binding puts on media chrome (shared MEDIA_CHROME_ATTRIBUTE). */
const MEDIA_CHROME = '[data-pptx-media-chrome]';

/** The shared play triangle (`MEDIA_FALLBACK_ICONS.play`), whatever wraps it. */
const PLAY_BADGE_PATH = 'M5 3 L19 12 L5 21 Z';

async function loadDeck(page: Page): Promise<void> {
	await openDeckFile(page, deck, 14);
}

/**
 * Enter the show on the video slide, assuming the deck is already loaded. Kept
 * separate from {@link loadDeck} so a test that samples the morph at several
 * fractions can re-enter the show without re-parsing the 5 MB deck each time:
 * doing the full load four times over ran Angular (the slowest binding, served
 * from `dist`) past the 60 s test timeout on CI.
 */
async function presentVideoSlide(page: Page): Promise<void> {
	await goToSlide(page, VIDEO_SLIDE);
	await startPresenting(page);
}

async function startShowOnVideoSlide(page: Page): Promise<void> {
	await loadDeck(page);
	await presentVideoSlide(page);
}

/**
 * Advance to the next slide and hold the resulting morph open, paused, at the
 * given fraction of its own clock. The overlay tears itself down on a single
 * long timer, so that timer is neutralised for the duration of the test.
 */
async function advanceAndFreezeMorph(page: Page, fraction: number): Promise<void> {
	await page.evaluate(() => {
		const real = window.setTimeout.bind(window);
		(window as unknown as { setTimeout: unknown }).setTimeout = (
			handler: TimerHandler,
			timeout?: number,
			...args: unknown[]
		) => (typeof timeout === 'number' && timeout >= 900 ? 0 : real(handler, timeout, ...args));
	});
	await page.keyboard.press('ArrowRight');
	// The morph is up once its overlay is mounted and its keyframes are running.
	await page.waitForFunction(
		(overlay) =>
			document.querySelector(overlay) !== null &&
			document.getAnimations().some((animation) => animation.playState === 'running'),
		TRANSITION_OVERLAY,
	);
	await page.evaluate((f) => {
		for (const animation of document.getAnimations()) {
			animation.pause();
			const duration = animation.effect?.getTiming().duration;
			animation.currentTime = typeof duration === 'number' ? duration * f : 0;
		}
	}, fraction);
	await nextPaint(page);
}

test.describe('issue #147 - media chrome must not ride along in a transition', () => {
	test('the outgoing background video paints no play badge mid-morph', async ({ page }) => {
		await startShowOnVideoSlide(page);

		// The live show really is playing the video, so the ghost below is the
		// only thing that could be painting a fallback.
		await expect(page.locator(`${SHOW_STAGE} video`).first()).toBeAttached();

		for (const fraction of [0.2, 0.5, 0.8]) {
			await advanceAndFreezeMorph(page, fraction);

			const overlay = page.locator(TRANSITION_OVERLAY).first();
			expect(
				await overlay.count(),
				`the transition overlay must be up at ${fraction * 100}% of the morph`,
			).toBeGreaterThan(0);

			expect(
				await page.locator(MEDIA_CHROME).count(),
				`media chrome was painted at ${fraction * 100}% of the morph`,
			).toBe(0);
			expect(
				await page.locator(`svg path[d="${PLAY_BADGE_PATH}"]`).count(),
				`a play triangle was painted at ${fraction * 100}% of the morph`,
			).toBe(0);

			// Re-enter the show on the video slide for the next sample. The deck is
			// already loaded, so only re-present it - re-parsing it each iteration
			// timed Angular out on CI.
			await stopPresenting(page);
			await presentVideoSlide(page);
		}
	});

	test('the outgoing video still paints its own frame, so the ghost is not a hole', async ({
		page,
	}) => {
		await startShowOnVideoSlide(page);
		await advanceAndFreezeMorph(page, 0.3);

		// Whatever each binding puts inside the ghost - a paused <video>, or the
		// poster frame the media map would have decoded - it must paint SOMETHING,
		// or suppressing the chrome would have traded one artefact for a hole.
		//
		// Searched three ways on purpose. Two bindings mount the overlay as a
		// SIBLING of the show stage rather than inside it, and the one that paints
		// per-shape ghosts labels them `data-pptx-morph-outgoing` instead of
		// re-emitting `data-element-id`, so any single selector finds no ghost at
		// all somewhere and the assertion goes vacuous.
		const painted = await page.evaluate(
			([overlay, stage, id]) =>
				[
					...document.querySelectorAll(
						`[data-pptx-morph-outgoing="${id}"], ${overlay} [data-element-id="${id}"], ${stage} [data-element-id="${id}"]`,
					),
				].some((node) => Boolean(node.querySelector('img, video, canvas'))),
			[TRANSITION_OVERLAY, SHOW_STAGE, BACKGROUND_VIDEO] as const,
		);
		expect(painted, 'the outgoing background video must still paint a frame').toBeTruthy();
	});
});
