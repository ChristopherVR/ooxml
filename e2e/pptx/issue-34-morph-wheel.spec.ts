/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Issue #34: clicking round this deck's wheel, the wedges must stay in their
 * sectors through every morph.
 *
 * From the third topic on, the reporter watched the wheel's wedges swap
 * places mid-morph. Wedges of one orientation share a box and the unselected
 * ones share a fill, and the identical-twin matching pass paired each wedge
 * with the FIRST such twin in document order - its mirror across the wheel -
 * while its own in-place counterpart (recoloured, because the selection moved
 * there) stood unclaimed. The fix lets a closer counterpart the proximity pass
 * accepts outrank any distance-agnostic twin.
 *
 * This starts the show on the third topic (slide 4), clicks the wedge that
 * leads to slide 5 (the very wedge that used to fly to its mirror), freezes
 * the morph halfway and checks that every copy of every wedge still sits where
 * that wedge sits on the settled slide.
 *
 * Run: bunx playwright test issue-34-morph-wheel
 */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { test, expect } from '@playwright/test';
import type { Page } from '@playwright/test';

import { resetTabSession } from './support/deck';

const deck = resolve(fileURLToPath(new URL('./fixtures/solution-explorer.pptx', import.meta.url)));

/** The 5 MB deck embeds a video; give the parse room on CI. */
const LOAD_TIMEOUT_MS = 60_000;

/** The third topic on the wheel. */
const TOPIC_SLIDE = 4;

/** The eight wheel wedges of slide 4, in spTree order. */
const WEDGES = Array.from({ length: 8 }, (_, i) => `ppt/slides/slide4.xml-shape-${i}`);

/** The wedge hyperlinked to slide 5, on the right of the wheel. */
const NEXT_TOPIC_WEDGE = 'ppt/slides/slide4.xml-shape-3';

/**
 * The running show's own stage. The editor canvas and the slide rail stay
 * mounted behind the show and carry the SAME element ids at their own scale.
 */
const SHOW_STAGE = '[data-pptx-presenting]';

/** The transition's own layer, which not every binding mounts inside the stage. */
const MORPH_OVERLAY = '[data-pptx-transition-overlay]';

/** How far (CSS px) a copy may sit from its wedge's settled centre. */
const TOLERANCE_PX = 4;

interface Centre {
	x: number;
	y: number;
}

async function startShowOnTopicSlide(page: Page): Promise<void> {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(deck);
	await page.locator('[aria-label="Go to slide 14"]').first().waitFor({ timeout: LOAD_TIMEOUT_MS });
	await page.waitForTimeout(1200);
	await page.locator(`[aria-label="Go to slide ${TOPIC_SLIDE}"]`).first().click();
	await page.waitForTimeout(900);
	await page
		.getByRole('button', { name: /^present$|slide show/iu })
		.first()
		.click();
	await page.locator(SHOW_STAGE).first().waitFor({ timeout: LOAD_TIMEOUT_MS });
	await page.waitForFunction(
		([stage, id]) => {
			const rect = document
				.querySelector(`${stage} [data-element-id="${id}"]`)
				?.getBoundingClientRect();
			return Boolean(rect && rect.width > 0 && rect.height > 0);
		},
		[SHOW_STAGE, NEXT_TOPIC_WEDGE] as const,
		{ timeout: LOAD_TIMEOUT_MS },
	);
	await page.mouse.move(4, 4);
	await page.waitForTimeout(600);
}

/** Every wedge's centre on the settled slide, before the morph starts. */
async function settledCentres(page: Page): Promise<Record<string, Centre>> {
	return page.evaluate(
		([stage, ids]) => {
			const out: Record<string, { x: number; y: number }> = {};
			for (const id of ids) {
				const rect = document
					.querySelector(`${stage} [data-element-id="${id}"]`)
					?.getBoundingClientRect();
				if (rect) {
					out[id] = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
				}
			}
			return out;
		},
		[SHOW_STAGE, WEDGES] as const,
	);
}

/**
 * Click the wedge and hold the resulting morph open, paused, so it can be
 * scrubbed. The freeze is injected BEFORE the click (pausing afterwards is a
 * race), and the overlay's teardown timer is neutralised.
 */
async function clickWedgeAndFreezeMorph(page: Page): Promise<void> {
	await page.evaluate(() => {
		const real = window.setTimeout.bind(window);
		(window as unknown as { setTimeout: unknown }).setTimeout = (
			handler: TimerHandler,
			timeout?: number,
			...args: unknown[]
		) => (typeof timeout === 'number' && timeout >= 900 ? 0 : real(handler, timeout, ...args));
		const style = document.createElement('style');
		style.textContent = '*, *::before, *::after { animation-play-state: paused !important; }';
		document.head.appendChild(style);
	});
	// Click the wedge's own pixels away from the disc: two bindings navigate
	// from a pointer event, which a synthetic click misses.
	const spot = await page.evaluate(
		([stage, id]) => {
			const rect = document
				.querySelector(`${stage} [data-element-id="${id}"]`)
				?.getBoundingClientRect();
			return rect ? { x: rect.x + rect.width * 0.75, y: rect.y + rect.height * 0.5 } : null;
		},
		[SHOW_STAGE, NEXT_TOPIC_WEDGE] as const,
	);
	expect(spot, 'the next topic wedge must be on screen').not.toBeNull();
	await page.mouse.click(spot!.x, spot!.y);
	await page.locator(MORPH_OVERLAY).first().waitFor({ timeout: 5_000 });
	await page.mouse.move(4, 4);
	const morphAnimations = await page.evaluate(
		() =>
			document
				.getAnimations()
				.filter((animation) =>
					(animation as { animationName?: string }).animationName?.startsWith('pptx-morph'),
				).length,
	);
	expect(morphAnimations, 'a morph must be playing and frozen').toBeGreaterThan(0);
}

/** Put every running animation at the same fraction of its own duration. */
async function scrubTo(page: Page, fraction: number): Promise<void> {
	await page.evaluate(async (f) => {
		for (const animation of document.getAnimations()) {
			const duration = animation.effect?.getTiming().duration;
			animation.currentTime = typeof duration === 'number' ? duration * f : 0;
		}
		await new Promise<void>((painted) => {
			requestAnimationFrame(() => {
				requestAnimationFrame(() => {
					painted();
				});
			});
		});
	}, fraction);
	await page.waitForTimeout(120);
}

/**
 * The painted centre of every visible copy of a wedge right now. Four bindings
 * re-render a ghost as an ordinary slide element with its `data-element-id`;
 * React marks its overlay copies with `data-pptx-morph-outgoing` / `-lifted`.
 */
async function copyCentres(page: Page, elementId: string): Promise<Centre[]> {
	return page.evaluate(
		([stage, overlay, id]) => {
			const markers = [
				`[data-element-id="${id}"]`,
				`[data-pptx-morph-outgoing="${id}"]`,
				`[data-pptx-morph-lifted="${id}"]`,
			];
			const nodes = new Set<Element>();
			for (const scope of [stage, overlay]) {
				for (const marker of markers) {
					for (const node of document.querySelectorAll(`${scope} ${marker}`)) {
						nodes.add(node);
					}
				}
			}
			const out: { x: number; y: number }[] = [];
			for (const node of nodes) {
				// The node carrying the morph animation is not always the marked one,
				// and React's marked wrapper spans the whole slide: without an
				// animation, measure the first descendant boxed apart from it (the
				// element's own positioned container).
				const outer = node.getBoundingClientRect();
				const subtree = [node, ...node.querySelectorAll('*')];
				const painted =
					subtree.find((candidate) =>
						getComputedStyle(candidate).animationName.includes('pptx-morph'),
					) ??
					subtree.find((candidate) => {
						const box = candidate.getBoundingClientRect();
						return (
							box.width > 0 &&
							box.height > 0 &&
							(Math.abs(box.width - outer.width) > 1 || Math.abs(box.height - outer.height) > 1)
						);
					}) ??
					node;
				const rect = painted.getBoundingClientRect();
				if (rect.width > 0 && rect.height > 0) {
					out.push({ x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 });
				}
			}
			return out;
		},
		[SHOW_STAGE, MORPH_OVERLAY, elementId] as const,
	);
}

test.describe('issue #34 - morphing round the wheel', () => {
	test('every wedge stays in its sector halfway through the third topic morph', async ({
		page,
	}) => {
		test.setTimeout(120_000);
		await startShowOnTopicSlide(page);
		const settled = await settledCentres(page);
		expect(Object.keys(settled), 'every wedge must be on screen').toHaveLength(WEDGES.length);

		await clickWedgeAndFreezeMorph(page);
		await scrubTo(page, 0.5);

		for (const id of WEDGES) {
			const home = settled[id]!;
			const copies = await copyCentres(page, id);
			expect(copies.length, `${id} must still be painted mid-morph`).toBeGreaterThan(0);
			for (const copy of copies) {
				const drift = Math.hypot(copy.x - home.x, copy.y - home.y);
				expect(drift, `${id} must not leave its sector`).toBeLessThanOrEqual(TOLERANCE_PX);
			}
		}
	});
});
