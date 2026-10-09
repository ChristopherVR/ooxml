/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * Exact line spacing keeps its pitch whatever runs a line holds (issue #35).
 *
 * Every paragraph of `fixed-line-spacing.pptx` authors an exact 11pt line
 * spacing; the `Left` table cell and the `Box` text box open each 7pt line
 * with a 9pt "•" run, the `Mid` cell with a 7pt one. PowerPoint puts all three
 * blocks' baselines 11pt apart and level with each other. The browser used to
 * grow each line holding the 9pt run by a pixel, so by the fifth line the text
 * sat several points lower than in `Mid`.
 *
 * The spec reads the bottom edge of each line's first 7pt character from a DOM
 * range, so it makes no assumption about any binding's span structure.
 *
 * Fixture: `fixed-line-spacing.pptx` (`generate-fixed-line-spacing-fixture.ts`).
 *
 * Run: bunx playwright test fixed-line-spacing
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { FIXED_SPACING_BLOCKS } from './fixtures/generate-fixed-line-spacing-fixture';
import { fixture, loadDeckAt, slideStage } from './support/deck';
import { acrossFrameworks } from './support/parity';

test.use({ viewport: { width: 1440, height: 900 } });

const FIXTURE = fixture('fixed-line-spacing.pptx');

/** Half a CSS pixel: well under the whole pixel per line the bug added. */
const TOLERANCE_PX = 0.5;

/** The bottom edge of the first character of each of `lines`, on the main stage. */
async function lineBottoms(page: Page, lines: string[]): Promise<number[]> {
	return page.evaluate((wanted) => {
		// The main stage only: a filmstrip thumbnail renders the same text tiny.
		const root = document.querySelector('[data-pptx-viewport]') ?? document.body;
		const nodes: Text[] = [];
		const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
		for (let node = walker.nextNode(); node; node = walker.nextNode()) {
			nodes.push(node as Text);
		}
		const text = nodes.map((node) => node.data).join('');
		return wanted.map((line) => {
			let at = text.indexOf(line);
			if (at < 0) {
				throw new Error(`no rendered text holds ${JSON.stringify(line)}`);
			}
			const node = nodes.find((candidate) => {
				if (at < candidate.data.length) {
					return true;
				}
				at -= candidate.data.length;
				return false;
			});
			if (!node) {
				throw new Error(`lost ${JSON.stringify(line)}`);
			}
			const range = document.createRange();
			range.setStart(node, at);
			range.setEnd(node, at + 1);
			return range.getBoundingClientRect().bottom;
		});
	}, lines);
}

async function readSlide(page: Page, origin: string): Promise<Record<string, number[]>> {
	await loadDeckAt(page, origin, FIXTURE);
	await slideStage(page).waitFor();
	await page.waitForFunction(() => document.fonts.status === 'loaded');
	await page.waitForTimeout(400);
	const out: Record<string, number[]> = {};
	for (const block of FIXED_SPACING_BLOCKS) {
		out[block.name] = await lineBottoms(page, [...block.lines]);
	}
	return out;
}

test.describe('Fixed line spacing', () => {
	test('a larger run does not grow a line with exact spacing', async ({ browser }, testInfo) => {
		test.slow();
		const results = await acrossFrameworks(browser, testInfo, readSlide);

		const failures = results.flatMap(({ framework, value }) => {
			const mid = value.Mid;
			const problems: string[] = [];
			for (const name of ['Left', 'Box']) {
				const bottoms = value[name];
				bottoms.forEach((bottom, i) => {
					if (i === 0) {
						return;
					}
					const pitch = bottom - bottoms[i - 1];
					const midPitch = mid[i] - mid[i - 1];
					if (Math.abs(pitch - midPitch) > TOLERANCE_PX) {
						problems.push(
							`${name} line ${i + 1}: pitch ${pitch.toFixed(2)}px, Mid ${midPitch.toFixed(2)}px`,
						);
					}
				});
			}
			// The two cells share a row and its top edge, so their lines are level.
			value.Left.forEach((bottom, i) => {
				if (Math.abs(bottom - mid[i]) > TOLERANCE_PX) {
					problems.push(`Left line ${i + 1} sits ${(bottom - mid[i]).toFixed(2)}px off Mid's`);
				}
			});
			return problems.length > 0 ? [`${framework.name}: ${problems.join('; ')}`] : [];
		});

		expect(failures.join('\n')).toBe('');
	});
});
