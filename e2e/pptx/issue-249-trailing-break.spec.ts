/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { test, expect } from '@playwright/test';

import { resetTabSession } from './support/deck';

const fixturePath = resolve(
	fileURLToPath(new URL('./fixtures/issue-249-random-newlines.pptx', import.meta.url)),
);

/**
 * Issue 249: each paragraph of the sample ends in `<a:br>`, and PowerPoint draws an empty
 * line (the break's 18pt at 90% spacing = 19.44pt) under it. Distances between the
 * paragraphs' tops, in points, measured in PowerPoint 16 through COM
 * (`TextRange.Paragraphs(i).BoundTop`): 93.67, 171.43, 239.75, 308.07. Each paragraph after
 * the first starts 10pt (spcBef) below its box top, which a CSS margin puts outside the
 * paragraph's rect, so the rect-to-rect gaps are the COM gaps plus 10pt for the first step.
 */
const PARAGRAPH_GAPS_PT = [87.76, 68.32, 68.32];
const SHAPE_WIDTH_PT = 820.2802;
/**
 * Characters on each line, from `TextRange.Paragraphs(i).Lines(j).Length` in PowerPoint with
 * the deck's Aptos (the bullet and the trailing break are not counted). A browser without
 * Aptos paints a narrower stand-in, so the viewer spreads Aptos's measured advances over it.
 */
const LINE_LENGTHS = [
	[93, 98, 40],
	[99, 57],
	[86, 101],
	[98, 61],
];

test('issue 249: a trailing line break keeps its empty line, sized like PowerPoint', async ({
	page,
}) => {
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(fixturePath);
	await page.locator('[data-pptx-viewport] [data-element-id]').first().waitFor({ timeout: 15_000 });
	await page.waitForTimeout(1500);
	const measured = await page.evaluate((shapeWidthPt) => {
		const stage = document.querySelector('[data-pptx-viewport] [aria-roledescription="slide"]')!;
		const shape = stage.querySelector('[data-element-id]')!;
		const box = shape.getBoundingClientRect();
		// Vue, Svelte and Angular paragraphs are <p>; React and the vanilla binding use divs.
		const marked = Array.from(shape.querySelectorAll('p'));
		const paragraphs = marked.length > 0 ? marked : Array.from(shape.firstElementChild!.children);
		const tops = paragraphs.map((child) => child.getBoundingClientRect().top);
		const lineLengths = paragraphs.map((paragraph) => {
			const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
			const lines: number[] = [];
			let lastTop = Number.NEGATIVE_INFINITY;
			for (let node = walker.nextNode(); node; node = walker.nextNode()) {
				const text = node as Text;
				for (let index = 0; index < text.length; index++) {
					const code = text.data.codePointAt(index);
					// The bullet glyph is a marker, the zero-width space is the break's empty line.
					if (code === 0x200b || (text.parentElement?.className ?? '').includes('bullet')) continue;
					const range = document.createRange();
					range.setStart(text, index);
					range.setEnd(text, index + 1);
					const rect = range.getClientRects()[0];
					if (!rect || rect.width === 0) continue;
					if (Math.abs(rect.top - lastTop) > 5) {
						lines.push(0);
						lastTop = rect.top;
					}
					lines[lines.length - 1]! += 1;
				}
			}
			return lines;
		});
		return { scale: box.width / shapeWidthPt, tops, lineLengths };
	}, SHAPE_WIDTH_PT);
	expect(measured.tops).toHaveLength(4);
	PARAGRAPH_GAPS_PT.forEach((expectedPt, index) => {
		const gapPx = measured.tops[index + 1]! - measured.tops[index]!;
		expect(Math.abs(gapPx / measured.scale - expectedPt)).toBeLessThan(0.5);
	});
	expect(measured.lineLengths).toEqual(LINE_LENGTHS);
});
