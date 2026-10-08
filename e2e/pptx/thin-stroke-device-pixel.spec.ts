/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
/**
 * Issue #23: PowerPoint never paints a line thinner than one device pixel on
 * screen, so a 0.1 pt and a 1 pt rule both read as a solid hairline at fit
 * zoom. The slide stage is shrunk with `transform: scale()`, which used to
 * thin a sub-pixel stroke further until it faded or vanished.
 *
 * Every binding's stage now publishes one device pixel in slide px and paints
 * strokes as `max(<authored>, <one device pixel>)`, so this measures the
 * painted width of each line in DEVICE pixels: never under one, and a line
 * that is already wider keeps its authored width.
 */
import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';
import JSZip from 'jszip';
import { PptxHandler, ShapeBuilder } from 'pptx-viewer-core';

import { resetTabSession } from './support/deck';

/** Authored widths in slide px: 0.1 pt, 0.5 pt and 6 pt. */
const WIDTHS = [0.1 * (96 / 72), 0.5 * (96 / 72), 6 * (96 / 72)];
const LINE_LENGTH = 900;

async function thinLineDeck(): Promise<Buffer> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 1 });
	data.slides[0].elements = WIDTHS.map((width, idx) => {
		const line = ShapeBuilder.create('line')
			.position(120, 120 + idx * 120)
			.size(LINE_LENGTH, 0)
			.stroke({ color: '#000000', width })
			.build();
		return line;
	});
	data.slides[0].isDirty = true;
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();
	return zip.generateAsync({ type: 'nodebuffer' });
}

/**
 * Painted stroke width of each line in DEVICE pixels, with the on-screen scale
 * of its element (screen box over layout box: the stage's `scale()`).
 */
async function paintedDeviceWidths(page: Page): Promise<Array<{ scale: number; width: number }>> {
	return page.evaluate((lineLength) => {
		// The painted line only: not the transparent hit target, nor an editor-only
		// hover / selection decoration (`data-export-ignore`, invisible at rest).
		const visible = (node: Element): boolean => {
			const { stroke, strokeOpacity } = getComputedStyle(node);
			return (
				!node.closest('[data-export-ignore="true"]') &&
				Number.parseFloat(strokeOpacity) > 0 &&
				stroke !== 'none' &&
				stroke !== 'transparent' &&
				stroke !== 'rgba(0, 0, 0, 0)'
			);
		};
		const viewport = document.querySelector('[data-pptx-viewport]') ?? document;
		// The deck's only elements are the three 900px-wide lines.
		return [...viewport.querySelectorAll<HTMLElement>('[data-element-id]')]
			.filter((el) => Math.abs(el.offsetWidth - lineLength) < 1)
			.filter((el) => [...el.querySelectorAll('svg path, svg line')].some(visible))
			.sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top)
			.map((el) => {
				const scale = el.getBoundingClientRect().width / el.offsetWidth;
				const strokes = [...el.querySelectorAll('svg path, svg line')].filter(visible);
				const css = Math.max(
					0,
					...strokes.map((node) => Number.parseFloat(getComputedStyle(node).strokeWidth) || 0),
				);
				return { scale, width: css * scale * window.devicePixelRatio };
			});
	}, LINE_LENGTH);
}

test.describe('thin strokes stay at least one device pixel on screen (issue #23)', () => {
	test('a sub-pixel line is held at one device pixel; a wide one keeps its width', async ({
		page,
	}) => {
		await resetTabSession(page);
		await page.goto('/');
		await page.locator('#file-input').setInputFiles({
			name: 'thin-lines.pptx',
			mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			buffer: await thinLineDeck(),
		});
		// A flat line has a zero-height box, which Playwright counts as hidden.
		await page
			.locator('[data-pptx-viewport] [data-element-id]')
			.first()
			.waitFor({ state: 'attached' });

		const lines = await paintedDeviceWidths(page);
		expect(lines).toHaveLength(WIDTHS.length);
		const dpr = await page.evaluate(() => window.devicePixelRatio);
		for (const [idx, { scale, width }] of lines.entries()) {
			expect(scale).toBeGreaterThan(0);
			const authoredDevice = WIDTHS[idx]! * scale * dpr;
			// Never thinner than one device pixel...
			expect(width, `line ${idx}`).toBeGreaterThanOrEqual(0.99);
			// ...and never wider than the larger of the two.
			expect(width, `line ${idx}`).toBeLessThanOrEqual(Math.max(authoredDevice, 1) + 0.01);
			if (authoredDevice > 1) {
				expect(width, `line ${idx} keeps its authored width`).toBeCloseTo(authoredDevice, 1);
			}
		}
	});
});
