// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';

import { clampScaleToCanvas, rasterizeElementClampedToCanvas } from './rasterize-element-clamped';

function fakeCanvas(width: number, height: number): HTMLCanvasElement {
	const canvas = document.createElement('canvas');
	canvas.width = width;
	canvas.height = height;
	return canvas;
}

describe('rasterizeElementClampedToCanvas', () => {
	afterEach(() => {
		vi.restoreAllMocks();
	});

	it('passes the requested scale through unchanged when it already fits the cap', async () => {
		const html2canvasFallback = vi.fn().mockResolvedValue(fakeCanvas(1920, 1080));
		const el = document.createElement('div');
		document.body.appendChild(el);

		const result = await rasterizeElementClampedToCanvas(el, 960, 540, document, {
			scale: 2,
			maxCanvasDim: 4096,
			mode: 'html2canvas',
			html2canvasFallback,
		});

		expect(result.clamped).toBeFalsy();
		expect(result.effectiveScale).toBe(2);
		expect(result.kind).toBe('canvas');
		document.body.removeChild(el);
	});

	it('reduces the scale to fit when the requested resolution exceeds the cap', async () => {
		const html2canvasFallback = vi
			.fn()
			.mockImplementation(async (_source, outputSize) =>
				fakeCanvas(outputSize.width, outputSize.height),
			);
		const el = document.createElement('div');
		document.body.appendChild(el);

		// natural 10000x10 wide element at 2x = 20000px, cap 4096.
		const result = await rasterizeElementClampedToCanvas(el, 10000, 10, document, {
			scale: 2,
			maxCanvasDim: 4096,
			mode: 'html2canvas',
			html2canvasFallback,
		});

		expect(result.clamped).toBeTruthy();
		expect(result.effectiveScale).toBeCloseTo(4096 / 10000, 5);
		expect(Math.max(result.width, result.height)).toBeLessThanOrEqual(4096);
		document.body.removeChild(el);
	});

	it('reduces both axes to fit the canvas area cap (iOS 4096 x 4096)', async () => {
		const html2canvasFallback = vi
			.fn()
			.mockImplementation(async (_source, outputSize) =>
				fakeCanvas(outputSize.width, outputSize.height),
			);
		const el = document.createElement('div');
		document.body.appendChild(el);

		// 1920 x 1080 at 4x is 7680 x 4320: under a 16384 edge, over 16,777,216 pixels.
		const result = await rasterizeElementClampedToCanvas(el, 1920, 1080, document, {
			scale: 4,
			maxCanvasDim: 16384,
			maxCanvasArea: 16_777_216,
			mode: 'html2canvas',
			html2canvasFallback,
		});

		expect(result.clamped).toBe(true);
		expect(result.width * result.height).toBeLessThanOrEqual(16_777_216);
		expect(result.width / result.height).toBeCloseTo(1920 / 1080, 2);
		document.body.removeChild(el);
	});

	it('never plans a rounded raster over the area cap', () => {
		for (const [w, h] of [
			[1000, 1000],
			[1333, 750],
			[4097, 4097],
		] as const) {
			const scale = clampScaleToCanvas(w, h, 5, 16384, 16_777_216);
			expect(Math.round(w * scale) * Math.round(h * scale)).toBeLessThanOrEqual(16_777_216);
		}
	});
});
