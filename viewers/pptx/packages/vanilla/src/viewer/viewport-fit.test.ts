import { afterEach, describe, expect, it } from 'vitest';

import { fitViewerViewport } from './viewport-fit';

/** A viewport reporting what a browser does for a box laid out at a fractional size. */
function laidOut(width: number, height: number): HTMLElement {
	const viewport = document.createElement('div');
	Object.defineProperties(viewport, {
		clientWidth: { get: () => Math.round(width) },
		clientHeight: { get: () => Math.round(height) },
		offsetWidth: { get: () => Math.round(width) },
		offsetHeight: { get: () => Math.round(height) },
	});
	viewport.getBoundingClientRect = () => new DOMRect(0, 0, width, height);
	document.body.append(viewport);
	return viewport;
}

const canvas = { width: 960, height: 540 };

describe('fitViewerViewport', () => {
	afterEach(() => {
		document.body.replaceChildren();
	});

	it.each([
		[801.5, 700],
		[1200, 451.5],
	])('fits a %spx x %spx viewport without exceeding it', (width, height) => {
		const fit = fitViewerViewport(
			laidOut(width, height),
			canvas,
			{ fitPadding: 0, maxFitScale: null },
			false,
		);
		expect(canvas.width * fit).toBeLessThanOrEqual(width + 1e-9);
		expect(canvas.height * fit).toBeLessThanOrEqual(height + 1e-9);
		expect(fit).toBeCloseTo(Math.min(width / canvas.width, height / canvas.height), 9);
	});

	it('fits the presentation stage, which has no padding, inside a fractional viewport', () => {
		const viewport = laidOut(1400, 451.5);
		const fit = fitViewerViewport(viewport, canvas, {}, true);
		expect(parseFloat(viewport.style.padding)).toBe(0);
		expect(canvas.height * fit).toBeLessThanOrEqual(451.5 + 1e-9);
	});
});
