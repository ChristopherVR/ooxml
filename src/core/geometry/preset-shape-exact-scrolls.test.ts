/**
 * Scrolls and the funnel, measured on PowerPoint 16 (150 x 90 pt, defaults;
 * the raster overlays match the outline exactly).
 */

import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

function pointsOf(name: string, w = 150, h = 90): Array<[number, number]> {
	const svg = evaluatePresetShape(name, w, h)!.svgPath;
	const out: Array<[number, number]> = [];
	for (const m of svg.matchAll(/([MLAQ])((?: -?[\d.]+)+)/g)) {
		const n = m[2]!.trim().split(/\s+/).map(Number);
		out.push([n[n.length - 2]!, n[n.length - 1]!]);
	}
	return out;
}

function has(points: Array<[number, number]>, x: number, y: number, tol = 0.15): boolean {
	return points.some(([px, py]) => Math.abs(px - x) < tol && Math.abs(py - y) < tol);
}

describe('verticalScroll', () => {
	// ch = ss * 12.5% = 11.25, ch2 = 5.625
	it('rolls the top across the width with a round cap at the right', () => {
		const pts = pointsOf('verticalScroll');
		expect(has(pts, 144.375, 0)).toBe(true);
		expect(has(pts, 144.375, 11.25)).toBe(true);
		expect(has(pts, 16.875, 0)).toBe(true);
	});

	it('hangs the paper between x = ch and x = r - ch with a rounded bottom corner', () => {
		const pts = pointsOf('verticalScroll');
		expect(has(pts, 11.25, 5.625)).toBe(true);
		expect(has(pts, 138.75, 84.375)).toBe(true);
		expect(has(pts, 133.125, 90)).toBe(true);
	});

	it('shades the curl as a darkenLess face', () => {
		const faces = evaluatePresetShape('verticalScroll', 150, 90)!.paths;
		expect(faces.some((f) => f.fill === 'darkenLess')).toBe(true);
	});
});

describe('horizontalScroll', () => {
	it('is the vertical scroll turned about its diagonal', () => {
		const pts = pointsOf('horizontalScroll');
		expect(has(pts, 5.625, 11.25)).toBe(true);
		expect(has(pts, 0, 16.875)).toBe(true);
		expect(has(pts, 150, 5.625)).toBe(true);
		expect(has(pts, 138.75, 5.625)).toBe(true);
	});
});

describe('funnel', () => {
	it('keeps the outline inside the box and has the opening inset by ss / 20', () => {
		const pts = pointsOf('funnel');
		expect(Math.max(...pts.map((p) => p[1]))).toBeLessThanOrEqual(90.01);
		// the opening starts at its right end, inset ss / 20 = 4.5
		expect(has(pts, 145.5, 22.5, 0.2)).toBe(true);
	});

	it('cuts the opening out of the body as a counter-clockwise hole', () => {
		const svg = evaluatePresetShape('funnel', 150, 90)!.svgPath;
		// opening: radii 75 - 4.5 and 22.5 - 4.5, drawn with a negative sweep
		expect(svg).toMatch(/M 145\.5 22\.5 A 70\.5 18 0 [01] 0 /);
	});

	it('joins the rim to the bottom with sides tangent to both ellipses', () => {
		const pts = pointsOf('funnel');
		// left side: slope about 0.94 at the middle of the funnel
		const side = pts.filter(([x, y]) => y > 28 && y < 87 && x < 75);
		expect(side.length).toBeGreaterThan(0);
		const [x0, y0] = side[0]!;
		const [x1, y1] = side[side.length - 1]!;
		expect((x1 - x0) / (y1 - y0)).toBeGreaterThan(0.9);
		expect((x1 - x0) / (y1 - y0)).toBeLessThan(1.0);
	});
});
