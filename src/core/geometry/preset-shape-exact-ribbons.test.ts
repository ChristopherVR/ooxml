/**
 * Ribbons, curved ribbons and the left-right ribbon, measured on PowerPoint 16
 * (150 x 90 pt shape, defaults; the raster overlays match pixel for pixel).
 */

import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

function pointsOf(name: string, adj?: Record<string, number>): Array<[number, number]> {
	const svg = evaluatePresetShape(name, 150, 90, adj)!.svgPath;
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

describe('ribbon and ribbon2', () => {
	it('puts the ribbon tails at the top and the centre banner full height', () => {
		const pts = pointsOf('ribbon');
		// tail notch: wd8 across, half way down the tail (0..75)
		expect(has(pts, 18.75, 37.5)).toBe(true);
		// centre banner left edge x2 = 37.5 reaches the bottom curl
		expect(has(pts, 37.5, 86.25)).toBe(true);
	});

	it('puts the ribbon2 tails at the bottom', () => {
		const pts = pointsOf('ribbon2');
		expect(has(pts, 18.75, 52.5)).toBe(true);
		expect(has(pts, 37.5, 3.75)).toBe(true);
	});

	it('keeps both folds as separate darkenLess faces', () => {
		const faces = evaluatePresetShape('ribbon2', 150, 90)!.paths;
		expect(faces.filter((f) => f.fill === 'darkenLess')).toHaveLength(2);
	});
});

describe('ellipseRibbon and ellipseRibbon2', () => {
	it('hangs the tails from the parabolic top edge', () => {
		const pts = pointsOf('ellipseRibbon');
		// tail notch: on the sagging centre of the tail, 18.75 across
		expect(has(pts, 18.75, 38.65)).toBe(true);
		// the fold apex where the tail top edge meets the banner
		expect(has(pts, 56.25, 10.5)).toBe(true);
		// the banner bottom edge reaches the box bottom at the middle (control point
		// at hc, end points 3 pt higher)
		expect(Math.max(...pts.map((p) => p[1]))).toBeLessThanOrEqual(90.01);
	});

	it('turns ellipseRibbon over for ellipseRibbon2', () => {
		const a = pointsOf('ellipseRibbon');
		const b = pointsOf('ellipseRibbon2');
		expect(has(b, 18.75, 90 - 38.65)).toBe(true);
		expect(has(b, 56.25, 90 - 10.5)).toBe(true);
		expect(b).toHaveLength(a.length);
	});
});

describe('leftRightRibbon', () => {
	it('points the left arrow up by adj3 and the right arrow down', () => {
		const pts = pointsOf('leftRightRibbon');
		expect(has(pts, 0, 37.5)).toBe(true);
		expect(has(pts, 150, 52.5)).toBe(true);
		// right head spans 15..90, left head 0..75
		expect(has(pts, 105, 15)).toBe(true);
		expect(has(pts, 45, 75)).toBe(true);
	});

	it('rolls the left body into the right with a curl at the middle', () => {
		const pts = pointsOf('leftRightRibbon');
		// the curl ends on the right body top (y = 30) at hc
		expect(has(pts, 75, 30)).toBe(true);
		expect(has(pts, 75, 22.5)).toBe(true);
	});
});
