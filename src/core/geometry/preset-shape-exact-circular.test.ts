/**
 * The circular arrows follow the ECMA-376 guide chain: a band centred on the
 * ellipse `wd2 - thh`, whose arrow head tip sits `adj2` past the band end angle
 * `adj3`. Measured on PowerPoint 16 (150 x 90 pt; the raster overlays match the
 * outline pixel for pixel at the defaults and at scaled adjustments).
 */

import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

function pointsOf(name: string, adj?: Record<string, number>): Array<[number, number]> {
	const svg = evaluatePresetShape(name, 150, 90, adj)!.svgPath;
	const out: Array<[number, number]> = [];
	for (const m of svg.matchAll(/([MLA])((?: -?[\d.]+)+)/g)) {
		const n = m[2]!.trim().split(/\s+/).map(Number);
		out.push([n[n.length - 2]!, n[n.length - 1]!]);
	}
	return out;
}

function has(points: Array<[number, number]>, x: number, y: number, tol = 0.3): boolean {
	return points.some(([px, py]) => Math.abs(px - x) < tol && Math.abs(py - y) < tol);
}

describe('circular arrows', () => {
	// ss = 90, thh = 11.25, centre line radii 75 - 11.25 and 45 - 11.25.
	it('circularArrow ends its head on the centre line at the right', () => {
		expect(has(pointsOf('circularArrow'), 138.75, 45)).toBe(true);
	});

	it('circularArrow starts its tail on the outer and inner ellipse at 180 degrees', () => {
		const pts = pointsOf('circularArrow');
		// rw1 = 75 + 5.625 - 11.25, rw2 = rw1 - 11.25
		expect(has(pts, 75 - 69.375, 45)).toBe(true);
		expect(has(pts, 75 - 58.125, 45)).toBe(true);
	});

	it('leftCircularArrow runs the other way and tips at the right', () => {
		const pts = pointsOf('leftCircularArrow');
		expect(has(pts, 138.75, 45)).toBe(true);
		// the tail is at 180 degrees and the band passes through the bottom
		expect(has(pts, 75 - 69.375, 45)).toBe(true);
		expect(Math.min(...pts.map((p) => p[1]))).toBeGreaterThan(44.5);
	});

	it('leftRightCircularArrow has a tip at both ends of the band', () => {
		const pts = pointsOf('leftRightCircularArrow');
		expect(has(pts, 138.75, 45)).toBe(true);
		expect(has(pts, 11.25, 45)).toBe(true);
	});

	it('keeps the outline inside the shape box', () => {
		for (const name of ['circularArrow', 'leftCircularArrow', 'leftRightCircularArrow']) {
			for (const [x, y] of pointsOf(name)) {
				expect(x).toBeGreaterThan(-0.5);
				expect(x).toBeLessThan(150.5);
				expect(y).toBeGreaterThan(-0.5);
				expect(y).toBeLessThan(90.5);
			}
		}
	});
});
