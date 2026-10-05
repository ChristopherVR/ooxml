import { describe, expect, it } from 'vitest';

import { at } from './indexed.js';
import { flattenSvgPath } from './svg-path-flatten.js';

describe('flattenSvgPath', () => {
	it('flattens a rectangle path (M L L L Z) to its 4 corners', () => {
		const loops = flattenSvgPath('M 0 0 L 100 0 L 100 50 L 0 50 Z');
		expect(loops).toHaveLength(1);
		const pts = at(loops, 0);
		expect(at(pts, 0)).toStrictEqual({ x: 0, y: 0 });
		expect(at(pts, 1)).toStrictEqual({ x: 100, y: 0 });
		expect(at(pts, 2)).toStrictEqual({ x: 100, y: 50 });
		expect(at(pts, 3)).toStrictEqual({ x: 0, y: 50 });
		// Z re-appends the subpath start.
		expect(at(pts, 4)).toStrictEqual({ x: 0, y: 0 });
	});

	it('handles relative commands identically to absolute ones', () => {
		const abs = flattenSvgPath('M 10 10 L 60 10 L 60 40 Z');
		const rel = flattenSvgPath('m 10 10 l 50 0 l 0 30 z');
		expect(rel).toStrictEqual(abs);
	});

	it('flattens H and V shorthand', () => {
		const loops = flattenSvgPath('M 0 0 H 40 V 20 H 0 Z');
		const pts = at(loops, 0);
		expect(pts).toStrictEqual([
			{ x: 0, y: 0 },
			{ x: 40, y: 0 },
			{ x: 40, y: 20 },
			{ x: 0, y: 20 },
			{ x: 0, y: 0 },
		]);
	});

	it('samples a cubic Bezier so the midpoint lands on the curve, not the chord', () => {
		// A quarter-circle-ish cubic from (0,0) to (100,100) via control points
		// that bow the curve well off the straight chord.
		const loops = flattenSvgPath('M 0 0 C 0 100 100 100 100 100', 8);
		const pts = at(loops, 0);
		const mid = at(pts, Math.floor(pts.length / 2));
		// The chord midpoint would be (50, 50); the true cubic midpoint bows
		// toward the control points, well above the chord's y at that x.
		expect(mid.y).toBeGreaterThan(60);
	});

	it('samples a full-circle arc (two semicircle A commands) close to a circle', () => {
		const r = 50;
		const d = `M ${-r} 0 A ${r} ${r} 0 1 1 ${r} 0 A ${r} ${r} 0 1 1 ${-r} 0`;
		const loops = flattenSvgPath(d, 32);
		const pts = at(loops, 0);
		expect(pts.length).toBeGreaterThan(20);
		for (const p of pts) {
			const dist = Math.hypot(p.x, p.y);
			expect(dist).toBeGreaterThan(r - 1);
			expect(dist).toBeLessThan(r + 1);
		}
	});

	it('starts a new loop (hole) on each additional M command', () => {
		const loops = flattenSvgPath('M 0 0 L 10 0 L 10 10 Z M 2 2 L 8 2 L 8 8 Z');
		expect(loops).toHaveLength(2);
		expect(at(at(loops, 0), 0)).toStrictEqual({ x: 0, y: 0 });
		expect(at(at(loops, 1), 0)).toStrictEqual({ x: 2, y: 2 });
	});

	it('returns an empty array for an empty or invalid path', () => {
		expect(flattenSvgPath('')).toStrictEqual([]);
	});

	it('ignores incomplete and nonfinite coordinates without leaking invalid points', () => {
		expect(flattenSvgPath('M')).toEqual([]);
		expect(flattenSvgPath('M 2')).toEqual([]);
		expect(flattenSvgPath('M 1e999 2')).toEqual([]);
		expect(flattenSvgPath('M 0 0 C 1 2 3 L 5 5 Z')).toEqual([
			[
				{ x: 0, y: 0 },
				{ x: 5, y: 5 },
				{ x: 0, y: 0 },
			],
		]);
	});

	it.each([0, -1, NaN, Infinity])(
		'uses safe sampling for invalid curve segment count %s',
		(segments) => {
			const path = 'M 0 0 Q 0 10 10 10';
			expect(flattenSvgPath(path, segments)).toEqual(flattenSvgPath(path));
		},
	);

	it('bounds excessive segment counts and preserves finite curve output', () => {
		const loops = flattenSvgPath('M 0 0 C 0 1 1 1 1 0', 1e12);
		expect(at(loops, 0)).toHaveLength(4097);
		expect(
			at(loops, 0).every((point) => Number.isFinite(point.x) && Number.isFinite(point.y)),
		).toBe(true);
	});
});
