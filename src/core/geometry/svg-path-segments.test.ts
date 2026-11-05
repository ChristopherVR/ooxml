import { describe, expect, it } from 'vitest';

import { arcToCubics, svgPathToSegments } from './svg-path-segments';
import { tokenizeSvgPath } from './svg-path-tokens';

describe('svgPathToSegments', () => {
	it('reads absolute and relative lines, H / V and close', () => {
		expect(svgPathToSegments('M10 10 h5 v5 H0 V0 l1 1 z m2 2 L4 4')).toEqual([
			{ type: 'moveTo', pt: { x: 10, y: 10 } },
			{ type: 'lineTo', pt: { x: 15, y: 10 } },
			{ type: 'lineTo', pt: { x: 15, y: 15 } },
			{ type: 'lineTo', pt: { x: 0, y: 15 } },
			{ type: 'lineTo', pt: { x: 0, y: 0 } },
			{ type: 'lineTo', pt: { x: 1, y: 1 } },
			{ type: 'close' },
			// after Z the current point is the subpath start (10, 10)
			{ type: 'moveTo', pt: { x: 12, y: 12 } },
			{ type: 'lineTo', pt: { x: 4, y: 4 } },
		]);
	});

	it('turns extra moveto pairs into lines and reads compact numbers', () => {
		expect(svgPathToSegments('M0,0 10,0-5.5.5')).toEqual([
			{ type: 'moveTo', pt: { x: 0, y: 0 } },
			{ type: 'lineTo', pt: { x: 10, y: 0 } },
			{ type: 'lineTo', pt: { x: -5.5, y: 0.5 } },
		]);
	});

	it('expands S and T against the previous control point', () => {
		const segs = svgPathToSegments('M0 0 C0 10 10 10 10 0 S20 -10 20 0 Q25 5 30 0 T40 0');
		expect(segs[2]).toEqual({
			type: 'cubicBezTo',
			pts: [
				{ x: 10, y: -10 },
				{ x: 20, y: -10 },
				{ x: 20, y: 0 },
			],
		});
		expect(segs[4]).toEqual({
			type: 'quadBezTo',
			pts: [
				{ x: 35, y: -5 },
				{ x: 40, y: 0 },
			],
		});
	});

	it('writes an arc as cubic Beziers that stay on the ellipse', () => {
		// half circle of radius 10 from (0,0) to (20,0), with compact flags
		const segs = svgPathToSegments('M0 0a10 10 0 0120 0');
		expect(segs.slice(1).every((s) => s.type === 'cubicBezTo')).toBe(true);
		expect(segs).toHaveLength(3);
		const cubic = (index: number) => {
			const segment = segs[index];
			if (segment?.type !== 'cubicBezTo') {
				throw new Error(`segment ${index} is not a cubic`);
			}
			return segment;
		};
		expect(cubic(2).pts[2]).toEqual({ x: 20, y: 0 });
		// sweep=1 from (0,0) to (20,0) about (10,0) passes through (10,-10)
		expect(cubic(1).pts[2].x).toBeCloseTo(10, 9);
		expect(cubic(1).pts[2].y).toBeCloseTo(-10, 9);
	});

	it('draws a zero-radius arc as a straight segment', () => {
		expect(arcToCubics({ x: 0, y: 0 }, 0, 5, 0, false, true, { x: 4, y: 0 })).toEqual([
			[
				{ x: 0, y: 0 },
				{ x: 4, y: 0 },
				{ x: 4, y: 0 },
			],
		]);
	});

	it('rejects a path that does not start with a command', () => {
		expect(() => svgPathToSegments('10 10')).toThrow(SyntaxError);
	});

	it('rejects a command with too few numbers, or an unknown character', () => {
		expect(() => svgPathToSegments('M0 0 L5')).toThrow(SyntaxError);
		expect(() => svgPathToSegments('M0 0 C1 1 2 2')).toThrow(SyntaxError);
		expect(() => svgPathToSegments('M0 0 L5 5 # 6 6')).toThrow(SyntaxError);
	});

	it('draws nothing for an arc to its own start point', () => {
		expect(svgPathToSegments('M5 5 A3 3 0 0 1 5 5')).toEqual([
			{ type: 'moveTo', pt: { x: 5, y: 5 } },
		]);
	});

	it('splits a full sweep into quarter-turn pieces that end on the ellipse', () => {
		// A large arc from (10,0) to (0,10) about the origin: three quarters of a circle.
		const pieces = arcToCubics({ x: 10, y: 0 }, 10, 10, 0, true, false, { x: 0, y: 10 });
		expect(pieces).toHaveLength(3);
		for (const [, , end] of pieces) {
			expect(Math.hypot(end.x, end.y)).toBeCloseTo(10, 9);
		}
		// The midpoint of each piece stays within 0.03 % of the radius.
		for (const [c1, c2, end] of pieces) {
			const index = pieces.findIndex((piece) => piece[2] === end);
			const start = index === 0 ? { x: 10, y: 0 } : (pieces[index - 1]?.[2] ?? { x: 0, y: 0 });
			const mx = (start.x + 3 * c1.x + 3 * c2.x + end.x) / 8;
			const my = (start.y + 3 * c1.y + 3 * c2.y + end.y) / 8;
			expect(Math.abs(Math.hypot(mx, my) - 10)).toBeLessThan(0.003);
		}
	});

	it('scales radii that are too small to reach the end point', () => {
		const pieces = arcToCubics({ x: 0, y: 0 }, 1, 1, 0, false, true, { x: 20, y: 0 });
		expect(pieces).toHaveLength(2);
		expect(pieces[1]?.[2]).toEqual({ x: 20, y: 0 });
		expect(pieces[0]?.[2].x).toBeCloseTo(10, 9);
	});
});

describe('tokenizeSvgPath', () => {
	it('reads compact arc flags and numbers that run together', () => {
		expect(tokenizeSvgPath('a1 1 0 011 1M0-1.5.5')).toEqual([
			{ type: 'a', values: [1, 1, 0, 0, 1, 1, 1] },
			{ type: 'M', values: [0, -1.5, 0.5] },
		]);
	});

	it('skips what it cannot read unless strict', () => {
		expect(tokenizeSvgPath('5 M1 1 # L2 2')).toEqual([
			{ type: 'M', values: [1, 1] },
			{ type: 'L', values: [2, 2] },
		]);
		expect(() => tokenizeSvgPath('5 M1 1', { strict: true })).toThrow(SyntaxError);
	});
});
