import { describe, expect, it } from 'vitest';

import { closedWarpPaths } from './text-warp-closed-paths';
import { hasGlyphEnvelope } from './text-warp-envelope-curves';
import { createEnvelopeWarp } from './text-warp-envelope-map';
import { buildWarpedGlyphOutlinePathD } from './text-warp-glyph-outline';
import { clipPolygonByX, clipPolygonConvex, ellipsePolygon } from './text-warp-polygon-clip';

const W = 200;
const H = 100;
const BLOCK = { left: 0, top: 0, right: 100, bottom: 40 };

function mid<T>(points: T[]): T {
	return points[Math.floor(points.length / 2)]!;
}

describe('closed-path presets', () => {
	it('render through the per-glyph envelope', () => {
		for (const preset of ['textRingInside', 'textRingOutside', 'textCirclePour']) {
			expect(hasGlyphEnvelope(preset)).toBe(true);
		}
	});

	it('Ring Inside: two offset ellipses, the text starting on the upper arc', () => {
		const paths = closedWarpPaths('textRingInside', W, H, undefined, undefined)!;
		// adj defaults to 60%: each ellipse is 60 tall, touching the top or the bottom.
		expect(paths.top[0]![0]).toEqual({ x: 0, y: 30 });
		expect(mid(paths.top[0]!).x).toBeCloseTo(100, 6);
		expect(mid(paths.top[0]!).y).toBeCloseTo(0, 6);
		// The glyph bottoms reach the top of the lower ellipse (40 down).
		expect(mid(paths.bottom[0]!).y).toBeCloseTo(40, 6);
		// The second half runs back along the bottom of the lower ellipse.
		expect(mid(paths.top[1]!).y).toBeCloseTo(100, 6);
		expect(paths.top[1]![0]!.x).toBeCloseTo(W, 6);
		expect(paths.clips).toHaveLength(2);
	});

	it('Ring Outside starts on the lower band with the glyph tops toward the middle', () => {
		const paths = closedWarpPaths('textRingOutside', W, H, undefined, undefined)!;
		expect(mid(paths.top[0]!).y).toBeCloseTo(60, 6);
		expect(mid(paths.bottom[0]!).y).toBeCloseTo(100, 6);
		expect(mid(paths.top[1]!).y).toBeCloseTo(40, 6);
		expect(mid(paths.bottom[1]!).y).toBeCloseTo(0, 6);
	});

	it('pins the ring adjust to 50%..99%', () => {
		const low = closedWarpPaths('textRingInside', W, H, 10000, undefined)!;
		expect(mid(low.bottom[0]!).y).toBeCloseTo(50, 6);
		const high = closedWarpPaths('textRingInside', W, H, 99999, undefined)!;
		expect(mid(high.bottom[0]!).y).toBeCloseTo(1, 6);
	});

	it('Circle Pour: the box ellipse and its adj2 scale, starting at the left tip', () => {
		const paths = closedWarpPaths('textCirclePour', W, H, undefined, undefined)!;
		expect(paths.top).toHaveLength(1);
		expect(paths.top[0]![0]!.x).toBeCloseTo(0, 6);
		expect(paths.top[0]![0]!.y).toBeCloseTo(50, 6);
		expect(paths.bottom[0]![0]!.x).toBeCloseTo(50, 6);
		// Clockwise: a quarter of the way round is the top of the box.
		expect(paths.top[0]![64]!.y).toBeCloseTo(0, 6);
		const start = closedWarpPaths('textCirclePour', W, H, 16200000, 25000)!;
		expect(start.top[0]![0]!.y).toBeCloseTo(0, 6);
		expect(start.bottom[0]![0]!.y).toBeCloseTo(37.5, 6);
	});

	it('leaves every other preset to the sampled paths', () => {
		expect(closedWarpPaths('textInflate', W, H, undefined, undefined)).toBeUndefined();
	});
});

describe('envelope warp over a closed loop', () => {
	it('maps the text ends to the loop start and the middle to the opposite tip', () => {
		const warp = createEnvelopeWarp('textCirclePour', W, H, undefined, undefined, BLOCK)!;
		const start = warp.map(0, 0);
		expect(start.x).toBeCloseTo(0, 6);
		expect(start.y).toBeCloseTo(50, 6);
		const half = warp.map(50, 0);
		expect(half.x).toBeCloseTo(W, 6);
		expect(half.y).toBeCloseTo(50, 6);
		// v = 1 is the inner ellipse.
		expect(warp.map(50, 40).x).toBeCloseTo(150, 6);
	});

	it('runs the first half of Ring Inside along the top of the upper ellipse', () => {
		const warp = createEnvelopeWarp('textRingInside', W, H, undefined, undefined, BLOCK)!;
		const top = warp.map(25, 0);
		expect(top.x).toBeCloseTo(100, 0);
		expect(top.y).toBeCloseTo(0, 0);
		// The second half is on the lower band, running right to left.
		const lower = warp.map(75, 0);
		expect(lower.x).toBeCloseTo(100, 0);
		expect(lower.y).toBeCloseTo(100, 0);
		expect(warp.clip?.splitsX).toEqual([50]);
	});

	it('clips a glyph to the ellipse of the band it lies in', () => {
		const warp = createEnvelopeWarp('textRingInside', W, H, undefined, undefined, BLOCK)!;
		// A tall box at the very start of the text: unclipped, its bottom edge sweeps down to
		// the lower ellipse's left tip, far below the upper ellipse.
		const box = [
			{ type: 'M', x: 0, y: 0 },
			{ type: 'L', x: 4, y: 0 },
			{ type: 'L', x: 4, y: 40 },
			{ type: 'L', x: 0, y: 40 },
			{ type: 'Z' },
		] as const;
		const d = buildWarpedGlyphOutlinePathD(box, warp, 1)!;
		const ys = [...d.matchAll(/[ML]([\d.-]+) ([\d.-]+)/gu)].map((m) => Number(m[2]));
		// The upper ellipse (centre 30, radius 30) ends at y = 60.
		expect(Math.max(...ys)).toBeLessThanOrEqual(60.01);
		expect(Math.max(...ys)).toBeGreaterThan(30);
	});
});

describe('polygon clipping', () => {
	const square = [
		{ x: -10, y: -10 },
		{ x: 10, y: -10 },
		{ x: 10, y: 10 },
		{ x: -10, y: 10 },
	];

	it('clips a polygon to an ellipse', () => {
		const clipped = clipPolygonConvex(square, ellipsePolygon(0, 0, 5, 5, 64));
		const radii = clipped.map((p) => Math.hypot(p.x, p.y));
		expect(Math.max(...radii)).toBeLessThanOrEqual(5.001);
		expect(clipped.length).toBeGreaterThan(8);
	});

	it('keeps a polygon that lies inside the clip', () => {
		const inside = [
			{ x: -1, y: -1 },
			{ x: 1, y: -1 },
			{ x: 0, y: 1 },
		];
		expect(clipPolygonConvex(inside, ellipsePolygon(0, 0, 5, 5, 32))).toHaveLength(3);
	});

	it('cuts a polygon at a vertical line', () => {
		const left = clipPolygonByX(square, 0, -1);
		expect(Math.max(...left.map((p) => p.x))).toBeCloseTo(0, 9);
		const right = clipPolygonByX(square, 0, 1);
		expect(Math.min(...right.map((p) => p.x))).toBeCloseTo(0, 9);
	});
});
