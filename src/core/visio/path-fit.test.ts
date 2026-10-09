import { describe, expect, it } from 'vitest';
import {
	visioArcSamples,
	visioCircleThrough,
	visioFitFreeform,
	visioFitPencil,
	visioPathBounds,
	visioQuarterArc,
	visioSimplifyPath,
} from './path-fit';

const range = (count: number) => Array.from({ length: count }, (_, index) => index);

describe('path fitting', () => {
	it('simplifies collinear and noisy samples with Ramer-Douglas-Peucker', () => {
		const line = range(50).map((index) => ({ x: index / 10, y: index % 2 ? 0.001 : 0 }));
		expect(visioSimplifyPath(line, 0.01)).toEqual([line[0], line[49]]);
		const corner = [
			...range(11).map((index) => ({ x: index / 10, y: 0 })),
			...range(10).map((index) => ({ x: 1, y: (index + 1) / 10 })),
		];
		expect(visioSimplifyPath(corner, 0.01)).toEqual([
			{ x: 0, y: 0 },
			{ x: 1, y: 0 },
			{ x: 1, y: 1 },
		]);
	});

	it('fits freeform samples with smooth cubic segments through the simplified points', () => {
		const wave = range(61).map((index) => ({ x: index / 20, y: Math.sin(index / 10) }));
		const segments = visioFitFreeform(wave, 0.02, false);
		expect(segments.length).toBeGreaterThan(2);
		expect(segments.every((segment) => segment.kind === 'cubic')).toBe(true);
		expect(segments.at(-1)).toMatchObject({ x: 3, y: Math.sin(6) });
		const ring = range(41).map((index) => ({
			x: Math.cos((index / 40) * 2 * Math.PI),
			y: Math.sin((index / 40) * 2 * Math.PI),
		}));
		ring[40] = { ...ring[0]! };
		const closed = visioFitFreeform(ring, 0.01, true);
		expect(closed.at(-1)).toMatchObject({ x: 1, y: 0 });
		const bounds = visioPathBounds(ring[0]!, closed);
		expect(bounds.minX).toBeCloseTo(-1, 1);
		expect(bounds.maxY).toBeCloseTo(1, 1);
	});

	it('splits pencil strokes into straight lines at corners and arcs where they bow', () => {
		const corner = [
			...range(11).map((index) => ({ x: index / 10, y: 0 })),
			...range(10).map((index) => ({ x: 1, y: (index + 1) / 10 })),
		];
		expect(visioFitPencil(corner, 1 / 32)).toEqual([
			{ kind: 'line', x: 1, y: 0 },
			{ kind: 'line', x: 1, y: 1 },
		]);
		const bow = range(21).map((index) => {
			const angle = Math.PI - (index / 20) * Math.PI;
			return { x: Math.cos(angle), y: Math.sin(angle) };
		});
		const [arc, ...rest] = visioFitPencil(bow, 1 / 32);
		expect(rest).toEqual([]);
		expect(arc).toMatchObject({ kind: 'arc', ratio: 1 });
		expect(arc!.x).toBeCloseTo(1, 9);
		expect(
			Math.hypot(arc!.kind === 'arc' ? arc!.a : 0, arc!.kind === 'arc' ? arc!.b : 0),
		).toBeCloseTo(1, 9);
	});

	it('samples arcs through their control point and rejects collinear circles', () => {
		expect(visioCircleThrough({ x: 0, y: 0 }, { x: 1, y: 1 }, { x: 2, y: 2 })).toBeUndefined();
		const arc = visioQuarterArc({ x: 0, y: 2 }, { x: 4, y: 0 })!;
		expect(arc).toMatchObject({ kind: 'arc', x: 4, y: 0, ratio: 2 });
		const samples = visioArcSamples({ x: 0, y: 2 }, arc, 16);
		// Every sample lies on the ellipse centred at (0, 0) with radii 4 and 2.
		for (const point of samples) expect((point.x / 4) ** 2 + (point.y / 2) ** 2).toBeCloseTo(1, 9);
		expect(samples.every((point) => point.x >= -1e-9 && point.y >= -1e-9)).toBe(true);
		expect(visioQuarterArc({ x: 0, y: 0 }, { x: 0, y: 1 })).toBeUndefined();
	});
});
