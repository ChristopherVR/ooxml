import { describe, expect, it } from 'vitest';

import type { ShadowLayer } from './morph-effect-shadow';
import { overlapColor, serializeOverlap, shadowOverlap } from './morph-effect-shadow';

const black = (dx: number, dy: number, alpha = 1, blur = 0): ShadowLayer => ({
	inset: false,
	lengths: [dx, dy, blur],
	color: [0, 0, 0, alpha],
});

/** Source-over of [r, g, b, a] layers listed top first, on a transparent backdrop. */
function stack(layers: Array<[number, number, number, number]>): [number, number] {
	let alpha = 0;
	let premul = 0;
	for (const [r, , , a] of [...layers].reverse()) {
		premul = r * a + premul * (1 - a);
		alpha = a + alpha * (1 - a);
	}
	return [alpha, alpha > 0 ? premul / alpha : 0];
}

describe('morph shadow overlap', () => {
	it('paints the overlap of two diagonal shadows as one shrunken box', () => {
		const overlap = shadowOverlap(black(14, 14), black(42, 42), 400, 200);
		expect(overlap?.tiles).toEqual([{ dx: 28, dy: 28, spread: -14 }]);
		expect(serializeOverlap(overlap!, [0, 0, 0, 1])).toEqual([
			'28px 28px 0px -14px rgba(0, 0, 0, 1)',
		]);
	});

	it('tiles an overlap that is not a shrunken box', () => {
		// 400 x 200 box, shadows 60 px apart horizontally and 10 px vertically:
		// the overlap is 340 x 190, tiled by two 340 x 140 boxes.
		const overlap = shadowOverlap(black(0, 0), black(60, 10), 400, 200);
		const tiles = overlap?.tiles ?? [];
		expect(tiles.length).toBe(2);
		for (const t of tiles) {
			expect(t.spread).toBeCloseTo(-30, 6);
			// Every tile stays inside the 340 x 190 overlap at (60, 10).
			expect(t.dx - t.spread).toBeGreaterThanOrEqual(60 - 1e-6);
			expect(400 + t.dx + t.spread).toBeLessThanOrEqual(400 + 1e-6);
			expect(t.dy - t.spread).toBeGreaterThanOrEqual(10 - 1e-6);
			expect(200 + t.dy + t.spread).toBeLessThanOrEqual(200 + 1e-6);
		}
		const left = Math.min(...tiles.map((t) => t.dx - t.spread));
		const bottom = Math.max(...tiles.map((t) => 200 + t.dy + t.spread));
		expect(left).toBeCloseTo(60, 6);
		expect(bottom).toBeCloseTo(200, 6);
	});

	it('paints a translucent overlap once, centred, rather than with tiles that overlap', () => {
		// Two overlapping tiles would composite the overlap layer twice where
		// they cross (measured: 0.56 instead of 0.52 at 50 % transparency).
		const overlap = shadowOverlap(black(0, 0, 0.5), black(60, 10, 0.5), 400, 200);
		expect(overlap?.tiles).toEqual([{ dx: 30, dy: 5, spread: -30 }]);
	});

	it('has no overlap for inset shadows or shadows that miss each other', () => {
		expect(shadowOverlap({ ...black(5, 5), inset: true }, black(9, 9), 100, 100)).toBeUndefined();
		expect(shadowOverlap(black(0, 0), black(150, 0), 100, 100)).toBeUndefined();
	});

	it('keeps an opaque overlap opaque at every progress, like PowerPoint (alone)', () => {
		for (const p of [0.1, 0.25, 0.5, 0.75, 0.9]) {
			const top = overlapColor([0, 0, 0, 1], [0, 0, 0, 1], p, 'alone')!;
			const [alpha] = stack([top, [0, 0, 0, p], [0, 0, 0, 1 - p]]);
			expect(alpha).toBeCloseTo(1, 6);
		}
	});

	it('holds a half-transparent overlap at its own alpha, like PowerPoint (alone)', () => {
		// PowerPoint, 50 % transparent shadow: the overlap reads 0.52 at every frame.
		for (const p of [0.2, 0.5, 0.8]) {
			const top = overlapColor([0, 0, 0, 0.5], [0, 0, 0, 0.5], p, 'alone')!;
			const [alpha] = stack([top, [0, 0, 0, 0.5 * p], [0, 0, 0, 0.5 * (1 - p)]]);
			expect(alpha).toBeCloseTo(0.5, 6);
		}
	});

	it('dissolves two colours additively where the shadows overlap', () => {
		const p = 0.3;
		const top = overlapColor([200, 0, 0, 1], [0, 0, 0, 1], p, 'alone')!;
		const [alpha, red] = stack([top, [0, 0, 0, p], [200, 0, 0, 1 - p]]);
		expect(alpha).toBeCloseTo(1, 6);
		expect(red).toBeCloseTo(200 * (1 - p), 4);
	});

	it('solves for the ghost painting the outgoing shadow on top (underGhost)', () => {
		for (const [c, p] of [
			[1, 0.5],
			[0.5, 0.3],
		] as const) {
			const top = overlapColor([0, 0, 0, c], [0, 0, 0, c], p, 'underGhost')!;
			const [mine] = stack([top, [0, 0, 0, c * p]]);
			const [alpha] = stack([
				[0, 0, 0, c * (1 - p)],
				[0, 0, 0, mine],
			]);
			expect(alpha).toBeCloseTo(c, 6);
		}
	});
});
