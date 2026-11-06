import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphConnectorCarry } from './morph-connector-carry';
import type { Matrix } from './morph-line-carry';
import { apply, buildMorphLineCarry, journeyOf, mul, wrapperAt } from './morph-line-carry';
import { reachCoupling, reachedExtent, reachFactors, reachOffsets } from './morph-line-reach';

const el = (
	type: string,
	box: { x: number; y: number; width: number; height: number },
	extra: Record<string, unknown> = {},
): PptxElement =>
	({
		id: 'c',
		name: '!!Card',
		type: 'connector',
		shapeType: type,
		...box,
		shapeStyle: { strokeColor: '#FFFFFF', strokeWidth: 3 },
		...extra,
	}) as unknown as PptxElement;

describe('Morph line reach (measured on PowerPoint 16)', () => {
	it('is exactly linear at the ends and for a box that grows about a fixed centre', () => {
		expect(reachedExtent(100, 300, 200, 0)).toBe(100);
		expect(reachedExtent(100, 300, 200, 1)).toBe(300);
		expect(reachedExtent(100, 300, 0, 0.5)).toBeCloseTo(200, 6);
	});

	it('grows 3x about a fixed left edge ahead of the linear blend (+0.75 * delta * p(1-p))', () => {
		// 100 -> 300 with the left edge fixed: the centre travels half the delta.
		expect(reachedExtent(100, 300, 100, 0.4)).toBeCloseTo(180 + 0.75 * 200 * 0.24, 0);
		// A shrink is larger than its blend as well.
		expect(reachedExtent(300, 100, 100, 0.4)).toBeCloseTo(220 + 0.75 * 200 * 0.24, 0);
	});

	it('fades with the centre travel, whichever way it goes', () => {
		expect(reachCoupling(1)).toBe(1);
		expect(reachCoupling(0.5)).toBeLessThan(0.6);
		expect(reachCoupling(2)).toBeCloseTo(0.57, 2);
		expect(reachedExtent(100, 300, 400, 0.5)).toBeLessThan(reachedExtent(100, 300, 100, 0.5));
	});

	it('leaves flipped and unresized axes linear for the factor form', () => {
		const a = el('line', { x: 0, y: 0, width: 100, height: 50 });
		const flipped = el('line', { x: 0, y: 0, width: 300, height: 50 }, { flipHorizontal: true });
		expect(reachFactors(a, flipped, 0.5)[0]).toBe(1);
		const wide = el('line', { x: 0, y: 0, width: 300, height: 50 });
		expect(reachFactors(a, wide, 0.5)[1]).toBe(1);
		expect(reachFactors(a, wide, 0.5)[0]).toBeGreaterThan(1);
	});

	it('draws a straight line larger than the linear blend mid-flight', () => {
		const from = el('line', { x: 100, y: 100, width: 100, height: 100 });
		const to = el('line', { x: 100, y: 100, width: 300, height: 300 });
		const keyframes = buildMorphLineCarry(from, to, 'from')?.keyframes ?? '';
		const mid = [...keyframes.matchAll(/\d+(?:\.\d+)?% \{([^}]*)\}/gu)][12]?.[1] ?? '';
		const m = (/matrix\(([^)]*)\)/u.exec(mid)?.[1]?.split(',').map(Number) ??
			[]) as unknown as Matrix;
		// The strand's own matrix undoes the wrapper; composed with it, it lands on the slide.
		const net = mul(wrapperAt(journeyOf(from, to, 'from'), 0.5), m);
		const [x0, y0] = apply(net, [0, 0]);
		const [x1, y1] = apply(net, [100, 100]);
		// 100 -> 300 about a left edge that stays put: 200 + 0.75 * 200 * 0.25 wide.
		expect(Math.abs(x1 - x0)).toBeGreaterThan(236);
		expect(Math.abs(x1 - x0)).toBeLessThan(239);
		expect(Math.abs(y1 - y0)).toBeGreaterThan(236);
	});

	it('draws an elbow connector larger than the linear blend mid-flight', () => {
		const from = el('bentConnector3', { x: 100, y: 100, width: 200, height: 150 });
		const to = el('bentConnector3', { x: 250, y: 150, width: 600, height: 300 });
		const keyframes = buildMorphConnectorCarry(from, to, 'to') ?? '';
		const mid = [...keyframes.matchAll(/\d+(?:\.\d+)?% \{([^}]*)\}/gu)][12]?.[1] ?? '';
		const d = /d: path\('([^']*)'\)/u.exec(mid)?.[1] ?? '';
		const xs = (d.match(/-?\d+(?:\.\d+)?/gu) ?? []).map(Number).filter((_, i) => i % 2 === 0);
		const span = Math.max(...xs) - Math.min(...xs);
		// Linear blend of 200 and 600 is 400; PowerPoint reach makes it about 449.
		expect(span).toBeGreaterThan(440);
		expect(span).toBeLessThan(456);
	});

	it('fades the overshoot out for small ratios (1.1x none, 1.4x about 70%)', () => {
		// 200 -> 220 about a fixed end: PowerPoint was measured at 0.5 pt, not 2.6.
		expect(reachedExtent(200, 220, 10, 0.5) - 210).toBeLessThan(1);
		// 200 -> 280: +80 * 0.583 * 0.72 * 0.25 = 8.4 where the unfaded law gave 11.7.
		expect(reachedExtent(200, 280, 40, 0.5) - 240).toBeCloseTo(8.4, 0);
		// A shrink fades the same way.
		expect(reachedExtent(280, 200, 40, 0.5) - 240).toBeCloseTo(8.4, 0);
	});

	it('overshoots through zero when the two ends flip differently', () => {
		// 150 high, then 300 high flipped: signed -300; PowerPoint reaches 198 * p(1-p) further.
		const a = el('line', { x: 100, y: 100, width: 200, height: 150 });
		const b = el('line', { x: 100, y: 100, width: 400, height: 300 }, { flipVertical: true });
		const [dx, dy] = reachOffsets(a, b, 0.5);
		expect(dy).toBeLessThan(-45);
		expect(dy).toBeGreaterThan(-52);
		expect(dx).toBeGreaterThan(0);
		// Equal magnitudes turned over about one place: just the linear blend through zero.
		const same = el('line', { x: 100, y: 100, width: 200, height: 100 });
		const over = el('line', { x: 100, y: 100, width: 200, height: 100 }, { flipVertical: true });
		expect(reachOffsets(same, over, 0.5)).toEqual([0, 0]);
		// An elbow's flipped axis was not measured: it stays linear.
		expect(reachOffsets(a, b, 0.5, false)[1]).toBe(0);
	});

	it('keeps a line turning about a fixed centre on the chord, and reaches when the centre moves', () => {
		const turnOf = (to: Record<string, unknown>, box: { x: number; y: number }) => {
			const a = el('line', { x: 330, y: 270, width: 300, height: 0 }, { rotation: 20 });
			const b = el('line', { ...box, width: 300, height: 0 }, { rotation: 70, ...to });
			const [dx, dy] = reachOffsets(a, b, 0.5);
			// The blended vector at the middle of the turn plus the overshoot.
			const [ax, ay] = [300 * Math.cos(Math.PI / 9), 300 * Math.sin(Math.PI / 9)];
			const [bx, by] = [300 * Math.cos((70 * Math.PI) / 180), 300 * Math.sin((70 * Math.PI) / 180)];
			return Math.hypot((ax + bx) / 2 + dx, (ay + by) / 2 + dy);
		};
		// In place the chord shortens mid-turn (272 of 300), as PowerPoint draws it.
		expect(turnOf({}, { x: 330, y: 270 })).toBeCloseTo(272.2, 0);
		// Moving 112 pt as it turns, the chord is made up again (measured 306 of 300).
		expect(turnOf({}, { x: 430, y: 320 })).toBeGreaterThan(300);
		expect(turnOf({}, { x: 430, y: 320 })).toBeLessThan(312);
	});

	it('draws a flipped, resized line past the linear blend mid-flight', () => {
		const from = el('line', { x: 100, y: 100, width: 200, height: 150 });
		const to = el('line', { x: 100, y: 100, width: 400, height: 300 }, { flipVertical: true });
		const keyframes = buildMorphLineCarry(from, to, 'from')?.keyframes ?? '';
		const mid = [...keyframes.matchAll(/\d+(?:\.\d+)?% \{([^}]*)\}/gu)][12]?.[1] ?? '';
		const m = (/matrix\(([^)]*)\)/u.exec(mid)?.[1]?.split(',').map(Number) ??
			[]) as unknown as Matrix;
		const net = mul(wrapperAt(journeyOf(from, to, 'from'), 0.5), m);
		const [x0, y0] = apply(net, [0, 0]);
		const [x1, y1] = apply(net, [200, 150]);
		// PowerPoint: 331 wide and 124.5 tall here (linear blend 301 by 77).
		expect(Math.abs(x1 - x0)).toBeGreaterThan(322);
		expect(Math.abs(y1 - y0)).toBeGreaterThan(115);
		expect(Math.abs(y1 - y0)).toBeLessThan(133);
	});

	it('draws a rotated, moving line longer than the blended chord mid-flight', () => {
		const from = el('line', { x: 330, y: 270, width: 300, height: 0 }, { rotation: 20 });
		const to = el('line', { x: 430, y: 320, width: 300, height: 0 }, { rotation: 70 });
		const keyframes = buildMorphLineCarry(from, to, 'from')?.keyframes ?? '';
		const mid = [...keyframes.matchAll(/\d+(?:\.\d+)?% \{([^}]*)\}/gu)][12]?.[1] ?? '';
		const m = (/matrix\(([^)]*)\)/u.exec(mid)?.[1]?.split(',').map(Number) ??
			[]) as unknown as Matrix;
		const net = mul(wrapperAt(journeyOf(from, to, 'from'), 0.5), m);
		const [x0, y0] = apply(net, [0, 0]);
		const [x1, y1] = apply(net, [300, 0]);
		// PowerPoint: 306 long here; the blend of the end points alone is 272.
		expect(Math.hypot(x1 - x0, y1 - y0)).toBeGreaterThan(298);
		expect(Math.hypot(x1 - x0, y1 - y0)).toBeLessThan(314);
	});
});
