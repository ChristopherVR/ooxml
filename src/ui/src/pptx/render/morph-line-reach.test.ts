import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphConnectorCarry } from './morph-connector-carry';
import type { Matrix } from './morph-line-carry';
import { apply, buildMorphLineCarry, journeyOf, mul, wrapperAt } from './morph-line-carry';
import { reachCoupling, reachedExtent, reachFactors } from './morph-line-reach';

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
		expect(reachedExtent(100, 300, 100, 0.4)).toBeCloseTo(180 + 0.75 * 200 * 0.24, 6);
		// A shrink is larger than its blend as well.
		expect(reachedExtent(300, 100, 100, 0.4)).toBeCloseTo(220 + 0.75 * 200 * 0.24, 6);
	});

	it('fades with the centre travel, whichever way it goes', () => {
		expect(reachCoupling(1)).toBe(1);
		expect(reachCoupling(0.5)).toBeLessThan(0.6);
		expect(reachCoupling(2)).toBeCloseTo(0.57, 2);
		expect(reachedExtent(100, 300, 400, 0.5)).toBeLessThan(reachedExtent(100, 300, 100, 0.5));
	});

	it('leaves flipped, rotated and unresized axes linear', () => {
		const a = el('line', { x: 0, y: 0, width: 100, height: 50 });
		const flipped = el('line', { x: 0, y: 0, width: 300, height: 50 }, { flipHorizontal: true });
		expect(reachFactors(a, flipped, 0.5)[0]).toBe(1);
		const turned = el('line', { x: 0, y: 0, width: 300, height: 50 }, { rotation: 30 });
		expect(reachFactors(a, turned, 0.5)).toEqual([1, 1]);
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
});
