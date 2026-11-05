import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { getConnectorPathGeometry } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphConnectorCarry } from './morph-connector-carry';
import { buildMorphAnimationRules, buildMorphTransitionPlan } from './morph-plan';

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, slideNumber: 1, elements } as unknown as PptxSlide;
}

function connector(
	id: string,
	box: { x: number; y: number; width: number; height: number },
	extra: Record<string, unknown> = {},
): PptxElement {
	return {
		id,
		name: '!!Card',
		type: 'connector',
		shapeType: 'bentConnector3',
		...box,
		shapeStyle: { strokeColor: '#FFFFFF', strokeWidth: 8 },
		...extra,
	} as unknown as PptxElement;
}

const small = { x: 133, y: 133, width: 267, height: 200 };
const big = { x: 333, y: 200, width: 800, height: 400 };

interface Stop {
	readonly width: number;
	readonly points: number[][];
}

/** Every baked stop in order (stop `i` is progress `i / 24`). */
function stops(keyframes: string): Stop[] {
	return [...keyframes.matchAll(/\d+(?:\.\d+)?% \{([^}]*)\}/gu)].map((block) => {
		const body = block[1] ?? '';
		const d = /d: path\('([^']*)'\)/u.exec(body)?.[1] ?? '';
		const numbers = (d.match(/-?\d+(?:\.\d+)?/gu) ?? []).map(Number);
		const points: number[][] = [];
		for (let i = 0; i + 1 < numbers.length; i += 2) points.push([numbers[i]!, numbers[i + 1]!]);
		return { width: Number(/stroke-width: ([\d.]+)px/u.exec(body)?.[1]), points };
	});
}

const carry = (from: PptxElement, to: PptxElement) =>
	stops(buildMorphConnectorCarry(from, to, 'to') ?? '');

/** The authored route of `element` in slide coordinates. */
function route(element: PptxElement): number[][] {
	const d = getConnectorPathGeometry(element as never).pathData;
	const numbers = (d.match(/-?\d+(?:\.\d+)?/gu) ?? []).map(Number);
	const points: number[][] = [];
	for (let i = 0; i + 1 < numbers.length; i += 2) {
		points.push([numbers[i]! + element.x, numbers[i + 1]! + element.y]);
	}
	return points;
}

describe('morph elbow and curved connector carry', () => {
	it.each(['bentConnector3', 'curvedConnector3'])(
		'keeps the %s stroke width and ends on both routes while it grows',
		(shapeType) => {
			const from = connector('a-1', small, { shapeType });
			const to = connector('b-1', big, { shapeType });
			const all = carry(from, to);
			expect(all).toHaveLength(25);
			expect(all.every((stop) => stop.width === 8)).toBe(true);
			const [first, last] = [all[0]!, all.at(-1)!];
			route(from).forEach((point, i) => {
				expect(first.points[i]![0]).toBeCloseTo(point[0]!, 0);
				expect(first.points[i]![1]).toBeCloseTo(point[1]!, 0);
			});
			route(to).forEach((point, i) => {
				expect(last.points[i]![0]).toBeCloseTo(point[0]!, 0);
				expect(last.points[i]![1]).toBeCloseTo(point[1]!, 0);
			});
		},
	);

	it('keeps the bend at its fraction of the interpolated box', () => {
		const mid = carry(connector('a-1', small), connector('b-1', big))[12]!;
		// bentConnector3: M0,0 L adj,0 L adj,h L w,h; at progress 0.5 the box is 200..., 0.5 wide.
		const left = mid.points[0]![0]!;
		const right = mid.points.at(-1)![0]!;
		expect(mid.points[1]![0]! - left).toBeCloseTo((right - left) / 2, 0);
	});

	it('collapses a flipped elbow through zero width and opens it out mirrored', () => {
		const from = connector('a-1', small);
		const to = connector('b-1', small, { flipHorizontal: true });
		const all = carry(from, to);
		const spread = (stop: Stop) => {
			const xs = stop.points.map((point) => point[0]!);
			return Math.max(...xs) - Math.min(...xs);
		};
		expect(spread(all[12]!)).toBeLessThan(1);
		expect(spread(all[6]!)).toBeCloseTo(spread(all[18]!), 0);
		// Start at the left edge, end at the right edge; flipped, the other way round.
		expect(all[0]!.points[0]![0]).toBeLessThan(all[0]!.points.at(-1)![0]!);
		expect(all[24]!.points[0]![0]).toBeGreaterThan(all[24]!.points.at(-1)![0]!);
	});

	it('moves the bend linearly between two adj values', () => {
		const from = connector('a-1', small, { shapeAdjustments: { adj1: 50000 } });
		const to = connector('b-1', small, { shapeAdjustments: { adj1: 15000 } });
		const all = carry(from, to);
		const bend = (stop: Stop) => (stop.points[1]![0]! - small.x) / small.width;
		expect(bend(all[0]!)).toBeCloseTo(0.5, 2);
		expect(bend(all[12]!)).toBeCloseTo(0.325, 2);
		expect(bend(all[24]!)).toBeCloseTo(0.15, 2);
	});

	it('leaves an unchanged pair, a preset change and a straight line alone', () => {
		const a = connector('a-1', small);
		expect(buildMorphConnectorCarry(a, connector('b-1', small), 'to')).toBeUndefined();
		expect(
			buildMorphConnectorCarry(a, connector('b-1', big, { shapeType: 'curvedConnector3' }), 'to'),
		).toBeUndefined();
		expect(
			buildMorphConnectorCarry(
				connector('a-1', small, { shapeType: 'line' }),
				connector('b-1', big, { shapeType: 'line' }),
				'to',
			),
		).toBeUndefined();
	});

	it('routes the strand animation to every binding through the plan rules', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [connector('a-1', small)]),
			slide('b', [connector('b-1', big)]),
			2000,
		);
		expect(plan?.lineAnimations?.get('b-1')).toMatch(/^pptx-morph-connector-0-b1 2000ms/u);
		expect(buildMorphAnimationRules(plan!, '', 'incoming')).toMatch(
			/\[data-element-id="b-1"\] \[data-pptx-line\] \{ animation: pptx-morph-connector-0-b1 /u,
		);
	});
	it('moves the bend of one opaque line instead of dissolving two and clipping the outline', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [connector('a-1', small, { shapeAdjustments: { adj1: 50000 } })]),
			slide('b', [connector('b-1', small, { shapeAdjustments: { adj1: 15000 } })]),
			2000,
		);
		expect(plan?.lineAnimations?.get('b-1')).toMatch(/^pptx-morph-connector-0-b1 /u);
		expect(plan?.keyframesCss).not.toContain('pptx-morph-geo-');
		// No ghost dissolves over the incoming half, so nothing in the plan fades the pair.
		expect(plan?.keyframesCss).not.toMatch(/opacity: 0\.\d/u);
	});
});
