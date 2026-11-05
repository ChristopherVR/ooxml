import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphLineCarry } from './morph-line-carry';
import { buildMorphAnimationRules, buildMorphTransitionPlan } from './morph-plan';

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, slideNumber: 1, elements } as unknown as PptxSlide;
}

function line(
	id: string,
	box: { x: number; y: number; width: number; height: number },
	extra: Record<string, unknown> = {},
): PptxElement {
	return {
		id,
		name: '!!Card',
		type: 'connector',
		shapeType: 'line',
		...box,
		shapeStyle: { strokeColor: '#FFFFFF', strokeWidth: 4 },
		...extra,
	} as unknown as PptxElement;
}

const small = { x: 100, y: 100, width: 300, height: 60 };
const big = { x: 300, y: 250, width: 660, height: 200 };

/** The `stroke-width` and `matrix` of the keyframe at `percent`. */
function stopAt(keyframes: string, percent: number) {
	const block =
		new RegExp(`(?:^|\\n)\\t${percent}%[ ]\\{([^}]*)\\}`, 'u').exec(keyframes)?.[1] ?? '';
	const width = /stroke-width: ([\d.]+)px/u.exec(block)?.[1];
	const matrix = /matrix\(([^)]*)\)/u.exec(block)?.[1]?.split(',').map(Number);
	return { width: Number(width), matrix };
}

describe('morph line carry', () => {
	it('keeps a growing line at its width by shrinking the strand against the journey', () => {
		const from = line('a-1', small);
		const to = line('b-1', big);
		const carry = buildMorphLineCarry(from, to, 'to');
		expect(carry).toBeDefined();
		const start = stopAt(carry!.keyframes, 0);
		const end = stopAt(carry!.keyframes, 100);
		// The journey starts the wrapper at the old box, so the strand must draw
		// at the old line's length: its stroke is the width over that scale.
		const ratio = Math.hypot(300, 60) / Math.hypot(660, 200);
		expect(start.width).toBeCloseTo(4 / ratio, 1);
		expect(end.width).toBeCloseTo(4, 3);
		expect(end.matrix?.map((v) => Math.round(v * 1000) / 1000)).toEqual([1, 0, 0, 1, 0, 0]);
	});

	it('carries a flipped line to the new angle without mirroring twice', () => {
		const from = line('a-1', small, { flipVertical: true });
		const to = line('b-1', big, { flipVertical: true });
		const carry = buildMorphLineCarry(from, to, 'to');
		const mid = stopAt(carry!.keyframes, 0);
		expect(Number.isFinite(mid.width)).toBe(true);
		expect(mid.matrix).toHaveLength(6);
	});

	it('leaves a compound or bent connector to the journey', () => {
		expect(
			buildMorphLineCarry(
				line('a-1', small, { shapeStyle: { strokeWidth: 12, compoundLine: 'dbl' } }),
				line('b-1', big, { shapeStyle: { strokeWidth: 12, compoundLine: 'dbl' } }),
				'to',
			),
		).toBeUndefined();
		expect(
			buildMorphLineCarry(
				line('a-1', small, { shapeType: 'bentConnector3' }),
				line('b-1', big, { shapeType: 'bentConnector3' }),
				'to',
			),
		).toBeUndefined();
	});

	it('routes the strand animation to every binding through the plan rules', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [line('a-1', small)]),
			slide('b', [line('b-1', big)]),
			2000,
		);
		expect(plan?.lineAnimations?.get('b-1')).toMatch(/^pptx-morph-line-0-b1 2000ms/u);
		expect(buildMorphAnimationRules(plan!, '', 'incoming')).toMatch(
			/\[data-element-id="b-1"\] \[data-pptx-line\] \{ animation: pptx-morph-line-0-b1 /u,
		);
	});

	it('does not mirror the wrapper of a flipped connector, whose geometry is flipped already', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [line('a-1', small, { flipVertical: true })]),
			slide('b', [line('b-1', big)]),
			2000,
		);
		expect(plan?.keyframesCss).not.toContain('scaleY(-1)');
	});
});
