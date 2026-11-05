import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { morphPairIncomingFadesIn, morphPairNeedsCrossfade } from './morph-animation';
import { buildMorphTransitionPlan } from './morph-plan';
import { morphSoftEdgeBlends } from './morph-soft-edge-blend';

function soft(id: string, width: number, height: number, radius: number): PptxElement {
	return {
		id,
		name: 'Soft',
		type: 'shape',
		shapeType: 'rect',
		shapeId: 5,
		x: 100,
		y: 100,
		width,
		height,
		shapeStyle: { fillMode: 'solid', fillColor: '#C07020', softEdgeRadius: radius },
	} as unknown as PptxElement;
}

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, slideNumber: 1, elements } as unknown as PptxSlide;
}

describe('morphSoftEdgeBlends', () => {
	it('is true for a soft edge on a box that changes size', () => {
		expect(morphSoftEdgeBlends(soft('a', 200, 120, 20), soft('b', 440, 240, 20))).toBe(true);
	});

	it('is true when the radius changes in place', () => {
		expect(morphSoftEdgeBlends(soft('a', 200, 120, 10), soft('b', 200, 120, 30))).toBe(true);
	});

	it('is false for an unchanged soft edge and for shapes without one', () => {
		expect(morphSoftEdgeBlends(soft('a', 200, 120, 20), soft('b', 200, 120, 20))).toBe(false);
		expect(morphSoftEdgeBlends(soft('a', 200, 120, 0), soft('b', 440, 240, 0))).toBe(false);
	});
});

describe('a soft-edged pair in the morph plan', () => {
	it('dissolves both halves, though the body is opaque', () => {
		const from = soft('a-1', 200, 120, 20);
		const to = soft('b-1', 440, 240, 20);

		expect(morphPairNeedsCrossfade(from, to)).toBe(true);
		expect(morphPairIncomingFadesIn(from, to, true)).toBe(true);
	});

	it('hands the two feathers to one additive crossfade group', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [soft('a-1', 200, 120, 20)]),
			slide('b', [soft('b-1', 440, 240, 20)]),
			500,
		);

		expect(
			plan?.crossfadeGroups.map((group) => [group.outgoing.id, group.incoming.id]),
		).toStrictEqual([['a-1', 'b-1']]);
	});

	it('leaves a soft edge that does not change alone', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [soft('a-1', 200, 120, 20)]),
			slide('b', [{ ...soft('b-1', 200, 120, 20), x: 300 } as PptxElement]),
			500,
		);

		expect(plan?.crossfadeGroups).toStrictEqual([]);
	});
});
