import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	buildMorphAnimationRules,
	buildMorphScopedCss,
	buildMorphTransitionPlan,
} from './morph-plan';

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, slideNumber: 1, elements } as unknown as PptxSlide;
}

function card(
	id: string,
	box: { x: number; y: number; width: number; height: number },
	style: Record<string, unknown> = {},
	extra: Record<string, unknown> = {},
): PptxElement {
	return {
		id,
		name: 'Card',
		type: 'shape',
		shapeType: 'roundRect',
		...box,
		shapeStyle: { fillColor: '#C07020', strokeColor: '#FFFFFF', strokeWidth: 8, ...style },
		...extra,
	} as unknown as PptxElement;
}

describe('morph geometry and effect carry in the plan', () => {
	it('re-paths a growing outline instead of scaling it with the box', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [card('a-1', { x: 100, y: 100, width: 400, height: 160 })]),
			slide('b', [card('b-1', { x: 300, y: 200, width: 560, height: 240 })]),
			2000,
		);
		expect(plan?.outlineAnimations?.get('b-1')).toMatch(/^pptx-morph-carry-0-b1-outline 2000ms/u);
		// The radii ride the journey's own curve, composed on the same element.
		expect(plan?.incomingAnimations.get('b-1')).toContain('pptx-morph-carry-0-b1 2000ms');
		expect(plan?.keyframesCss).toContain('border-radius: 37.334px / 40.001px');
		expect(plan?.keyframesCss).not.toContain('pptx-morph-geo-');
		// Every binding gets the outline rule, behind a `d` support check.
		for (const css of [
			buildMorphAnimationRules(plan!, '', 'incoming', 'image'),
			buildMorphScopedCss(plan!, 'data-pptx-morph-incoming', 'incoming'),
		]) {
			expect(css).toContain(`@supports (d: path('M0 0'))`);
			expect(css).toMatch(
				/\[data-element-id="b-1"\] \[data-pptx-outline\] path \{ animation: pptx-morph-carry-0-b1-outline /u,
			);
		}
	});

	it('gives a dissolving pair the same geometry on both halves', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [card('a-1', { x: 100, y: 100, width: 400, height: 200 })]),
			slide('b', [
				card(
					'b-1',
					{ x: 100, y: 100, width: 400, height: 200 },
					{},
					{ shapeAdjustments: { adj: 50000 } },
				),
			]),
			2000,
		);
		expect(plan?.outlineAnimations?.has('b-1')).toBe(true);
		expect(plan?.outlineAnimations?.has('a-1')).toBe(true);
		expect(plan?.outgoingAnimations.get('a-1')).toContain('pptx-morph-carry-g0-a1');
	});

	it('fades a shadow that changes where it stands instead of showing the new one at once', () => {
		const shadow = {
			shadowColor: '#000000',
			shadowOpacity: 1,
			shadowBlur: 0,
			shadowAngle: 45,
			shadowDistance: 20,
			strokeWidth: 0,
		};
		const plan = buildMorphTransitionPlan(
			slide('a', [card('a-1', { x: 100, y: 100, width: 400, height: 200 }, shadow)]),
			slide('b', [
				card('b-1', { x: 100, y: 100, width: 400, height: 200 }, { ...shadow, shadowDistance: 60 }),
			]),
			2000,
		);
		expect(plan?.incomingAnimations.get('b-1')).toContain('pptx-morph-carry-0-b1-fx 2000ms linear');
		// The overlap layer first, then the incoming shadow fading in over the outgoing one.
		expect(plan?.keyframesCss).toMatch(
			/box-shadow: 28px 28px 0px -14px rgba\(0, 0, 0, 0\), 42px 42px 0px rgba\(0, 0, 0, 0\), 14px 14px/u,
		);
	});

	it('fades the incoming shadow in under a ghost that dissolves over an opaque body', () => {
		// The fill changes, so the outgoing ghost dissolves over the incoming half,
		// which stays at full strength: PowerPoint still fades the incoming shadow
		// in on p (measured: the overlap holds 1.00 and each alone runs 1 - p, p).
		const shadow = {
			shadowColor: '#000000',
			shadowOpacity: 1,
			shadowBlur: 0,
			shadowAngle: 45,
			shadowDistance: 20,
			strokeWidth: 0,
		};
		const plan = buildMorphTransitionPlan(
			slide('a', [
				card('a-1', { x: 100, y: 100, width: 200, height: 120 }, shadow, { name: '!!Card' }),
			]),
			slide('b', [
				card(
					'b-1',
					{ x: 300, y: 200, width: 440, height: 240 },
					{ ...shadow, fillColor: '#2070C0' },
					{ name: '!!Card' },
				),
			]),
			2000,
		);
		expect(plan?.incomingAnimations.get('b-1')).toContain('pptx-morph-carry-0-b1-fx 2000ms linear');
		const body = plan?.keyframesCss.split('@keyframes pptx-morph-carry-0-b1-fx {')[1] ?? '';
		// No outgoing layer on this half: the ghost paints that one itself.
		// (the outgoing 14 px shadow rides at 30.8 x 28 px in this frame).
		expect(body).not.toContain('30.8px 28px 0px rgba');
		expect(body).toContain('px rgba(0, 0, 0, 1)');
		expect(body).toMatch(/0% \{\n\t\tbox-shadow: [^;]*14px 14px 0px rgba\(0, 0, 0, 0\);/u);
		// The ghost keeps painting its own shadow rather than a carry of it.
		expect(plan?.outgoingAnimations.get('a-1') ?? '').not.toContain('-fx');
	});

	it('leaves a shape that only moves exactly as it was', () => {
		const plan = buildMorphTransitionPlan(
			slide('a', [card('a-1', { x: 100, y: 100, width: 400, height: 160 })]),
			slide('b', [card('b-1', { x: 300, y: 200, width: 400, height: 160 })]),
			2000,
		);
		expect(plan?.outlineAnimations).toBeUndefined();
		expect(plan?.keyframesCss).not.toContain('pptx-morph-carry');
	});
});
