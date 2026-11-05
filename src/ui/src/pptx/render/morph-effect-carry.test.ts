import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphEffectCarry } from './morph-effect-carry';

function rect(
	id: string,
	width: number,
	height: number,
	style: Record<string, unknown>,
): PptxElement {
	return {
		id,
		type: 'shape',
		shapeType: 'rect',
		x: 0,
		y: 0,
		width,
		height,
		shapeStyle: { fillColor: '#C07020', ...style },
	} as unknown as PptxElement;
}

const shadow = {
	shadowColor: '#000000',
	shadowOpacity: 1,
	shadowBlur: 0,
	shadowAngle: 45,
	shadowDistance: Math.SQRT2 * 20,
};

describe('morph effect carry', () => {
	it('crossfades the two ends shadows, each riding the box', () => {
		// PowerPoint: one shadow offset 20 * W(p) / W0 fading out on 1 - p, one
		// offset 20 * W(p) / W1 fading in on p.
		const carry = buildMorphEffectCarry(
			rect('a', 200, 100, shadow),
			rect('b', 400, 300, shadow),
			'to',
		);
		expect(carry?.startDeclarations).toEqual([
			'box-shadow: 20px 20px 0px rgba(0, 0, 0, 0), 40px 60px 0px rgba(0, 0, 0, 1)',
		]);
		expect(carry?.endDeclarations).toEqual([
			'box-shadow: 20px 20px 0px rgba(0, 0, 0, 1), 40px 60px 0px rgba(0, 0, 0, 0)',
		]);
	});

	it('writes the ghost half in its own frame', () => {
		const carry = buildMorphEffectCarry(
			rect('a', 200, 100, shadow),
			rect('b', 400, 300, shadow),
			'from',
		);
		expect(carry?.startDeclarations[0]).toContain('10px 6.667px 0px rgba(0, 0, 0, 0)');
		expect(carry?.startDeclarations[0]).toContain('20px 20px 0px rgba(0, 0, 0, 1)');
	});

	it('fades in a shadow that changes where it stands', () => {
		const carry = buildMorphEffectCarry(
			rect('a', 200, 100, shadow),
			rect('b', 200, 100, { ...shadow, shadowDistance: Math.SQRT2 * 50, shadowBlur: 30 }),
			'to',
		);
		expect(carry?.startDeclarations[0]).toMatch(/^box-shadow: 50px 50px 30px rgba\(0, 0, 0, 0\)/u);
		expect(carry?.endDeclarations[0]).toMatch(/^box-shadow: 50px 50px 30px rgba\(0, 0, 0, 1\)/u);
	});

	it('leaves a pair alone when both ends land the same effects', () => {
		expect(buildMorphEffectCarry(rect('a', 200, 100, {}), rect('b', 400, 300, {}), 'to')).toBe(
			undefined,
		);
		expect(
			buildMorphEffectCarry(rect('a', 200, 100, shadow), rect('b', 200, 100, shadow), 'to'),
		).toBe(undefined);
	});

	it('widens a glow filter from the outgoing radius riding the box to the incoming one', () => {
		const glow = { glowColor: '#FFFF00', glowRadius: 20, glowOpacity: 0.75 };
		const carry = buildMorphEffectCarry(rect('a', 200, 100, glow), rect('b', 400, 200, glow), 'to');
		const start = carry?.startDeclarations.find((d) => d.startsWith('filter:'));
		const end = carry?.endDeclarations.find((d) => d.startsWith('filter:'));
		expect(start).toBeDefined();
		expect(end).toBeDefined();
		const radius = (value: string | undefined) =>
			Number(/drop-shadow\([^)]*?([\d.]+)px rgba/u.exec(value ?? '')?.[1]);
		expect(radius(start)).toBeCloseTo(radius(end) * 2, 1);
	});
});
