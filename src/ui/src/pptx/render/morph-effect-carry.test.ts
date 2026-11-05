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

/** The overlap layers of the 200 x 100 -> 400 x 300 pair below, in the incoming frame. */
const OVERLAP = (alpha: number) =>
	`20px 40px 0px -20px rgba(0, 0, 0, ${alpha}), 40px 40px 0px -20px rgba(0, 0, 0, ${alpha})`;

/**
 * The declaration of `property` at eased progress `percent` (0, 50 or 100) of
 * a baked body: its keyframes sit at stops 0, 0.1, ... 1 of the eased curve.
 */
function at(body: string | undefined, percent: number, property: string): string | undefined {
	const blocks = [...(body ?? '').matchAll(/\t[\d.]+% \{([^}]*)\}/gu)].map((m) => m[1] ?? '');
	const block = blocks[Math.round(((blocks.length - 1) * percent) / 100)];
	return new RegExp(`${property}: ([^;]*);`, 'u').exec(block ?? '')?.[1];
}

describe('morph effect carry', () => {
	it('crossfades the two ends shadows, each riding the box, with their overlap held opaque', () => {
		// PowerPoint: one shadow offset 20 * W(p) / W0 fading out on 1 - p, one
		// offset 20 * W(p) / W1 fading in on p, and alpha 1 where they overlap.
		const carry = buildMorphEffectCarry(
			rect('a', 200, 100, shadow),
			rect('b', 400, 300, shadow),
			'to',
		);
		// Overlap of (20, 20) and (40, 60) on 400 x 300: 380 x 260, tiled by
		// two 360 x 260 boxes (spread -20) side by side.
		expect(at(carry?.keyframes, 0, 'box-shadow')).toBe(
			`${OVERLAP(0)}, 20px 20px 0px rgba(0, 0, 0, 0), 40px 60px 0px rgba(0, 0, 0, 1)`,
		);
		expect(at(carry?.keyframes, 100, 'box-shadow')).toBe(
			`${OVERLAP(0)}, 20px 20px 0px rgba(0, 0, 0, 1), 40px 60px 0px rgba(0, 0, 0, 0)`,
		);
		const mid = at(carry?.keyframes, 50, 'box-shadow') ?? '';
		expect(mid.startsWith(`${OVERLAP(1)}, `)).toBe(true);
	});

	it('writes the ghost half in its own frame', () => {
		const carry = buildMorphEffectCarry(
			rect('a', 200, 100, shadow),
			rect('b', 400, 300, shadow),
			'from',
		);
		const start = at(carry?.keyframes, 0, 'box-shadow');
		expect(start).toContain('10px 6.667px 0px rgba(0, 0, 0, 0)');
		expect(start).toContain('20px 20px 0px rgba(0, 0, 0, 1)');
	});

	it('fades in a shadow that changes where it stands', () => {
		const carry = buildMorphEffectCarry(
			rect('a', 200, 100, shadow),
			rect('b', 200, 100, { ...shadow, shadowDistance: Math.SQRT2 * 50, shadowBlur: 30 }),
			'to',
		);
		expect(at(carry?.keyframes, 0, 'box-shadow')).toContain('50px 50px 30px rgba(0, 0, 0, 0)');
		expect(at(carry?.keyframes, 100, 'box-shadow')).toContain('50px 50px 30px rgba(0, 0, 0, 1)');
	});

	it('leaves a pair alone when both ends land the same effects', () => {
		expect(buildMorphEffectCarry(rect('a', 200, 100, {}), rect('b', 400, 300, {}), 'to')).toBe(
			undefined,
		);
		expect(
			buildMorphEffectCarry(rect('a', 200, 100, shadow), rect('b', 200, 100, shadow), 'to'),
		).toBe(undefined);
	});

	it('dissolves two glows, each riding the box, as a chain narrow glow first', () => {
		const glow = { glowColor: '#FFFF00', glowRadius: 20, glowOpacity: 0.75 };
		const carry = buildMorphEffectCarry(rect('a', 200, 100, glow), rect('b', 400, 200, glow), 'to');
		const radii = (value: string | undefined) =>
			[
				...(value ?? '').matchAll(
					/drop-shadow\(0px 0px ([\d.]+)px rgba\(255, 255, 0, ([\d.]+)\)\)/gu,
				),
			].map((m) => [Number(m[1]), Number(m[2])]);
		// In the incoming frame the outgoing glow rides at twice its radius, so
		// the incoming (narrower) one is painted first.
		expect(radii(at(carry?.keyframes, 0, 'filter'))).toEqual([
			[10, 0],
			[20, 0.75],
		]);
		expect(radii(at(carry?.keyframes, 100, 'filter'))).toEqual([
			[10, 0.75],
			[20, 0],
		]);
	});

	it('fades only the incoming effects in under a ghost that paints the outgoing ones', () => {
		const carry = buildMorphEffectCarry(
			rect('a', 200, 100, shadow),
			rect('b', 400, 300, shadow),
			'to',
			'underGhost',
		);
		const start = at(carry?.keyframes, 0, 'box-shadow');
		// No outgoing layer: the overlap layer and the incoming shadow only.
		expect(start).toBe(`${OVERLAP(0)}, 20px 20px 0px rgba(0, 0, 0, 0)`);
		// Opaque where the ghost's shadow lies over it, at every stop.
		expect(at(carry?.keyframes, 50, 'box-shadow')?.startsWith(OVERLAP(1))).toBe(true);
		expect(at(carry?.keyframes, 100, 'box-shadow')).toBe(
			`${OVERLAP(0)}, 20px 20px 0px rgba(0, 0, 0, 1)`,
		);
	});
});
