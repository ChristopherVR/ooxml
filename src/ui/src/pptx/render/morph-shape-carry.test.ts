import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphShapeCarry, morphShapeCarryApplies } from './morph-shape-carry';
import { buildMorphOutlineRules } from './morph-plan';
import type { MorphTransitionPlan } from './morph-plan';
import { buildStrokeOutline } from './stroke-outline';

function card(
	id: string,
	box: { x: number; y: number; width: number; height: number },
	extra: Record<string, unknown> = {},
): PptxElement {
	return {
		id,
		type: 'shape',
		shapeType: 'roundRect',
		...box,
		shapeStyle: { fillColor: '#C07020', strokeColor: '#FFFFFF', strokeWidth: 8 },
		...extra,
	} as unknown as PptxElement;
}

/** The value of `property` in every stop of a baked keyframes body. */
function stops(body: string, property: string): string[] {
	return [...body.matchAll(new RegExp(`\\t\\t${property}: ([^;]*);`, 'gu'))].map((m) => m[1] ?? '');
}

function scaleOf(value: string): [number, number] {
	const [, x, y] = /scale\(([^,]+), ([^)]+)\)/u.exec(value) ?? [];
	return [Number(x), Number(y)];
}

describe('morph shape carry', () => {
	const small = card('a', { x: 80, y: 80, width: 80, height: 80 });
	const large = card('b', { x: 400, y: 260, width: 640, height: 320 });

	it('applies to one preset whose box or handles change, and to nothing else', () => {
		expect(morphShapeCarryApplies(small, large)).toBe(true);
		expect(morphShapeCarryApplies(small, { ...small, x: 300 } as PptxElement)).toBe(false);
		expect(
			morphShapeCarryApplies(small, {
				...small,
				shapeAdjustments: { adj: 50000 },
			} as unknown as PptxElement),
		).toBe(true);
		// Another preset reaches the carry only as a `!!`-named pair, whose
		// outline PowerPoint morphs (see `morph-outline-blend`).
		expect(
			morphShapeCarryApplies(small, { ...large, shapeType: 'ellipse' } as unknown as PptxElement),
		).toBe(true);
		expect(morphShapeCarryApplies(small, { ...large, flipHorizontal: true } as PptxElement)).toBe(
			false,
		);
	});

	it('keeps the outline at the measured fraction of the growing box on each axis', () => {
		// PowerPoint: the stroke is lerp(w / W0, w / W1, p) of the box, so a
		// 60 pt square growing to 480 x 240 pt thickens mid-way (18.6 px measured
		// on the vertical edges at p 0.52 for an 8 px line), then settles.
		const carry = buildMorphShapeCarry(small, large, 'to');
		const scales = stops(carry?.outlineKeyframes ?? '', 'transform').map(scaleOf);
		const paths = stops(carry?.outlineKeyframes ?? '', 'd');
		expect(scales.length).toBeGreaterThan(13);
		scales.forEach(([ax, ay], i) => {
			// The x share lerps 8/80 -> 8/640, so ax = 8 - 7p in the 640 px box...
			const p = (8 - ax) / 7;
			// ...and the y share 8/80 -> 8/320 on the same p: ay = 4 - 3p.
			expect(ay).toBeCloseTo(4 - 3 * p, 2);
			// The path itself always spans the local box once scaled.
			const right = Math.max(
				...[...(paths[i] ?? '').matchAll(/L([\d.]+) /gu)].map((m) => Number(m[1])),
			);
			expect(right * ax).toBeCloseTo(640, -1);
		});
		expect(scales[0]).toEqual([8, 4]);
		expect(scales.at(-1)).toEqual([1, 1]);
	});

	it('lands the outline path on each end, in the half that carries it', () => {
		const carry = buildMorphShapeCarry(small, large, 'to');
		const paths = stops(carry?.outlineKeyframes ?? '', 'd');
		// At p = 1 the path is the incoming shape's own outline in its own box.
		expect(paths.at(-1)).toContain('L640 ');
		// At p = 0, scaled by (8, 4), the outgoing 80 x 80 outline fills the local box.
		expect(paths[0]).toContain('L80 ');
		const ghost = buildMorphShapeCarry(small, large, 'from');
		const ghostScales = stops(ghost?.outlineKeyframes ?? '', 'transform').map(scaleOf);
		expect(ghostScales[0]).toEqual([1, 1]);
		expect(ghostScales.at(-1)?.[0]).toBeCloseTo(0.125, 3);
	});

	it('lerps a rounded rectangle corner between the two normalised radii', () => {
		const from = card('a', { x: 0, y: 0, width: 400, height: 160 });
		const to = card('b', { x: 0, y: 0, width: 560, height: 240 });
		const carry = buildMorphShapeCarry(from, to, 'to');
		// 16667 of half the short side: 26.667 px, then 40 px.
		expect(carry?.startDeclarations).toEqual(['border-radius: 37.334px / 40.001px']);
		expect(carry?.endDeclarations).toEqual(['border-radius: 40.001px / 40.001px']);
	});

	it('re-evaluates a clipped preset per stop when its handles move', () => {
		const arrow = (id: string, adj: Record<string, number>) =>
			card(
				id,
				{ x: 0, y: 0, width: 400, height: 200 },
				{ shapeType: 'rightArrow', shapeAdjustments: adj },
			);
		const carry = buildMorphShapeCarry(
			arrow('a', { adj1: 30000, adj2: 30000 }),
			arrow('b', { adj1: 70000, adj2: 90000 }),
			'to',
		);
		const clips = stops(carry?.clipKeyframes ?? '', 'clip-path');
		expect(clips).toHaveLength(13);
		expect(new Set(clips).size).toBe(13);
		expect(clips.every((clip) => clip.startsWith("path('M"))).toBe(true);
		expect(carry?.outlineKeyframes).toBeDefined();
	});

	it('rides the morph curve, crowding stops where the outline share changes fastest', () => {
		const carry = buildMorphShapeCarry(small, large, 'to');
		const times = [...(carry?.outlineKeyframes ?? '').matchAll(/\t([\d.]+)% \{/gu)].map((m) =>
			Number(m[1]),
		);
		expect(times[0]).toBe(0);
		expect(times.at(-1)).toBe(100);
		expect(times.every((time, i) => i === 0 || time > (times[i - 1] ?? 0))).toBe(true);
		// Every stop but the last hands its slice of the morph curve to the next.
		expect(carry?.outlineKeyframes?.match(/animation-timing-function/gu)).toHaveLength(
			times.length - 1,
		);
		// A box that only grows 1.4x needs only a few more than the even twelfths.
		const plain = buildMorphShapeCarry(
			card('a', { x: 0, y: 0, width: 400, height: 160 }),
			card('b', { x: 0, y: 0, width: 560, height: 240 }),
			'to',
		);
		const count = (body: string | undefined) => body?.match(/% \{/gu)?.length ?? 0;
		expect(count(plain?.outlineKeyframes)).toBeLessThanOrEqual(17);
		expect(count(carry?.outlineKeyframes)).toBeGreaterThan(30);
	});

	describe('compound outlines', () => {
		const compound = (id: string, kind: string, box: { width: number; height: number }) =>
			card(
				id,
				{ x: 0, y: 0, ...box },
				{
					shapeStyle: {
						fillColor: '#C07020',
						strokeColor: '#FFFFFF',
						strokeWidth: 16,
						compoundLine: kind,
					},
				},
			);
		const fromBox = { width: 160, height: 107 };
		const toBox = { width: 640, height: 320 };

		it.each([
			['dbl', 2],
			['thickThin', 2],
			['thinThick', 2],
			['tri', 3],
		])('carries every strand of a %s line as its own track', (kind, strands) => {
			const from = compound('a', kind, fromBox);
			const to = compound('b', kind, toBox);
			const carry = buildMorphShapeCarry(from, to, 'to');
			expect(carry?.outlineKeyframes).toBeUndefined();
			const tracks = carry?.outlineStrandKeyframes ?? [];
			expect(tracks).toHaveLength(strands);
			// Every strand follows the plain line's measured scale on each axis.
			const plain = buildMorphShapeCarry(
				card(
					'a',
					{ x: 0, y: 0, ...fromBox },
					{ shapeStyle: { strokeWidth: 16, strokeColor: '#FFF' } },
				),
				card(
					'b',
					{ x: 0, y: 0, ...toBox },
					{ shapeStyle: { strokeWidth: 16, strokeColor: '#FFF' } },
				),
				'to',
			);
			const expected = stops(plain?.outlineKeyframes ?? '', 'transform');
			for (const track of tracks) {
				expect(stops(track, 'transform')).toEqual(expected);
			}
			// At the end each strand is the static strand of the incoming shape.
			const staticStrands = buildStrokeOutline(to)?.strands ?? [];
			tracks.forEach((track, i) => {
				const last = stops(track, 'd').at(-1) ?? '';
				const numbers = (d: string) => (d.match(/-?\d+\.?\d*/gu) ?? []).map(Number);
				// The centre strand of a `tri` line is the outline itself (no traced path).
				if (staticStrands[i]?.d === undefined) return;
				const want = numbers(staticStrands[i]?.d ?? '');
				const got = numbers(last);
				expect(got).toHaveLength(want.length);
				got.forEach((value, k) => expect(value).toBeCloseTo(want[k] ?? 0, 1));
				// A strand keeps its number of points, so CSS can tween it.
				const counts = new Set(stops(track, 'd').map((d) => d.split('L').length));
				expect(counts.size).toBe(1);
			});
		});

		it('leaves a pair whose line kind changes to the box journey', () => {
			const carry = buildMorphShapeCarry(
				compound('a', 'dbl', fromBox),
				compound('b', 'tri', toBox),
				'to',
			);
			expect(carry?.outlineStrandKeyframes).toBeUndefined();
		});

		it('addresses each strand path by position in the plan rules', () => {
			const plan = {
				outlineStrandAnimations: new Map([['a-1', ['t0 1s', 't1 1s']]]),
			} as unknown as MorphTransitionPlan;
			const rules = buildMorphOutlineRules(plan, 'data-element-id');
			expect(rules).toContain('[data-pptx-outline] path:nth-of-type(1) { animation: t0 1s; }');
			expect(rules).toContain('[data-pptx-outline] path:nth-of-type(2) { animation: t1 1s; }');
		});
	});
});
