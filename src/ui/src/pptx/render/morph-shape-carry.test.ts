import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphShapeCarry, morphShapeCarryApplies } from './morph-shape-carry';

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
		expect(
			morphShapeCarryApplies(small, { ...large, shapeType: 'ellipse' } as unknown as PptxElement),
		).toBe(false);
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
});
