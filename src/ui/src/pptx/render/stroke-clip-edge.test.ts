import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { outerStrokeEdgePath } from './stroke-clip-edge';
import { buildStrokeOutline } from './stroke-outline';

const points = (d: string): Array<[number, number]> => {
	const nums = (d.match(/-?\d*\.?\d+/gu) ?? []).map(Number);
	const out: Array<[number, number]> = [];
	for (let i = 0; i + 1 < nums.length; i += 2) {
		out.push([nums[i] as number, nums[i + 1] as number]);
	}
	return out;
};

describe('outerStrokeEdgePath', () => {
	const miter = { join: 'miter', miterLimit: 4 } as const;

	it('traces the widened clip: the square pushed out by half the line, mitred', () => {
		const ring = points(outerStrokeEdgePath('M 0 0 L 10 0 L 10 10 L 0 10 Z', 4, miter) ?? '');
		expect(ring).toEqual([
			[-2, -2],
			[12, -2],
			[12, 12],
			[-2, 12],
		]);
	});

	it('bevels a corner beyond the mitre limit, as the clip does', () => {
		const ring = points(
			outerStrokeEdgePath('M 0 0 L 10 0 L 10 10 L 0 10 Z', 4, { join: 'miter', miterLimit: 1 }) ??
				'',
		);
		expect(ring).toHaveLength(8);
		expect(ring[0]).toEqual([-2, 0]);
		expect(ring[1]).toEqual([0, -2]);
	});

	it('offsets away from the fill whichever way the outline is wound', () => {
		const ring = points(outerStrokeEdgePath('M 0 0 L 0 10 L 10 10 L 10 0 Z', 4, miter) ?? '');
		expect(ring[0]).toEqual([-2, -2]);
	});

	it('meets a concave corner at the crossing of its two edges', () => {
		// The notch's inner corner (5, 5): the outer edge passes 1px outside both
		// of its edges, at (6, 4), rather than mitring into the shape.
		const ring = points(
			outerStrokeEdgePath('M 0 0 L 5 0 L 5 5 L 10 5 L 10 10 L 0 10 Z', 2, miter) ?? '',
		);
		expect(ring).toContainEqual([6, 4]);
	});
});

describe('the outline edge bleed (issue #144)', () => {
	const shape = (overrides: object = {}, type = 'shape'): PptxElement =>
		({
			id: 'm',
			type,
			shapeType: 'flowChartOffpageConnector',
			x: 0,
			y: 0,
			width: 15,
			height: 26,
			shapeStyle: { strokeColor: '#ffffff', strokeWidth: 1.333, lineJoin: 'miter', ...overrides },
		}) as unknown as PptxElement;

	it('strokes the widened clip edge so the fill cannot show along it', () => {
		const outline = buildStrokeOutline(shape());
		const bleed = outline?.strands.find((strand) => strand.d);
		expect(bleed).toBeDefined();
		expect(bleed?.d).not.toBe(outline?.d);
		expect(bleed?.cssStrokeWidth).toContain('--pptx-device-px');
		// The outline itself is still the first strand, unchanged.
		expect(outline?.strands[0]?.d).toBeUndefined();
	});

	it('adds nothing where the clip is not widened or a second pass would show', () => {
		for (const element of [
			shape({ lineAlignment: 'in' }),
			shape({ strokeDash: 'dash' }),
			shape({ strokeOpacity: 0.5 }),
			{ ...shape(), shapeType: 'rect' } as PptxElement,
		]) {
			expect(buildStrokeOutline(element)?.strands.some((strand) => strand.d) ?? false).toBe(false);
		}
	});
});
