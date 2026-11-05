import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { blendOutlineCommands } from './morph-outline-blend';
import { buildMorphShapeCarry, morphShapeCarryApplies } from './morph-shape-carry';
import { parseSvgPath, serializeSvgPath } from './morph-svg-path';

function freeform(id: string, d: string, extra: Record<string, unknown> = {}): PptxElement {
	return {
		id,
		type: 'shape',
		shapeType: 'custom',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		pathData: d,
		pathWidth: 100,
		pathHeight: 100,
		shapeStyle: { fillColor: '#C07020', strokeColor: '#FFFFFF', strokeWidth: 8 },
		...extra,
	} as unknown as PptxElement;
}

describe('morph outline blend', () => {
	it('pairs two rings of different vertex counts, each keeping its own corners', () => {
		// A square (4 corners) and a hexagon-ish ring (6): every corner of either
		// is a vertex of both blended rings, at the same arc-length fraction.
		const square = parseSvgPath('M 0 0 L 100 0 L 100 100 L 0 100 Z');
		const six = parseSvgPath('M 0 0 L 50 -20 L 100 0 L 100 100 L 50 120 L 0 100 Z');
		const [a, b] = blendOutlineCommands(square, six) ?? [[], []];
		expect(a).toHaveLength(b.length);
		const points = (commands: typeof a) =>
			commands.filter((c) => c.type !== 'Z').map((c) => `${c.values[0]},${c.values[1]}`);
		for (const corner of ['0,0', '100,0', '100,100', '0,100']) {
			expect(points(a)).toContain(corner);
		}
		for (const corner of ['0,0', '50,-20', '100,0', '100,100', '50,120', '0,100']) {
			expect(points(b)).toContain(corner);
		}
	});

	it('walks both outlines in path order from their own first point', () => {
		// PowerPoint 16: a `!!` rectangle's first corner (top left) walks to the
		// ellipse's first point (its left end), not to the nearest point.
		const rect = parseSvgPath('M 0 0 L 100 0 L 100 50 L 0 50 Z');
		const ellipse = parseSvgPath(
			'M 0 25 A 50 25 0 0 1 50 0 A 50 25 0 0 1 100 25 A 50 25 0 0 1 50 50 A 50 25 0 0 1 0 25 Z',
		);
		const [a, b] = blendOutlineCommands(rect, ellipse) ?? [[], []];
		expect(a[0]?.values).toStrictEqual([0, 0]);
		expect(b[0]?.values).toStrictEqual([0, 25]);
	});

	it('refuses an outline of several sub-paths', () => {
		const two = parseSvgPath('M 0 0 L 10 0 L 10 10 Z M 20 20 L 30 20 L 30 30 Z');
		expect(blendOutlineCommands(two, parseSvgPath('M 0 0 L 9 0 L 9 9 Z'))).toBeUndefined();
	});

	it('carries a freeform whose nodes move, and one that gains nodes', () => {
		const from = freeform('a', 'M 0 0 L 100 0 L 80 100 L 0 60 Z');
		const moved = freeform('b', 'M 0 0 L 100 0 L 60 100 L 0 90 Z');
		const grown = freeform('c', 'M 0 0 L 100 0 L 100 40 L 80 100 L 20 100 L 0 60 Z');
		expect(morphShapeCarryApplies(from, from)).toBe(false);
		expect(morphShapeCarryApplies(from, moved)).toBe(true);
		expect(morphShapeCarryApplies(from, grown)).toBe(true);

		const carry = buildMorphShapeCarry(from, grown, 'to');
		// The fill is clipped to the walking outline, the line re-pathed, and a
		// freeform painted as a vector re-pathed in its own path units.
		expect(carry?.clipKeyframes).toContain('clip-path: path(');
		expect(carry?.outlineKeyframes).toContain('d: path(');
		expect(carry?.customKeyframes).toContain('d: path(');
		const first = /d: path\('([^']*)'\)/u.exec(carry?.customKeyframes ?? '')?.[1] ?? '';
		// At the start it is the outgoing outline, its corners kept, with the
		// incoming ring's extra vertices inserted along its edges.
		const start = serializeSvgPath(parseSvgPath(first));
		for (const corner of ['M0 0', 'L100 0', 'L80 100', 'L0 60']) {
			expect(start).toContain(corner);
		}
		expect(parseSvgPath(first).filter((c) => c.type === 'L').length).toBeGreaterThan(5);
	});
});
