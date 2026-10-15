import { describe, expect, it } from 'vitest';

import {
	buildHierarchyRoots,
	hierarchyBranchLabels,
	layoutTreemap,
	normalizeLevels,
} from './hierarchy';
import { computeHierarchicalSunburstArcs, computeSunburstArcs } from './sunburst-arcs';

const leaves = ['a', 'b', 'c', 'd', 'e'];
const levels = [leaves, ['Fruit', '', 'Veg', '', 'Nuts']];
const values = [30, 20, 10, 5, 8];

describe('hierarchy', () => {
	it('fills parent labels written once down their rows', () => {
		expect(normalizeLevels(levels, 5)[1]).toEqual(['Fruit', 'Fruit', 'Veg', 'Veg', 'Nuts']);
	});

	it('builds one root per outermost branch with aggregated weights and one colour each', () => {
		const roots = buildHierarchyRoots([{ name: 'Size', values }], leaves, levels);
		expect(roots.map((r) => [r.label, r.weight, r.colorIndex])).toEqual([
			['Fruit', 50, 0],
			['Veg', 15, 1],
			['Nuts', 8, 2],
		]);
		expect(roots[0]?.children?.map((c) => [c.label, c.colorIndex])).toEqual([
			['a', 0],
			['b', 0],
		]);
	});

	it('wraps several series as roots of their own', () => {
		const roots = buildHierarchyRoots(
			[
				{ name: 'One', values: [3, 1] },
				{ name: 'Two', values: [2, 2] },
			],
			['x', 'y'],
			undefined,
		);
		expect(roots.map((r) => [r.label, r.weight])).toEqual([
			['One', 4],
			['Two', 4],
		]);
		expect(
			hierarchyBranchLabels(
				[
					{ name: 'One', values: [] },
					{ name: 'Two', values: [] },
				],
				[],
				undefined,
			),
		).toEqual(['One', 'Two']);
	});

	it('lists the unique outermost level as the legend of one series', () => {
		expect(hierarchyBranchLabels([{ name: 'Size', values }], leaves, levels)).toEqual([
			'Fruit',
			'Veg',
			'Nuts',
		]);
		expect(hierarchyBranchLabels([{ name: 'Size', values }], leaves, undefined)).toEqual(leaves);
	});

	it('lays leaves out inside the box with a parent banner label', () => {
		const roots = buildHierarchyRoots([{ name: 'Size', values }], leaves, levels);
		const cells = layoutTreemap(roots, { x: 0, y: 0, w: 400, h: 300 });
		const leafCells = cells.filter((c) => c.kind === 'leaf');
		expect(leafCells).toHaveLength(5);
		expect(cells.filter((c) => c.kind === 'parent').map((c) => c.label)).toContain('Fruit');
		const area = leafCells.reduce((sum, c) => sum + c.w * c.h, 0);
		expect(area).toBeGreaterThan(0.8 * 400 * 300 - 5 * 16 * 400);
		for (const c of leafCells) {
			expect(c.x).toBeGreaterThanOrEqual(0);
			expect(c.x + c.w).toBeLessThanOrEqual(401);
			expect(c.y + c.h).toBeLessThanOrEqual(301);
		}
	});
});

describe('sunburst arcs', () => {
	it('draws one ring per series with one wedge per point', () => {
		const arcs = computeSunburstArcs([{ values: [1, 1] }, { values: [1, 2, 1] }], 100, 100, 80);
		expect(arcs.map((a) => a.level)).toEqual([0, 0, 1, 1, 1]);
		expect(arcs[0]?.opacity).toBeCloseTo(0.9);
	});

	it('draws one ring per level, aggregating contiguous leaves into parents', () => {
		const arcs = computeHierarchicalSunburstArcs(normalizeLevels(levels, 5), values, 100, 100, 90);
		expect(arcs.filter((a) => a.level === 0).map((a) => a.label)).toEqual(['Fruit', 'Veg', 'Nuts']);
		expect(arcs.filter((a) => a.level === 1)).toHaveLength(5);
		expect(arcs.filter((a) => a.pointIndex !== undefined)).toHaveLength(5);
		// Leaves under one parent share its colour.
		const fruit = arcs.filter((a) => a.label === 'a' || a.label === 'b').map((a) => a.colorIndex);
		expect(new Set(fruit).size).toBe(1);
	});

	it('keeps labels upright', () => {
		for (const arc of computeSunburstArcs([{ values: [1, 1, 1, 1, 1, 1] }], 0, 0, 50))
			expect(arc.labelRotation > 90 && arc.labelRotation < 270).toBe(false);
	});
});
