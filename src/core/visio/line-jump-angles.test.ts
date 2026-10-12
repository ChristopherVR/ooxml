import { describe, expect, it } from 'vitest';
import evidence from './__fixtures__/native-line-jump-angles.json';
import type { VisioPage, VisioShape } from './model';
import { visioLineJumpPaths } from './line-jumps';

describe('native Visio slanted connector jumps', () => {
	// Native SVG rounds coordinates and the 45-degree classifier is sensitive to
	// floating-point endpoint rounding. Keep those observations in the corpus,
	// without asserting an exact boundary tie rule from rounded output.
	const cases = evidence.cases.filter((row) => row.curved || row.angle % 90 !== 45);
	it.each(cases)('matches curve=$curved angle=$angle peer=$peerAngle page=$pageCode', (row) => {
		const shapes = row.ends.map((end, index) => {
			const [x1, y1, x2, y2] = end as [number, number, number, number];
			const path = row.curved
				? `M ${x1} ${y1} C ${x1} ${y1} ${x2} ${y2} ${x2} ${y2}`
				: `M ${x1} ${y1} L ${x2} ${y2}`;
			return {
				id: String(index),
				name: String(index),
				kind: 'connector',
				width: 4,
				height: 4,
				transform: [1, 0, 0, 1, 0, 0],
				geometry: [{ path, fill: false, stroke: true }],
				connectorRoute: row.curved ? 'curved' : 'straight',
				children: [],
				hidden: false,
			} as unknown as VisioShape;
		});
		const page = {
			id: '0',
			name: 'Page',
			width: 10,
			height: 10,
			isBackground: false,
			shapes,
			connectors: [],
			layout: {
				lineJumpCode: row.pageCode,
				lineJumpStyle: 1,
				lineJumpFactorX: evidence.factorX,
				lineJumpFactorY: evidence.factorY,
			},
		} as unknown as VisioPage;
		const paths = visioLineJumpPaths(page);
		for (const [index, shape] of shapes.entries()) {
			const native = row.paths[index]!.match(/A([\d.]+)/);
			const actual = paths.get(shape)?.[0]?.match(/A ([\d.]+)/);
			expect(Boolean(actual)).toBe(Boolean(native));
			if (native && actual) expect(Number(actual[1]) * 72).toBeCloseTo(Number(native[1]), 3);
		}
	});
});
