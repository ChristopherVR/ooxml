import { describe, expect, it } from 'vitest';
import { roundedOrthogonalPath, type RectanglePoint as Point } from './rounded-geometry.js';
import { geometryPaths } from './geometry.js';
import { readSheet } from './sheet.js';
import { parseXml } from '../xml/index.js';
import { cell, row, section, xml } from './test-fixtures.js';

// Independently observed Visio 16 SVG radii in points, divided by 72.
// Reproduce with scripts/verify-visio-rounding.ps1; no native document is shipped.
const cases: { name: string; points: Point[]; radii: number[] }[] = [
	{
		name: 'short shared segment',
		points: [
			[0, 0],
			[2, 0],
			[2, 0.2],
			[4, 0.2],
		],
		radii: [0.1, 0.1],
	},
	{
		name: 'short endpoint',
		points: [
			[0, 0],
			[0.1, 0],
			[0.1, 2],
		],
		radii: [0.05],
	},
	{
		name: 'endpoint reservation',
		points: [
			[0, 0],
			[0.4, 0],
			[0.4, 2],
		],
		radii: [0.2],
	},
	{
		name: 'unequal available space',
		points: [
			[0, 0],
			[0.05, 0],
			[0.05, 0.2],
			[2, 0.2],
		],
		radii: [0.025, 0.175],
	},
	{
		name: 'reversed unequal space',
		points: [
			[2, 0.2],
			[0.05, 0.2],
			[0.05, 0],
			[0, 0],
		],
		radii: [0.1, 0.025],
	},
	{
		name: 'two short endpoints',
		points: [
			[0, 0],
			[0.05, 0],
			[0.05, 0.2],
			[0.1, 0.2],
		],
		radii: [0.025, 0.025],
	},
	{
		name: 'three dependent corners',
		points: [
			[0, 0],
			[2, 0],
			[2, 0.2],
			[2.3, 0.2],
			[2.3, 2],
		],
		radii: [0.1, 0.1, 0.2],
	},
	{
		name: 'collinear vertex reservation',
		points: [
			[0, 0],
			[1.95, 0],
			[2, 0],
			[2, 1],
			[4, 1],
		],
		radii: [0.05, 0.25],
	},
	{
		name: 'later unbounded corners',
		points: [
			[0, 0],
			[2, 0],
			[2, 0.2],
			[4, 0.2],
			[4, 2],
			[6, 2],
		],
		radii: [0.1, 0.1, 0.25, 0.25],
	},
];
const serialize = (x: number, y: number) => `${Number(x.toFixed(9))} ${Number(y.toFixed(9))}`;

describe('native-observed ordered connector clamping', () => {
	it.each(cases)('$name agrees with native radii and retains endpoints', ({ points, radii }) => {
		for (const [sx, sy] of [
			[1, 1],
			[-1, 1],
			[1, -1],
			[-1, -1],
		]) {
			const transformed = points.map(([x, y]) => [x * sx! + 3, y * sy! - 7] as const);
			const rounded = roundedOrthogonalPath(transformed, 0.25, serialize)!;
			const arcs = [
				...rounded.path.matchAll(/A ([\d.]+) ([\d.]+) 0 0 ([01]) ([\d.\-]+) ([\d.\-]+)/g),
			];
			expect(arcs.map((arc) => Number(arc[1]))).toEqual(radii);
			expect(arcs.map((arc) => Number(arc[2]))).toEqual(radii);
			expect(rounded.path.startsWith(`M ${serialize(...transformed[0]!)} `)).toBe(true);
			expect(rounded.path.endsWith(`L ${serialize(...transformed.at(-1)!)}`)).toBe(true);
			expect(rounded.extraCommands).toBe(radii.length);
		}
	});
	it('normalizes relative coordinates, charges arcs and preserves section paint', () => {
		const contents =
			cell('Rounding', 0.25) +
			section(
				'Geometry',
				cell('NoFill', 1) +
					[
						[0, 0],
						[0.5, 0],
						[0.5, 0.1],
						[1, 0.1],
					]
						.map(([x, y], i) =>
							row(i, i ? 'RelLineTo' : 'RelMoveTo', cell('X', x!) + cell('Y', y!)),
						)
						.join(''),
			);
		const warnings: string[] = [];
		let consumed = 0;
		const sheet = readSheet(parseXml(xml('Shape', contents)).documentElement);
		const paths = geometryPaths(
			sheet,
			4,
			2,
			(code) => warnings.push(code),
			() => {
				consumed++;
			},
		);
		expect(paths).toEqual([
			{
				path: 'M 0 0 L 1.9 0 A 0.1 0.1 0 0 1 2 0.1 L 2 0.1 A 0.1 0.1 0 0 0 2.1 0.2 L 4 0.2',
				fill: false,
				stroke: true,
			},
		]);
		expect(consumed).toBe(6);
		expect(warnings).toEqual([]);
		expect(() =>
			geometryPaths(
				sheet,
				4,
				2,
				() => {},
				() => {
					if (++consumed > 11) throw new Error('budget');
				},
			),
		).toThrow('budget');
	});
	it('declines a clamped arc below serialization precision', () => {
		expect(
			roundedOrthogonalPath(
				[
					[0, 0],
					[1e-10, 0],
					[1e-10, 1],
				],
				0.25,
				serialize,
			),
		).toBeUndefined();
	});
});
