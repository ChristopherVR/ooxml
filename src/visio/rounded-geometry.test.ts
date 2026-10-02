import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index.js';
import { geometryPaths } from './geometry.js';
import { readSheet } from './sheet.js';
import { parseVsdx } from './index.js';
import { cell, fixture, rectangle, row, section, shape, xml } from './test-fixtures.js';

type Point = readonly [number, number];
const ccw: Point[] = [
	[0, 0],
	[4, 0],
	[4, 2],
	[0, 2],
];
const rows = (points: readonly Point[], relative = false) =>
	points
		.map((p, i) =>
			row(
				i + 1,
				(relative ? 'Rel' : '') + (i ? 'LineTo' : 'MoveTo'),
				cell('X', p[0]) + cell('Y', p[1]),
			),
		)
		.join('');
function parse(contents: string, rounding: number | string, width = 4, height = 2) {
	const sheet = readSheet(
		parseXml(xml('Shape', cell('Rounding', rounding) + contents)).documentElement,
	);
	const warnings: string[] = [];
	let commands = 0;
	const paths = geometryPaths(
		sheet,
		width,
		height,
		(code) => warnings.push(code),
		() => {
			commands++;
		},
	);
	return { paths, warnings, commands };
}
const rect = (points = ccw, relative = false) =>
	section('Geometry', rows([...points, points[0]!], relative));

describe('bounded cached rectangle rounding', () => {
	it('does not round a rectangle inferred from invalid coordinate defaults', () => {
		const source = rect().replace('N="X" V="0"', 'N="X" V="NaN"');
		const result = parse(source, 0.3);
		expect(result.paths).toEqual(parse(source, 0).paths);
		expect(result.warnings).toContain('missing-cached-value');
		expect(result.warnings).toContain('unsupported-corner-rounding');
	});
	it('emits exact circular arcs with the original CCW winding', () => {
		const result = parse(rect(), 0.5);
		expect(result.paths[0]!.path).toBe(
			'M 0.5 0 L 3.5 0 A 0.5 0.5 0 0 1 4 0.5 L 4 1.5 A 0.5 0.5 0 0 1 3.5 2 L 0.5 2 A 0.5 0.5 0 0 1 0 1.5 L 0 0.5 A 0.5 0.5 0 0 1 0.5 0 Z',
		);
		expect(result.commands).toBe(10);
		expect(result.warnings).toEqual([]);
	});
	it.each([0, 1, 2, 3])(
		'preserves both windings with starting corner %i and translated local origin',
		(start) => {
			for (const reverse of [false, true]) {
				const base = (reverse ? [...ccw].reverse() : ccw).map(([x, y]) => [x - 7, y + 3] as const);
				const rotated = [...base.slice(start), ...base.slice(0, start)];
				const result = parse(rect(rotated), 0.25);
				const first = rotated[0]!,
					next = rotated[1]!;
				const dx = Math.sign(next[0] - first[0]),
					dy = Math.sign(next[1] - first[1]);
				expect(
					result.paths[0]!.path.startsWith(`M ${first[0] + dx * 0.25} ${first[1] + dy * 0.25} `),
				).toBe(true);
				expect(
					result.paths[0]!.path.match(new RegExp(`A 0.25 0.25 0 0 ${reverse ? 0 : 1} `, 'g')),
				).toHaveLength(4);
				expect(result.paths[0]!.path.endsWith(' Z')).toBe(true);
				expect(result.warnings).toEqual([]);
			}
		},
	);
	it('normalizes relative coordinates before calculating an inch-based radius', () => {
		const relative: Point[] = [
			[0.25, 0.25],
			[0.75, 0.25],
			[0.75, 0.75],
			[0.25, 0.75],
		];
		const result = parse(rect(relative, true), 0.25, 8, 4);
		expect(result.paths[0]!.path).toContain('M 2.25 1 L 5.75 1 A 0.25 0.25 0 0 1 6 1.25');
	});
	it.each([1, 10, 1e9])('clamps requested radius %s to half the shortest edge', (radius) => {
		const result = parse(rect(), radius);
		expect(result.paths[0]!.path.match(/A 1 1 /g)).toHaveLength(4);
		expect(result.paths[0]!.path).toContain('L 4 1 A 1 1');
		expect(result.warnings).toEqual([]);
	});
	it('rounds a square to four circle quarters at the clamped limit', () => {
		const result = parse(
			rect([
				[0, 0],
				[2, 0],
				[2, 2],
				[0, 2],
			]),
			100,
		);
		expect(result.paths[0]!.path.match(/A 1 1 /g)).toHaveLength(4);
		expect(result.paths[0]!.path.startsWith('M 1 0 L 1 0')).toBe(true);
	});
	it.each([0, -0, -1, 'NaN', 'Themed', 1e-12])(
		'safely retains unrounded geometry for radius %s',
		(radius) => {
			const result = parse(rect(), radius);
			expect(result.paths[0]!.path).toBe('M 0 0 L 4 0 L 4 2 L 0 2 L 0 0');
			if (Number(radius) < 0) expect(result.warnings).toContain('invalid-corner-rounding');
			if (radius === 1e-12) expect(result.warnings).toContain('unsupported-corner-rounding');
		},
	);
	it.each(
		(
			[
				[
					[0, 0],
					[4, 0],
					[3, 2],
					[0, 2],
					[0, 0],
				],
				[
					[0, 0],
					[4, 1e-10],
					[4, 2],
					[0, 2],
					[0, 0],
				],
				[
					[0, 0],
					[4, 0],
					[4, 0],
					[0, 2],
					[0, 0],
				],
				[
					[0, 0],
					[2, 0],
					[4, 0],
					[4, 2],
					[0, 0],
				],
			] as Point[][]
		).map((points) => ({ points })),
	)('leaves diagonal, zero-length and nonrectangular paths unchanged', ({ points }) => {
		const source = section('Geometry', rows(points));
		const result = parse(source, 0.3);
		expect(result.paths).toEqual(parse(source, 0).paths);
		expect(result.warnings).toContain('unsupported-corner-rounding');
	});
	it('does not replace curved or multiple-subpath geometry with a rounded rectangle', () => {
		for (const source of [
			section(
				'Geometry',
				rows([ccw[0]!, ccw[1]!]) + row(3, 'ArcTo', cell('X', 4) + cell('Y', 2) + cell('A', 0.5)),
			),
			section('Geometry', rows([...ccw, ccw[0]!]) + row(6, 'MoveTo', cell('X', 5) + cell('Y', 5))),
		]) {
			const result = parse(source, 0.3);
			expect(result.paths).toEqual(parse(source, 0).paths);
			expect(result.warnings).toContain('unsupported-corner-rounding');
		}
	});
	it('preserves fill/stroke switches and skips hidden/deleted sections', () => {
		const result = parse(
			section('Geometry', cell('NoFill', 1) + cell('NoLine', 1) + rows([...ccw, ccw[0]!])) +
				section('Geometry', cell('NoShow', 1) + rows(ccw), 1) +
				'<Section N="Geometry" IX="2" Del="1"/>',
			0.2,
		);
		expect(result.paths).toHaveLength(1);
		expect(result.paths[0]).toMatchObject({ fill: false, stroke: false });
		expect(result.warnings).toEqual([]);
	});
	it('can round one section while warning and preserving another unsupported section', () => {
		const source =
			rect() +
			section(
				'Geometry',
				rows([
					[0, 0],
					[4, 1],
					[4, 2],
				]),
				1,
			);
		const result = parse(source, 0.3);
		expect(result.paths[0]!.path).toContain('A 0.3 0.3');
		expect(result.paths[1]).toEqual(parse(source, 0).paths[1]);
		expect(result.warnings).toContain('unsupported-corner-rounding');
	});
	it('ignores deleted geometry rows when validating the rectangle', () => {
		const source = section(
			'Geometry',
			rows(ccw) +
				'<Row IX="5" T="SplineStart" Del="1"/>' +
				row(6, 'LineTo', cell('X', 0) + cell('Y', 0)),
		);
		const result = parse(source, 0.2);
		expect(result.paths[0]!.path.match(/A 0.2 0.2 /g)).toHaveLength(4);
		expect(result.warnings).toEqual([]);
	});
	it('charges every generated command against the geometry budget', async () => {
		const bytes = await fixture({
			pages: [
				{ id: '0', contents: `<Shapes>${shape('1', cell('Rounding', 0.2) + rectangle)}</Shapes>` },
			],
		});
		await expect(parseVsdx(bytes, { maxGeometryCommands: 9 })).rejects.toMatchObject({
			code: 'GEOMETRY_LIMIT',
		});
		await expect(parseVsdx(bytes, { maxGeometryCommands: 10 })).resolves.toHaveProperty(
			'format',
			'vsdx',
		);
	});
	it('leaves shape dimensions, rotation, flips and pin transforms unchanged', async () => {
		const cells =
			cell('Width', 4) +
			cell('Height', 2) +
			cell('PinX', 10) +
			cell('PinY', -3) +
			cell('LocPinX', 1) +
			cell('LocPinY', 0.25) +
			cell('Angle', 0.7) +
			cell('FlipX', 1);
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', cells + cell('Rounding', 0.3) + rectangle)}${shape('2', cells + rectangle)}</Shapes>`,
				},
			],
		});
		const [rounded, original] = (await parseVsdx(bytes)).pages[0]!.shapes;
		expect(rounded!.transform).toEqual(original!.transform);
		expect([rounded!.width, rounded!.height]).toEqual([original!.width, original!.height]);
	});
	it('inherits cached radius through line styles and masters while honoring local overrides', async () => {
		const styles = `<StyleSheets><StyleSheet ID="2">${cell('Rounding', 0.4)}</StyleSheet></StyleSheets>`;
		const bytes = await fixture({
			document: styles,
			masters: [{ id: '7', shapes: shape('10', rectangle, 'LineStyle="2"') }],
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', '', 'Master="7"')}${shape('2', cell('Rounding', 0.2), 'Master="7"')}${shape('3', cell('Rounding', 0), 'Master="7"')}</Shapes>`,
				},
			],
		});
		const shapes = (await parseVsdx(bytes)).pages[0]!.shapes;
		expect(shapes[0]!.geometry[0]!.path).toContain('A 0.4 0.4');
		expect(shapes[1]!.geometry[0]!.path).toContain('A 0.2 0.2');
		expect(shapes[2]!.geometry[0]!.path).not.toContain(' A ');
		await expect(parseVsdx(bytes, { maxGeometryCommands: 9 })).rejects.toMatchObject({
			code: 'GEOMETRY_LIMIT',
		});
	});
});
