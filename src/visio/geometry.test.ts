import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index.js';
import { geometryPaths, shapeTransform } from './geometry.js';
import { readSheet, type Cells } from './sheet.js';
import { cell, row, section, xml } from './test-fixtures.js';

function geometry(rows: string, width = 4, height = 2) {
	const warnings: string[] = [];
	const sheet = readSheet(parseXml(xml('Shape', section('Geometry', rows))).documentElement);
	return {
		paths: geometryPaths(
			sheet,
			width,
			height,
			(code) => warnings.push(code),
			() => {},
		),
		warnings,
	};
}
describe('Visio geometry normalization', () => {
	it('normalizes relative quadratic and cubic Bezier control points', () => {
		const { paths } = geometry(
			row(1, 'RelMoveTo', cell('X', 0) + cell('Y', 0)) +
				row(2, 'RelQuadBezTo', cell('X', 0.5) + cell('Y', 0.5) + cell('A', 0.25) + cell('B', 1)) +
				row(
					3,
					'RelCubBezTo',
					cell('X', 1) +
						cell('Y', 1) +
						cell('A', 0.5) +
						cell('B', 0) +
						cell('C', 0.75) +
						cell('D', 0.5),
				),
		);
		expect(paths[0]!.path).toBe('M 0 0 Q 1 2 2 1 C 2 0 3 1 4 2');
	});
	it('converts circular arcs and zero-bow lines', () => {
		const { paths } = geometry(
			row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) +
				row(2, 'ArcTo', cell('X', 2) + cell('Y', 0) + cell('A', 1)) +
				row(3, 'ArcTo', cell('X', 3) + cell('Y', 0) + cell('A', 0)),
		);
		expect(paths[0]!.path).toBe('M 0 0 A 1 1 0 0 0 2 0 L 3 0');
	});
	it('converts ellipses into two arcs and closes them', () => {
		const { paths } = geometry(
			row(
				1,
				'Ellipse',
				cell('X', 2) + cell('Y', 1) + cell('A', 4) + cell('B', 1) + cell('C', 2) + cell('D', 2),
			),
		);
		expect(paths[0]!.path).toBe('M 4 1 A 2 1 0 0 1 0 1 A 2 1 0 0 1 4 1 Z');
	});
	it('converts three-point elliptical arcs and collinear arcs safely', () => {
		const arc =
			row(1, 'MoveTo', cell('X', -2) + cell('Y', 0)) +
			row(
				2,
				'EllipticalArcTo',
				cell('X', 2) + cell('Y', 0) + cell('A', 0) + cell('B', 1) + cell('C', 0) + cell('D', 2),
			);
		expect(geometry(arc).paths[0]!.path).toBe('M -2 0 A 2 1 0 0 0 2 0');
	});
	it('omits an entire section rather than joining across an unsupported curve', () => {
		const result = geometry(
			row(1, 'MoveTo', cell('X', 0)) +
				row(2, 'UnknownCurve', cell('X', 1)) +
				row(3, 'LineTo', cell('X', 2)),
		);
		expect(result.paths).toEqual([]);
		expect(result.warnings).toContain('unsupported-geometry');
	});
	it('honors geometry visibility, fill and stroke switches', () => {
		expect(geometry(cell('NoShow', 1) + row(1, 'MoveTo', '')).paths).toEqual([]);
		expect(
			geometry(cell('NoFill', 1) + cell('NoLine', 1) + row(1, 'MoveTo', '')).paths[0],
		).toMatchObject({ fill: false, stroke: false });
	});
	it('normalizes transforms without NaN from hostile cached values', () => {
		const cells: Cells = new Map([
			['PinX', { value: 'Infinity' }],
			['Angle', { value: 'NaN' }],
			['FlipX', { value: '1' }],
		]);
		expect(shapeTransform(cells, 4, 2).every(Number.isFinite)).toBe(true);
	});
	it('integrates literal polyline vertices and NURBS approximation', () => {
		const poly = geometry(
			row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) +
				row(
					2,
					'PolyLineTo',
					cell('X', 4) + cell('Y', 2) + '<Cell N="A" F="POLYLINE(0,0,0.5,0.5)"/>',
				),
		);
		expect(poly.paths[0]!.path).toBe('M 0 0 L 2 1 L 4 2');
		const nurbs = geometry(
			row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) +
				row(
					2,
					'NURBSTo',
					cell('X', 2) +
						cell('Y', 0) +
						cell('A', 0) +
						cell('B', 1) +
						cell('C', 0) +
						cell('D', 1) +
						'<Cell N="E" F="NURBS(1,2,1,1,1,1,0,1)"/>',
				),
		);
		expect(nurbs.paths[0]!.path).toContain('L 1 0.5');
		expect(nurbs.paths[0]!.path.endsWith('L 2 0')).toBe(true);
		expect(nurbs.warnings).toContain('geometry-approximation');
	});

	it.each([undefined, 'Width', '1e999'])(
		'does not synthesize NURBS first control from invalid predecessor %s',
		(value) => {
			const before = row(1, 'MoveTo', cell('Y', 0) + (value === undefined ? '' : cell('X', value)));
			const curve = row(
				2,
				'NURBSTo',
				cell('X', 2) +
					cell('Y', 0) +
					cell('A', 0) +
					cell('B', 1) +
					cell('C', 0) +
					cell('D', 1) +
					'<Cell N="E" F="NURBS(1,2,1,1,1,1,0,1)"/>',
			);
			expect(geometry(before + curve).paths).toEqual([]);
		},
	);
	it('bounds numeric cached tokens before number parsing', () => {
		const result = geometry(row(1, 'MoveTo', cell('X', '1'.repeat(60000) + 'x') + cell('Y', 0)));
		expect(result.paths[0]!.path).toBe('M 0 0');
		expect(result.warnings).toContain('missing-cached-value');
	});
	it('accepts the canonical schema PolylineTo row spelling', () => {
		const result = geometry(
			row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) +
				row(
					2,
					'PolylineTo',
					cell('X', 4) + cell('Y', 2) + '<Cell N="A" F="POLYLINE(0,0,0.5,0.5)"/>',
				),
		);
		expect(result.paths[0]!.path).toBe('M 0 0 L 2 1 L 4 2');
	});
});
