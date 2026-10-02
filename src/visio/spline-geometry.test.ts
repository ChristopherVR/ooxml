import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index.js';
import { geometryPaths } from './geometry.js';
import { readSheet, mergeSheets } from './sheet.js';
import { parseVsdx } from './index.js';
import { cell, row, section, xml, fixture, shape } from './test-fixtures.js';

const cells = (v: Record<string, string | number>) =>
	Object.entries(v)
		.map(([k, v]) => cell(k, v))
		.join('');
const move = row(1, 'MoveTo', cells({ X: 0, Y: 0 }));
const start = (v: Record<string, string | number> = {}) =>
	row(2, 'SplineStart', cells({ X: 1, Y: 1, A: 0, B: 0, C: 1, D: 2, ...v }));
const knot = (v: Record<string, string | number> = {}, index = 3) =>
	row(index, 'SplineKnot', cells({ X: 2, Y: 0, A: 0, ...v }));
const quadratic = move + start() + knot();
function run(rows: string, consumeExtra = () => {}) {
	const warnings: { code: string; message: string }[] = [];
	let consumed = 0;
	const sheet = readSheet(parseXml(xml('Shape', section('Geometry', rows))).documentElement);
	const paths = geometryPaths(
		sheet,
		4,
		2,
		(code, message) => warnings.push({ code, message }),
		() => {
			consumed++;
			consumeExtra();
		},
	);
	return { paths, warnings, consumed };
}
const vertices = (path: string) =>
	Array.from(path.matchAll(/L ([^ ]+) ([^ ]+)/g), (m) => [Number(m[1]), Number(m[2])]);
describe('cached contiguous Visio spline sequences', () => {
	it('approximates the unit-weight quadratic with exact endpoints and midpoint', () => {
		const result = run(quadratic),
			points = vertices(result.paths[0]!.path);
		expect(result.paths[0]!.path).toMatch(/^M 0 0 L /);
		expect(points).toContainEqual([1, 0.5]);
		expect(points.at(-1)).toEqual([2, 0]);
		expect(result.consumed).toBe(3 + points.length);
		expect(result.warnings).toEqual([
			{
				code: 'geometry-approximation',
				message: expect.stringContaining('SplineStart/SplineKnot'),
			},
		]);
	});

	it('bounds analytic quadratic chord error at dense interior samples', () => {
		const points = [[0, 0], ...vertices(run(quadratic).paths[0]!.path)];
		for (let i = 1; i < points.length; i++) {
			const a = points[i - 1]!,
				b = points[i]!;
			const dx = b[0]! - a[0]!,
				dy = b[1]! - a[1]!;
			for (let j = 1; j < 10; j++) {
				const t = (a[0]! + ((b[0]! - a[0]!) * j) / 10) / 2,
					x = 2 * t,
					y = 2 * t * (1 - t);
				const u = Math.max(
					0,
					Math.min(1, ((x - a[0]!) * dx + (y - a[1]!) * dy) / (dx * dx + dy * dy)),
				);
				expect(Math.hypot(x - a[0]! - u * dx, y - a[1]! - u * dy)).toBeLessThanOrEqual(0.000101);
			}
		}
	});
	it('matches analytic cubic vertices with linear x and bounds sampled chord error', () => {
		const points = [
			[0, 0],
			...vertices(
				run(
					move +
						start({ X: 1 / 3, Y: 1, D: 3 }) +
						knot({ X: 2 / 3, Y: 1 }, 3) +
						knot({ X: 1, Y: 0 }, 4),
				).paths[0]!.path,
			),
		];
		for (const [x, y] of points) expect(y).toBeCloseTo(3 * x! * (1 - x!), 12);
		for (let i = 1; i < points.length; i++) {
			const a = points[i - 1]!,
				b = points[i]!,
				x = (a[0]! + b[0]!) / 2,
				y = 3 * x * (1 - x);
			const dx = b[0]! - a[0]!,
				dy = b[1]! - a[1]!,
				u = ((x - a[0]!) * dx + (y - a[1]!) * dy) / (dx * dx + dy * dy);
			expect(Math.hypot(x - a[0]! - u * dx, y - a[1]! - u * dy)).toBeLessThanOrEqual(0.000101);
		}
	});
	it('accepts exactly 256 controls and the documented maximum degree 25', () => {
		const max = run(
			move +
				start({ D: 1, C: 255 }) +
				Array.from({ length: 254 }, (_, i) => knot({ X: i + 2, Y: 0, A: i + 1 }, i + 3)).join(''),
		);
		expect(vertices(max.paths[0]!.path).at(-1)).toEqual([255, 0]);
		const high = run(
			move +
				start({ D: 25, X: 0, Y: 0 }) +
				Array.from({ length: 24 }, (_, i) => knot({ X: 0, Y: 0, A: 0 }, i + 3)).join(''),
		);
		expect(vertices(high.paths[0]!.path)).toEqual([
			[0, 0],
			[0, 0],
			[0, 0],
			[0, 0],
		]);
	});
	it('rejects numeric curves that cannot fit the bounded approximation work', () => {
		const result = run(move + start({ X: 1e9, Y: 1e9 }) + knot({ X: 1e9, Y: 0 }));
		expect(result.paths).toEqual([]);
		expect(result.warnings).toContainEqual(expect.objectContaining({ code: 'geometry-limit' }));
		expect(result.consumed).toBeLessThanOrEqual(2051);
	});
	it('uses cached cubic controls, rather than drawing through their coordinates', () => {
		const result = run(
			move + start({ X: 0, Y: 1, D: 3 }) + knot({ X: 1, Y: 1 }, 3) + knot({ X: 1, Y: 0 }, 4),
		);
		const points = vertices(result.paths[0]!.path);
		expect(points).toContainEqual([0.5, 0.75]);
		expect(points.at(-1)).toEqual([1, 0]);
		expect(points).not.toContainEqual([0, 1]);
	});
	it('handles a degree-one multi-span polyline and a non-unit parameter range', () => {
		const result = run(move + start({ D: 1, B: 3, A: 3, C: 7 }) + knot({ A: 5 }));
		const points = vertices(result.paths[0]!.path);
		expect(points).toContainEqual([1, 1]);
		expect(points.at(-1)).toEqual([2, 0]);
	});
	it('keeps local coordinates independent of shape width and height', () => {
		const result = run(
			row(1, 'RelMoveTo', cells({ X: 0.5, Y: 0.5 })) + start({ X: 3, Y: 2 }) + knot({ X: 4, Y: 1 }),
		);
		expect(result.paths[0]!.path).toMatch(/^M 2 1 /);
		expect(vertices(result.paths[0]!.path)).toContainEqual([3, 1.5]);
	});
	it('uses the last spline control point as the next arc endpoint origin', () => {
		const result = run(quadratic + row(4, 'ArcTo', cells({ X: 4, Y: 0, A: 1 })));
		expect(result.paths[0]!.path).toMatch(/L 2 0 A 1 1 0 0 0 4 0$/);
	});
	it('supports successive spline sequences without leaking control points', () => {
		const second = start({ X: 3, Y: 1 }).replace('IX="2"', 'IX="4"') + knot({ X: 4 }, 5);
		const result = run(quadratic + second);
		expect(vertices(result.paths[0]!.path)).toContainEqual([3, 0.5]);
		expect(vertices(result.paths[0]!.path).at(-1)).toEqual([4, 0]);
		expect(result.warnings.filter((d) => d.code === 'geometry-approximation')).toHaveLength(2);
	});
	it.each([0, 26, 1.5, -1, 'Width', '1e999'])(
		'rejects unusable degree %s without partial section output',
		(D) => {
			const result = run(move + start({ D }) + knot() + row(4, 'LineTo', cells({ X: 5, Y: 5 })));
			expect(result.paths).toEqual([]);
			expect(result.warnings.some((d) => d.code === 'invalid-geometry')).toBe(true);
		},
	);
	it.each([
		start() + knot(),
		move + start(),
		move + knot(),
		move + start() + row(3, 'LineTo', cells({ X: 2, Y: 0 })) + knot({}, 4),
		row(1, 'Ellipse', cells({ X: 0, Y: 0, A: 1, B: 0, C: 0, D: 1 })) + start() + knot(),
	])('rejects malformed ordering and refuses to bridge omitted content', (rows) => {
		expect(run(rows).paths).toEqual([]);
	});

	it.each([undefined, 'Width', '1e999'])(
		'rejects unusable preceding endpoint %s instead of inventing origin',
		(value) => {
			const before = row(
				1,
				'MoveTo',
				value === undefined ? cells({ Y: 0 }) : cells({ X: value, Y: 0 }),
			);
			const result = run(before + start() + knot());
			expect(result.paths).toEqual([]);
			expect(result.warnings.some((d) => d.code === 'invalid-spline')).toBe(true);
		},
	);
	it('enforces aggregate weighted curve work separately from output commands', async () => {
		const bytes = await fixture({
			pages: [
				{ id: '0', contents: `<Shapes>${shape('1', section('Geometry', quadratic))}</Shapes>` },
			],
		});
		await expect(parseVsdx(bytes, { maxCurveWork: 1 })).rejects.toMatchObject({
			code: 'CURVE_WORK_LIMIT',
		});
		await expect(parseVsdx(bytes, { maxCurveWork: 0 })).rejects.toMatchObject({
			code: 'INVALID_LIMIT',
		});
		await expect(parseVsdx(bytes, { maxCurveWork: 200000 })).resolves.toMatchObject({
			format: 'vsdx',
		});
	});
	it('does not infer absent cached numbers from formula text or default zero', () => {
		for (const field of ['X', 'Y', 'A', 'B', 'C', 'D']) {
			// Replace directly on the start row, so MoveTo may keep its valid saved endpoint.
			const rawStart = start().replace(
				new RegExp(`<Cell N="${field}"[^>]*/>`),
				`<Cell N="${field}" F="GUARD(1)"/>`,
			);
			expect(run(move + rawStart + knot()).paths).toEqual([]);
		}
	});
	it('rejects descending, repeated-end, empty and unsupported nonclamped knot vectors', () => {
		for (const rows of [
			move + start({ B: 1, A: 0 }) + knot(),
			move + start({ C: 0 }) + knot(),
			move + start() + knot({ A: 1 }),
			move + start({ A: 0.1 }) + knot({ A: 0.2 }),
		])
			expect(run(rows).paths).toEqual([]);
	});
	it('bounds cached tokens and coordinates before curve evaluation', () => {
		for (const value of ['1e999', '9'.repeat(10000), '1+2', 'NaN'])
			expect(run(move + start({ X: value }) + knot()).paths).toEqual([]);
	});
	it('rejects over 256 controls while charging all source rows', () => {
		const result = run(
			move +
				start({ D: 1 }) +
				Array.from({ length: 255 }, (_, i) => knot({ A: i / 256 }, i + 3)).join(''),
		);
		expect(result.paths).toEqual([]);
		expect(result.consumed).toBe(257);
		expect(result.warnings).toContainEqual(expect.objectContaining({ code: 'geometry-limit' }));
	});
	it('propagates the caller work budget during source rows and generated segments', () => {
		let count = 0;
		expect(() =>
			run(quadratic, () => {
				if (++count > 10) throw new Error('budget');
			}),
		).toThrow('budget');
		expect(count).toBe(11);
	});
	it('honors inherited/deleted rows before grouping a spline sequence', () => {
		const base = readSheet(parseXml(xml('Shape', section('Geometry', quadratic))).documentElement);
		const local = readSheet(
			parseXml(xml('Shape', section('Geometry', '<Row IX="3"><Cell N="Y" V="2"/></Row>')))
				.documentElement,
		);
		const geometry = geometryPaths(
			mergeSheets(base, local),
			1,
			1,
			() => {},
			() => {},
		);
		expect(vertices(geometry[0]!.path).at(-1)).toEqual([2, 2]);
		const deleted = readSheet(
			parseXml(xml('Shape', section('Geometry', '<Row IX="3" Del="1"/>'))).documentElement,
		);
		expect(
			geometryPaths(
				mergeSheets(base, deleted),
				1,
				1,
				() => {},
				() => {},
			),
		).toEqual([]);
	});
	it('integrates through the bounded package parser and records approximation evidence', async () => {
		const model = await parseVsdx(
			await fixture({
				pages: [
					{ id: '0', contents: `<Shapes>${shape('1', section('Geometry', quadratic))}</Shapes>` },
				],
			}),
		);
		expect(vertices(model.pages[0]!.shapes[0]!.geometry[0]!.path)).toContainEqual([1, 0.5]);
		expect(model.diagnostics.some((d) => d.code === 'geometry-approximation')).toBe(true);
	});
});
