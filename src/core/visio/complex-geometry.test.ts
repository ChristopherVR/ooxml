import { describe, expect, it } from 'vitest';
import { geometryRow } from './complex-geometry.js';
import type { Cell, Row } from './sheet.js';

function run(
	type: string,
	formula: string,
	options: {
		values?: Record<string, string | number>;
		start?: [number, number];
		width?: number;
		height?: number;
		consume?: () => void;
	} = {},
) {
	const values = { X: 2, Y: 0, A: 0, B: 1, C: 0, D: 1, ...options.values };
	const cells = new Map<string, Cell>(
		Object.entries(values).map(([k, v]) => [k, { value: String(v) }]),
	);
	cells.set(type === 'PolyLineTo' ? 'A' : 'E', { formula });
	const row: Row = { type, index: '1', cells, deleted: false };
	const warnings: { code: string; message: string }[] = [];
	let consumed = 0;
	const path = geometryRow(
		row,
		...(options.start ?? [0, 0]),
		options.width ?? 4,
		options.height ?? 2,
		(code, message) => warnings.push({ code, message }),
		() => {
			consumed++;
			options.consume?.();
		},
	);
	return { path, warnings, consumed };
}
const vertices = (path: string) =>
	Array.from(path.matchAll(/L ([^ ]+) ([^ ]+)/g), (m) => [Number(m[1]), Number(m[2])]);
const quadratic = 'NURBS(1,2,1,1,1,1,0,1)';

describe('bounded literal Visio complex geometry', () => {
	it('expands the Microsoft POLYLINE rectangle example in local y-up inches', () => {
		// https://learn.microsoft.com/en-us/office/client-developer/visio/polyline-function
		const result = run('PolyLineTo', 'POLYLINE (0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0)', {
			values: { X: 0, Y: 0 },
		});
		expect(result.path).toBe('L 0 0 L 0 2 L 4 2 L 4 0 L 0 0');
		expect(result.consumed).toBe(5);
		expect(result.warnings).toEqual([]);
	});
	it('supports independent relative/local axes and appends the unscaled cached endpoint', () => {
		const result = run('PolyLineTo', 'polyline(0,1,+.5,1e-1,1.,2)', { values: { X: 7, Y: 8 } });
		expect(result.path).toBe('L 2 0.1 L 4 2 L 7 8');
		expect(result.consumed).toBe(3);
	});
	it('does not scale local coordinates or repeat a final cached point', () => {
		expect(run('PolyLineTo', 'POLYLINE(1,1,2,0)').path).toBe('L 2 0');
	});
	it.each([
		'POLYLINE(1,1,Width,0)',
		'POLYLINE(1,1,Sheet.3!PinX,0)',
		'POLYLINE(1,1,SIN(0),0)',
		'POLYLINE(1,1,1+2,0)',
		'POLYLINE(1,1,1 in,0)',
		'GUARD(POLYLINE(1,1,1,0))',
		'POLYLINE(1;1;1;0)',
		'POLYLINE(1,1,,0)',
		'POLYLINE(1,1,1,0,)',
		'POLYLINE(1,1,"1",0)',
		'POLYLINE(1,1,1,0); alert(1)',
		'POLYLINE(1,1,1,0',
	])('rejects expressions and malformed literal syntax: %s', (formula) => {
		const result = run('PolyLineTo', formula);
		expect(result.path).toBeUndefined();
		expect(result.warnings[0]!.code).toBe('unsupported-geometry-formula');
		expect(result.consumed).toBe(0);
	});
	it.each(['POLYLINE(2,1,1,1)', 'POLYLINE(1,1,1)', 'POLYLINE(1,1)', 'POLYLINE(1,1,1e999,0)'])(
		'rejects invalid numeric data: %s',
		(formula) => {
			expect(run('PolyLineTo', formula).path).toBeUndefined();
		},
	);
	it('rejects missing numeric cached endpoints and coordinates overflowing after scale', () => {
		expect(run('PolyLineTo', 'POLYLINE(1,1,1,1)', { values: { X: 'Width' } }).path).toBeUndefined();
		expect(run('PolyLineTo', 'POLYLINE(0,0,1e9,0)').path).toBeUndefined();
	});
	it('rejects malicious million-coordinate data before allocating argument arrays', () => {
		const result = run('NURBSTo', `NURBS(1,2,1,1,${'0,0,0,1,'.repeat(250_000)}0)`);
		expect(result.path).toBeUndefined();
		expect(result.consumed).toBe(0);
		expect(result.warnings[0]!.message).toContain('oversized');
	});
	it('rejects long near-numeric tokens without ambiguous digit-run backtracking', () => {
		const hostile = `${'1'.repeat(60_000)}x`;
		expect(run('PolyLineTo', `POLYLINE(1,1,${hostile},0)`).path).toBeUndefined();
		expect(run('NURBSTo', quadratic, { values: { X: hostile } }).path).toBeUndefined();
	});
	it('enforces a separate bounded point count on short numeric input', () => {
		const result = run('PolyLineTo', `POLYLINE(1,1,${Array(257).fill('0,0').join(',')})`);
		expect(result.path).toBeUndefined();
		expect(result.consumed).toBe(0);
	});
	it('samples a unit-weight quadratic with an explicit approximation diagnostic', () => {
		const result = run('NURBSTo', quadratic);
		const points = vertices(result.path!);
		expect(points).toContainEqual([1, 0.5]);
		expect(points.at(-1)).toEqual([2, 0]);
		expect(result.consumed).toBe(points.length);
		expect(result.warnings).toEqual([
			{ code: 'geometry-approximation', message: expect.stringContaining('control-hull flatness') },
		]);
		for (const [x, y] of points) expect(y).toBeCloseTo((x! * (2 - x!)) / 2, 10);
	});
	it('uses rational weights rather than pretending that the curve is polynomial', () => {
		const result = run('NURBSTo', 'NURBS(1,2,1,1,1,1,0,2)');
		expect(vertices(result.path!)).toContainEqual([1, 2 / 3]);
	});
	it('uses both cached endpoint weights in rational evaluation', () => {
		const result = run('NURBSTo', quadratic, { values: { B: 3, D: 2 } });
		expect(vertices(result.path!)).toContainEqual([8 / 7, 2 / 7]);
	});
	it('retains full cached endpoint precision and accepts the maximum supported degree', () => {
		const end = 1.123456789123;
		const points = Array.from({ length: 24 }, (_, i) => `${(i + 1) / 25},0,0,1`).join(',');
		const result = run('NURBSTo', `NURBS(1,25,1,1,${points})`, { values: { X: end } });
		expect(vertices(result.path!).at(-1)).toEqual([end, 0]);
		expect(result.consumed).toBe(4);
	});
	it('represents a rational quarter-circle and preserves the cached endpoint exactly', () => {
		const result = run('NURBSTo', `NURBS(1,2,1,1,1,1,0,${Math.SQRT1_2})`, {
			start: [1, 0],
			values: { X: 0, Y: 1 },
		});
		const points = vertices(result.path!);
		expect(points.at(-1)).toEqual([0, 1]);
		for (const [x, y] of points) expect(Math.hypot(x!, y!)).toBeCloseTo(1, 12);
	});
	it('scales only formula controls, leaving start and cached endpoint in local inches', () => {
		const result = run('NURBSTo', 'NURBS(1,2,0,0,.5,1,0,1)', { values: { X: 4, Y: 0 } });
		expect(vertices(result.path!)).toContainEqual([2, 1]);
		expect(result.path!.endsWith('L 4 0')).toBe(true);
	});
	it('accepts the documented DrawNURBS nonuniform knot sequence with implied end knots', () => {
		// Microsoft example: degree 2, compact knots (0,0,0,2,5,8), five controls.
		const result = run('NURBSTo', 'NURBS(8,2,1,1,1,1,0,1,2,1,0,1,3,0,2,1)', {
			values: { X: 4, Y: 0, A: 5 },
		});
		expect(result.path).toBeDefined();
		expect(result.path!.endsWith('L 4 0')).toBe(true);
		expect(result.consumed).toBeLessThanOrEqual(2048);
	});
	it('allows repeated interior knots up to the degree, retaining a sharp corner', () => {
		const result = run('NURBSTo', 'NURBS(1,2,1,1,1,0,0,1,1,1,0,1,2,1,.5,1)', {
			values: { X: 2, Y: 2, A: 0.5 },
		});
		expect(vertices(result.path!)).toContainEqual([1, 1]);
	});
	it.each([
		['NURBS(1,0,1,1,1,1,0,1)', {}],
		['NURBS(1,26,1,1,1,1,0,1)', {}],
		['NURBS(1,2.5,1,1,1,1,0,1)', {}],
		['NURBS(1,3,1,1,1,1,0,1)', {}],
		['NURBS(0,2,1,1,1,1,0,1)', {}],
		['NURBS(1,2,1,1,1,1,.5,1)', {}],
		['NURBS(1,2,1,1,1,1,0,0)', {}],
		['NURBS(1,2,1,1,1,1,0,-1)', {}],
		[quadratic, { B: 0 }],
		[quadratic, { D: 'NaN' }],
		[quadratic, { A: 1 }],
		[quadratic, { A: -1 }],
	] as const)('rejects invalid NURBS parameters: %s %j', (formula, values) => {
		const result = run('NURBSTo', formula, { values });
		expect(result.path).toBeUndefined();
		expect(result.consumed).toBe(0);
	});
	it('explicitly reports periodic splines as unsupported', () => {
		const result = run('NURBSTo', 'NURBS(1,2,1,1,1,1,.25,1)', { values: { A: 0.75 } });
		expect(result.path).toBeUndefined();
		expect(result.warnings[0]!.message).toContain('Periodic');
	});
	it('omits pathological curves when bounded sampling cannot reach its flatness target', () => {
		const result = run('NURBSTo', 'NURBS(1,2,1,1,1e9,1e9,0,1)', { values: { X: 1e9, Y: 0 } });
		expect(result.path).toBeUndefined();
		expect(result.consumed).toBeLessThanOrEqual(2048);
		expect(result.warnings.at(-1)!.code).toBe('geometry-limit');
	});
	it.each(['PolyLineTo', 'NURBSTo'])(
		'propagates the caller budget exception for every %s segment',
		(type) => {
			expect(() =>
				run(type, type === 'PolyLineTo' ? 'POLYLINE(1,1,1,1)' : quadratic, {
					consume: () => {
						throw new Error('budget');
					},
				}),
			).toThrow('budget');
		},
	);
});
