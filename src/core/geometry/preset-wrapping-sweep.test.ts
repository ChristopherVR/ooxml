import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { PRESET_SHAPE_GEOMETRY_TABLE } from './preset-shape-definitions-table';
import { evaluatePresetShape } from './preset-shape-evaluator';

/** Operand count of every DrawingML guide operator (ISO/IEC 29500-1 20.1.9.11). */
const ARITY: Readonly<Record<string, number>> = {
	val: 1,
	abs: 1,
	sqrt: 1,
	'+-': 3,
	'*/': 3,
	'+/': 3,
	'?:': 3,
	min: 2,
	max: 2,
	mod: 3,
	pin: 3,
	sin: 2,
	cos: 2,
	tan: 2,
	atan: 2,
	at2: 2,
	cat2: 3,
	sat2: 3,
};

/** End point of the first elliptical-arc (`A`) command in an SVG path. */
function firstArcEnd(d: string): [number, number] {
	const arc = /A\s+(-?[\d.]+)\s+(-?[\d.]+)\s+\S+\s+\S+\s+\S+\s+(-?[\d.]+)\s+(-?[\d.]+)/u.exec(d);
	expect(arc, d).not.toBeNull();
	return [Number(arc![3]), Number(arc![4])];
}

describe('preset sweeps that cross 0 degrees (issue #232)', () => {
	it('declares no guide with more operands than its operator takes', () => {
		const offenders = Object.values(PRESET_SHAPE_GEOMETRY_TABLE).flatMap((def) =>
			(def?.gdLst ?? [])
				.filter((guide) => {
					const [op, ...args] = guide.formula.trim().split(/\s+/u);
					return op !== undefined && ARITY[op] !== undefined && args.length > ARITY[op]!;
				})
				.map((guide) => `${def!.name}.${guide.name}: ${guide.formula}`),
		);
		expect(offenders).toEqual([]);
	});

	it('writes every guide literal in the geometry sources with its operator arity', () => {
		// Connection-site guides live outside the table, as `gd('name', 'formula')` literals.
		const dir = fileURLToPath(new URL('.', import.meta.url));
		const offenders = readdirSync(dir)
			.filter((file) => file.endsWith('.ts') && !file.endsWith('.test.ts'))
			.flatMap((file) =>
				[...readFileSync(join(dir, file), 'utf8').matchAll(/gd\('([^']+)',\s*'([^']+)'\)/gu)]
					.filter((match) => {
						const [op, ...args] = match[2]!.trim().split(/\s+/u);
						return op !== undefined && ARITY[op] !== undefined && args.length !== ARITY[op];
					})
					.map((match) => `${file} ${match[1]}: ${match[2]}`),
			);
		expect(offenders).toEqual([]);
	});

	// 314.5 degrees round to 8.3 degrees: the sweep wraps through 3 o'clock.
	const start = 18_872_393;
	const end = 498_077;
	const endRadians = (end / 60_000) * (Math.PI / 180);

	it.each(['arc', 'pie', 'chord', 'blockArc'])(
		'%s sweeps to where adj2 points when the sweep wraps',
		(name) => {
			const result = evaluatePresetShape(name, 200, 200, { adj1: start, adj2: end })!;
			const [x, y] = firstArcEnd(result.paths[0]!.d);
			expect(x).toBeCloseTo(100 + 100 * Math.cos(endRadians), 1);
			expect(y).toBeCloseTo(100 + 100 * Math.sin(endRadians), 1);
		},
	);
});
