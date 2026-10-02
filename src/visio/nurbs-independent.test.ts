import { describe, expect, it } from 'vitest';
import { normalizedNurbs, type NurbsControl } from './nurbs.js';

type Point = readonly [number, number];
function sample(controls: NurbsControl[], knots: number[], degree: number) {
	let consumed = 0;
	const warnings: string[] = [];
	const path = normalizedNurbs(
		controls,
		knots,
		degree,
		(code) => warnings.push(code),
		() => consumed++,
	);
	const points: Point[] = path
		? Array.from(path.matchAll(/L ([^ ]+) ([^ ]+)/g), (m) => [Number(m[1]), Number(m[2])])
		: [];
	return { path, points, warnings, consumed };
}
function bernstein(controls: number[], t: number): number {
	const values = [...controls];
	for (let level = 1; level < values.length; level++)
		for (let i = 0; i < values.length - level; i++)
			values[i] = values[i]! * (1 - t) + values[i + 1]! * t;
	return values[0]!;
}
function distance(p: Point, a: Point, b: Point): number {
	const dx = b[0] - a[0],
		dy = b[1] - a[1],
		squared = dx * dx + dy * dy;
	const t = squared
		? Math.min(1, Math.max(0, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / squared))
		: 0;
	return Math.hypot(p[0] - a[0] - t * dx, p[1] - a[1] - t * dy);
}
/** Independent Cox basis recurrence: no knot insertion or homogeneous subdivision. */
function referencePoint(
	controls: NurbsControl[],
	compact: number[],
	degree: number,
	t: number,
): Point {
	if (t === compact.at(-1)) return [controls.at(-1)!.x, controls.at(-1)!.y];
	const knots: number[] = [...compact, ...Array<number>(degree).fill(compact.at(-1)!)];
	let basis: number[] = knots.slice(0, -1).map((k, i) => (k <= t && t < knots[i + 1]! ? 1 : 0));
	for (let d = 1; d <= degree; d++) {
		const next: number[] = [];
		for (let i = 0; i < basis.length - 1; i++) {
			const a = knots[i + d]! - knots[i]!,
				b = knots[i + d + 1]! - knots[i + 1]!;
			next.push(
				(a ? ((t - knots[i]!) / a) * basis[i]! : 0) +
					(b ? ((knots[i + d + 1]! - t) / b) * basis[i + 1]! : 0),
			);
		}
		basis = next;
	}
	let x = 0,
		y = 0,
		weight = 0;
	for (let i = 0; i < controls.length; i++) {
		const control = controls[i]!,
			factor = basis[i]! * control.weight;
		x += factor * control.x;
		y += factor * control.y;
		weight += factor;
	}
	return [x / weight, y / weight];
}
function compareReference(
	controls: NurbsControl[],
	knots: number[],
	degree: number,
	count: number,
) {
	const result = sample(controls, knots, degree);
	expect(result.path).toBeDefined();
	const points: Point[] = [[controls[0]!.x, controls[0]!.y], ...result.points];
	for (let j = 0; j <= count; j++) {
		const t = j / count,
			p = referencePoint(controls, knots, degree, t);
		const error = Math.min(...points.slice(1).map((point, i) => distance(p, points[i]!, point)));
		expect(error, `degree ${degree}, parameter ${t}`).toBeLessThanOrEqual(0.000102);
	}
}

describe('independent bounded NURBS reference oracles', () => {
	it('matches independent Bernstein ordinates for every degree from 1 through 25', () => {
		for (let degree = 1; degree <= 25; degree++) {
			const y = Array.from({ length: degree + 1 }, (_, i) => Math.sin(i) * 0.01);
			const result = sample(
				y.map((value, i) => ({ x: i / degree, y: value, weight: 1 })),
				[...Array<number>(degree + 1).fill(0), 1],
				degree,
			);
			expect(result.path).toBeDefined();
			// Uniform x control coordinates make the Bezier's x equal its parameter.
			for (const [x, actualY] of result.points)
				expect(Math.abs(actualY - bernstein(y, x))).toBeLessThan(1e-12);
			expect(result.points.at(-1)).toEqual([1, y.at(-1)]);
		}
	});

	it('matches the rational quarter-circle equation and exact cached endpoint', () => {
		const result = sample(
			[
				{ x: 1, y: 0, weight: 1 },
				{ x: 1, y: 1, weight: Math.SQRT1_2 },
				{ x: 0, y: 1, weight: 1 },
			],
			[0, 0, 0, 1],
			2,
		);
		expect(result.path).toBeDefined();
		for (const [x, y] of result.points) expect(Math.abs(x * x + y * y - 1)).toBeLessThan(1e-14);
		expect(result.points.at(-1)).toEqual([0, 1]);
	});

	it('bounds 48 seeded multi-span rational curves against an independent Cox basis', () => {
		let seed = 827361;
		const random = () => {
			seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
			return seed / 2 ** 32;
		};
		for (let test = 0; test < 48; test++) {
			const degree = 1 + (test % 8),
				count = degree + 2 + (test % 5),
				internal = count - degree - 1;
			const controls = Array.from({ length: count }, (_, i) => ({
				x: i / (count - 1),
				y: (random() - 0.5) * 0.1,
				weight: 0.2 + random() * 3,
			}));
			const knots = [
				...Array<number>(degree + 1).fill(0),
				...Array.from({ length: internal }, (_, i) => (i + 1) / (internal + 1)),
				1,
			];
			compareReference(controls, knots, degree, 200);
		}
	});

	it('preserves a degree-25 repeated C0 knot while inserting a second internal knot', () => {
		const controls = Array.from({ length: 52 }, (_, i) => ({
			x: i / 51,
			y: 0.01 * Math.sin(i),
			weight: 0.5 + (i % 5) / 3,
		}));
		const knots = [...Array<number>(26).fill(0), ...Array<number>(25).fill(0.25), 0.75, 1];
		compareReference(controls, knots, 25, 1000);
	});

	it('keeps simultaneous degree-25 and 256-control input subject to the work limit', () => {
		const controls = Array.from({ length: 256 }, (_, i) => ({ x: i / 255, y: 0, weight: 1 }));
		const knots = [...Array<number>(26).fill(-5), ...Array.from({ length: 231 }, (_, i) => i - 4)];
		const result = sample(controls, knots, 25);
		// The maxima are independent: large refinement can be explicitly omitted.
		if (!result.path) {
			expect(result.warnings).toContain('geometry-limit');
			expect(result.warnings).not.toContain('geometry-approximation');
		} else {
			expect(result.consumed).toBeGreaterThan(0);
			expect(result.consumed).toBeLessThanOrEqual(2048);
			expect(result.points.every((point) => point.every(Number.isFinite))).toBe(true);
			expect(result.points.at(-1)).toEqual([1, 0]);
		}
	});
});
