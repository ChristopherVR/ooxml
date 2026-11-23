import { describe, expect, it } from 'vitest';
import { normalizedNurbs, type NurbsControl } from './nurbs';

type Point = readonly [number, number];

function parsePeriodicPath(path: string): { start: Point; vertices: Point[]; closed: boolean } {
	const move = /^M ([^ ]+) ([^ ]+) /.exec(path);
	expect(move).not.toBeNull();
	const start: Point = [Number(move![1]), Number(move![2])];
	const vertices: Point[] = Array.from(path.matchAll(/L ([^ ]+) ([^ ]+)/g), (m) => [
		Number(m[1]),
		Number(m[2]),
	]);
	return { start, vertices, closed: path.endsWith(' Z') };
}

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
	return { path, warnings, consumed };
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

/** Independent Cox-de Boor reference evaluator for periodic Visio compact knots. */
function periodicReference(
	controls: readonly NurbsControl[],
	compact: readonly number[],
	degree: number,
	fraction: number,
): Point {
	const n = controls.length;
	const base = compact[0]!;
	const period = compact.at(-1)! - base;
	const wrapped: NurbsControl[] = [
		...controls,
		...Array.from({ length: degree }, (_, i) => controls[i % n]!),
	];
	const knots = Array.from({ length: n + 2 * degree + 1 }, (_, i) => {
		const cycles = Math.floor(i / n);
		return cycles * period + (compact[i % n]! - base);
	});
	const uStart = knots[degree]!;
	const uEnd = knots[n + degree]!;
	const t = uStart + (uEnd - uStart) * (fraction >= 1 ? 0 : fraction);
	let basis: number[] = knots.slice(0, -1).map((k, i) => (k <= t && t < knots[i + 1]! ? 1 : 0));
	for (let d = 1; d <= degree; d++) {
		const next: number[] = [];
		for (let i = 0; i < basis.length - 1; i++) {
			const a = knots[i + d]! - knots[i]!;
			const b = knots[i + d + 1]! - knots[i + 1]!;
			next.push(
				(a ? ((t - knots[i]!) / a) * basis[i]! : 0) +
					(b ? ((knots[i + d + 1]! - t) / b) * basis[i + 1]! : 0),
			);
		}
		basis = next;
	}
	let x = 0,
		y = 0,
		w = 0;
	for (let i = 0; i < wrapped.length; i++) {
		const c = wrapped[i]!;
		const f = basis[i]! * c.weight;
		x += f * c.x;
		y += f * c.y;
		w += f;
	}
	return [x / w, y / w];
}

function assertMatchesOracle(
	controls: NurbsControl[],
	knots: number[],
	degree: number,
	samples = 200,
) {
	const result = sample(controls, knots, degree);
	expect(result.path).toBeDefined();
	const { start, vertices, closed } = parsePeriodicPath(result.path!);
	expect(closed).toBe(true);
	expect(vertices.at(-1)).toEqual(start);
	expect(result.consumed).toBe(vertices.length);
	const poly: Point[] = [start, ...vertices];
	for (let j = 0; j <= samples; j++) {
		const p = periodicReference(controls, knots, degree, j / samples);
		const error = Math.min(...poly.slice(1).map((pt, i) => distance(p, poly[i]!, pt)));
		expect(error, `degree ${degree}, fraction ${j / samples}`).toBeLessThanOrEqual(0.000105);
	}
	return { start, vertices };
}

describe('periodic Visio NURBS evaluation', () => {
	it('closes a degree-1 periodic triangle through all three vertices', () => {
		const controls: NurbsControl[] = [
			{ x: 0, y: 0, weight: 1 },
			{ x: 2, y: 0, weight: 1 },
			{ x: 1, y: 2, weight: 1 },
		];
		const { start, vertices } = assertMatchesOracle(controls, [0, 1, 2, 3], 1, 60);
		expect(start).toEqual([0, 0]);
		expect(vertices).toContainEqual([2, 0]);
		expect(vertices).toContainEqual([1, 2]);
		expect(vertices.at(-1)).toEqual([0, 0]);
	});

	it('evaluates a degree-2 periodic triangle tangent to edge midpoints with C1 seam continuity', () => {
		const controls: NurbsControl[] = [
			{ x: 0, y: 0, weight: 1 },
			{ x: 2, y: 0, weight: 1 },
			{ x: 1, y: 2, weight: 1 },
		];
		const { start, vertices } = assertMatchesOracle(controls, [0, 1, 2, 3], 2, 120);
		expect(start[0]).toBeCloseTo(1, 12);
		expect(start[1]).toBeCloseTo(0, 12);
		// Tangent entering the seam matches tangent leaving the seam.
		const first = vertices[0]!;
		const beforeEnd = vertices.at(-2)!;
		const outDx = first[0] - start[0],
			outDy = first[1] - start[1];
		const inDx = start[0] - beforeEnd[0],
			inDy = start[1] - beforeEnd[1];
		const cosAngle =
			(outDx * inDx + outDy * inDy) / (Math.hypot(outDx, outDy) * Math.hypot(inDx, inDy));
		expect(cosAngle).toBeGreaterThan(0.999);
	});

	it('represents an exact rational 360-degree unit circle as a periodic degree-2 NURBS', () => {
		const s3 = Math.sqrt(3);
		// Alternating tangent corners (w = 0.5) and on-circle points (w = 1) with double knots.
		const controls: NurbsControl[] = [
			{ x: 1, y: -s3, weight: 0.5 },
			{ x: 1, y: 0, weight: 1 },
			{ x: 1, y: s3, weight: 0.5 },
			{ x: -0.5, y: s3 / 2, weight: 1 },
			{ x: -2, y: 0, weight: 0.5 },
			{ x: -0.5, y: -s3 / 2, weight: 1 },
		];
		const { start, vertices } = assertMatchesOracle(controls, [0, 0, 1, 1, 2, 2, 3], 2, 180);
		for (const [x, y] of [start, ...vertices]) {
			expect(Math.abs(x * x + y * y - 1)).toBeLessThan(1e-12);
		}
	});

	it('evaluates a degree-3 periodic cubic B-spline on a square with 4-fold symmetry and C2 seam', () => {
		const controls: NurbsControl[] = [
			{ x: 0, y: 0, weight: 1 },
			{ x: 2, y: 0, weight: 1 },
			{ x: 2, y: 2, weight: 1 },
			{ x: 0, y: 2, weight: 1 },
		];
		const { start, vertices } = assertMatchesOracle(controls, [0, 1, 2, 3, 4], 3, 200);
		// At u = 3: (1/6)P0 + (4/6)P1 + (1/6)P2 = (5/3, 1/3).
		expect(start[0]).toBeCloseTo(5 / 3, 12);
		expect(start[1]).toBeCloseTo(1 / 3, 12);
		// Seam tangent continuity.
		const first = vertices[0]!;
		const beforeEnd = vertices.at(-2)!;
		const outDx = first[0] - start[0],
			outDy = first[1] - start[1];
		const inDx = start[0] - beforeEnd[0],
			inDy = start[1] - beforeEnd[1];
		const cosAngle =
			(outDx * inDx + outDy * inDy) / (Math.hypot(outDx, outDy) * Math.hypot(inDx, inDy));
		expect(cosAngle).toBeGreaterThan(0.999);
	});

	it('handles partial initial knot multiplicity (initial = 2, degree = 3)', () => {
		const controls: NurbsControl[] = [
			{ x: 0, y: 0, weight: 1 },
			{ x: 2, y: 0, weight: 1.5 },
			{ x: 2, y: 2, weight: 0.75 },
			{ x: 0, y: 2, weight: 2 },
		];
		assertMatchesOracle(controls, [0, 0, 1, 1, 2], 3, 160);
	});

	it('respects rational weights and is invariant to uniform weight scaling', () => {
		const base: NurbsControl[] = [
			{ x: 0, y: 0, weight: 1 },
			{ x: 3, y: 0, weight: 4 },
			{ x: 3, y: 3, weight: 0.5 },
			{ x: 0, y: 3, weight: 2 },
		];
		const uniform = base.map((c) => ({ ...c, weight: 1 }));
		const scaled = base.map((c) => ({ ...c, weight: c.weight * 5 }));
		const knots = [0, 1, 2, 3, 4];
		const weighted = assertMatchesOracle(base, knots, 3, 160);
		const unweighted = assertMatchesOracle(uniform, knots, 3, 160);
		const rescaled = assertMatchesOracle(scaled, knots, 3, 160);
		// High weight on (3, 0) pulls start point closer to (3, 0) than unweighted curve.
		expect(Math.hypot(weighted.start[0] - 3, weighted.start[1])).toBeLessThan(
			Math.hypot(unweighted.start[0] - 3, unweighted.start[1]),
		);
		expect(rescaled.start[0]).toBeCloseTo(weighted.start[0], 12);
		expect(rescaled.start[1]).toBeCloseTo(weighted.start[1], 12);
	});

	it('is invariant under affine knot normalization (shifts and positive scales)', () => {
		const controls: NurbsControl[] = [
			{ x: 0, y: 0, weight: 1 },
			{ x: 2, y: 0.5, weight: 2 },
			{ x: 1.5, y: 2.5, weight: 0.8 },
			{ x: -0.5, y: 1.5, weight: 1.3 },
		];
		const baseKnots = [0, 0.15, 0.45, 0.8, 1];
		const shifted = baseKnots.map((k) => k + 1e7);
		const negative = baseKnots.map((k) => k * 7.5 - 500);
		const tiny = baseKnots.map((k) => k * 1e-4);

		const ref = parsePeriodicPath(sample(controls, baseKnots, 3).path!);
		for (const variant of [shifted, negative, tiny]) {
			const res = assertMatchesOracle(controls, variant, 3, 100);
			// Input quantization at 1e7 is ~2e-9 in float64; verify agreement within 5e-9.
			expect(res.start[0]).toBeCloseTo(ref.start[0], 8);
			expect(res.start[1]).toBeCloseTo(ref.start[1], 8);
		}
		// Binary-exact knots shifted by 1e7 preserve full precision after subtracting base.
		const dyadic = [0, 0.25, 0.5, 0.75, 1];
		const dyadicRef = parsePeriodicPath(sample(controls, dyadic, 3).path!);
		const dyadicShifted = assertMatchesOracle(
			controls,
			dyadic.map((k) => k + 1e7),
			3,
			100,
		);
		expect(dyadicShifted.start[0]).toBeCloseTo(dyadicRef.start[0], 12);
		expect(dyadicShifted.start[1]).toBeCloseTo(dyadicRef.start[1], 12);
	});

	it('matches the independent Cox-de Boor oracle across seeded degrees 1 through 8', () => {
		let seed = 419283;
		const random = () => {
			seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
			return seed / 2 ** 32;
		};
		for (let degree = 1; degree <= 8; degree++) {
			const count = degree + 2 + (degree % 3);
			const controls: NurbsControl[] = Array.from({ length: count }, (_, i) => {
				const angle = (2 * Math.PI * i) / count;
				const radius = 0.8 + random() * 0.4;
				return {
					x: Math.cos(angle) * radius,
					y: Math.sin(angle) * radius,
					weight: 0.3 + random() * 2.5,
				};
			});
			const knots = [0];
			for (let i = 1; i <= count; i++) knots.push(knots[i - 1]! + 0.2 + random());
			assertMatchesOracle(controls, knots, degree, 150);
		}
	});

	it('still rejects invalid periodic knot vectors (repeated end, excessive internal multiplicity)', () => {
		const controls: NurbsControl[] = [
			{ x: 0, y: 0, weight: 1 },
			{ x: 1, y: 1, weight: 1 },
			{ x: 2, y: 0, weight: 1 },
			{ x: 1, y: -1, weight: 1 },
		];
		// Repeated last knot
		expect(sample(controls, [0, 0.25, 0.5, 1, 1], 2).path).toBeUndefined();
		// Internal knot multiplicity 3 > degree 2
		expect(sample(controls, [0, 0.5, 0.5, 0.5, 1], 2).path).toBeUndefined();
		// Descending knot
		expect(sample(controls, [0, 0.6, 0.4, 0.8, 1], 2).path).toBeUndefined();
	});

	it('does not regress clamped NURBS open path format and exact endpoints', () => {
		const controls: NurbsControl[] = [
			{ x: 0, y: 0, weight: 1 },
			{ x: 1, y: 1, weight: 1 },
			{ x: 2, y: 0, weight: 1 },
		];
		const result = sample(controls, [0, 0, 0, 1], 2);
		expect(result.path).toBeDefined();
		expect(result.path!.startsWith('L ')).toBe(true);
		expect(result.path!.endsWith(' Z')).toBe(false);
		expect(result.path!.endsWith('L 2 0')).toBe(true);
	});
});
