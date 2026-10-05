import { describe, expect, it } from 'vitest';
import { normalizedNurbs } from './nurbs.js';

// Independently generated Bernstein coefficients of k*product(t-j/16), j=0..16,
// scaled so y(1/64)=1. Fixed quarter/midpoint probes previously missed every loop.
const ALIAS = [
	0, 9.28059844954701, -44.189179606705224, 131.96610459207767, -299.9224747307209,
	554.991357023364, -863.8055237910094, 1151.019022475428, -1325.79936188791, 1325.79936188791,
	-1151.019022475428, 863.8055237910094, -554.991357023364, 299.9224747307209, -131.96610459207767,
	44.189179606705224, -9.28059844954701, 0,
];
function curve(
	ys: number[],
	knots: number[],
	degree: number,
	work: (n: number) => void = () => {},
) {
	const notes: string[] = [];
	const path = normalizedNurbs(
		ys.map((y, i) => ({ x: i / (ys.length - 1), y, weight: 1 })),
		knots,
		degree,
		(code) => notes.push(code),
		() => {},
		'test',
		work,
	);
	const points = [
		[0, ys[0]!],
		...Array.from(path?.matchAll(/L ([^ ]+) ([^ ]+)/g) ?? [], (m) => [Number(m[1]), Number(m[2])]),
	];
	return { path, points, notes };
}
function bernstein(ys: number[], t: number): number {
	const work = [...ys];
	for (let r = 1; r < work.length; r++)
		for (let i = 0; i < work.length - r; i++) work[i] = (1 - t) * work[i]! + t * work[i + 1]!;
	return work[0]!;
}
function nearest(points: number[][], x: number, y: number): number {
	let best = Infinity;
	for (let i = 1; i < points.length; i++) {
		const a = points[i - 1]!,
			b = points[i]!,
			dx = b[0]! - a[0]!,
			dy = b[1]! - a[1]!,
			sq = dx * dx + dy * dy;
		const t = sq ? Math.max(0, Math.min(1, ((x - a[0]!) * dx + (y - a[1]!) * dy) / sq)) : 0;
		best = Math.min(best, Math.hypot(x - a[0]! - t * dx, y - a[1]! - t * dy));
	}
	return best;
}
describe('control-hull bounded NURBS flattening', () => {
	it('preserves degree17 lobes invisible to fixed flatness probes', () => {
		const result = curve(ALIAS, [...Array<number>(18).fill(0), 1], 17);
		expect(result.path).toBeDefined();
		expect(Math.max(...result.points.map((p) => Math.abs(p[1]!)))).toBeGreaterThan(0.9);
		for (let i = 0; i <= 1024; i++)
			expect(nearest(result.points, i / 1024, bernstein(ALIAS, i / 1024))).toBeLessThanOrEqual(
				0.000101,
			);
	});
	it.each([1, 2, 3, 5, 9, 17, 25])(
		'matches independent polynomial evaluation for degree %i',
		(degree) => {
			const ys = Array.from({ length: degree + 1 }, (_, i) => Math.sin(i) * 0.2);
			const result = curve(ys, [...Array<number>(degree + 1).fill(0), 1], degree);
			expect(result.path).toBeDefined();
			for (let i = 0; i <= 128; i++)
				expect(nearest(result.points, i / 128, bernstein(ys, i / 128))).toBeLessThanOrEqual(
					0.000101,
				);
		},
	);
	it('propagates aggregate work failures instead of converting them into omission diagnostics', () => {
		let work = 0;
		expect(() =>
			curve(ALIAS, [...Array<number>(18).fill(0), 1], 17, (n) => {
				if ((work += n) > 100) throw new Error('work budget');
			}),
		).toThrow('work budget');
		expect(work).toBeGreaterThan(100);
	});
	it('retains a valid curve over a subnormal knot range without parameter midpoint underflow', () => {
		const result = curve([0, 1, 0], [0, 0, 0, Number.MIN_VALUE], 2);
		expect(result.path).toBeDefined();
		expect(result.points.every((p) => p.every(Number.isFinite))).toBe(true);
		expect(nearest(result.points, 0.5, 0.5)).toBeLessThanOrEqual(0.000101);
	});
});
