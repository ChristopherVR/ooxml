/**
 * Rasterising for the PDF export's soft effects: paths flattened to polylines, filled with
 * anti-aliased scanlines (non-zero or even-odd) or stroked by distance to their lines, and blurred
 * with a Gaussian made of three box passes. Coverage is one float per pixel, row 0 at the top.
 * DOM-free.
 */
import type { PathSegment } from './pdf-path';

/** Sub-scanlines per pixel row for fills. */
const SUB = 4;

export type Point = readonly [number, number];
export interface Polyline {
	points: Point[];
	closed: boolean;
}

export const IDENTITY = [1, 0, 0, 1, 0, 0] as const;

export function transformed(m: readonly number[], x: number, y: number): Point {
	return [m[0]! * x + m[2]! * y + m[4]!, m[1]! * x + m[3]! * y + m[5]!];
}

/** The part's subpaths as polylines in silhouette units, curves cut finer than `step`. */
export function flattenSegments(
	segments: readonly PathSegment[],
	matrix: readonly number[] = IDENTITY,
	step = 1,
): Polyline[] {
	const out: Polyline[] = [];
	let current: Polyline | undefined;
	let x = 0;
	let y = 0;
	for (const segment of segments) {
		if (segment.t === 'M') {
			current = { points: [transformed(matrix, segment.x, segment.y)], closed: false };
			out.push(current);
			[x, y] = [segment.x, segment.y];
		} else if (segment.t === 'Z') {
			if (current) current.closed = true;
		} else {
			if (!current) {
				current = { points: [transformed(matrix, x, y)], closed: false };
				out.push(current);
			}
			if (segment.t === 'L') current.points.push(transformed(matrix, segment.x, segment.y));
			else {
				const p0 = transformed(matrix, x, y);
				const p1 = transformed(matrix, segment.x1, segment.y1);
				const p2 = transformed(matrix, segment.x2, segment.y2);
				const p3 = transformed(matrix, segment.x, segment.y);
				const length =
					Math.hypot(p1[0] - p0[0], p1[1] - p0[1]) +
					Math.hypot(p2[0] - p1[0], p2[1] - p1[1]) +
					Math.hypot(p3[0] - p2[0], p3[1] - p2[1]);
				const steps = Math.min(64, Math.max(2, Math.ceil(length / Math.max(step, 1e-9))));
				for (let index = 1; index <= steps; index++) {
					const t = index / steps;
					const u = 1 - t;
					const a = u * u * u,
						b = 3 * u * u * t,
						c = 3 * u * t * t,
						d = t * t * t;
					current.points.push([
						a * p0[0] + b * p1[0] + c * p2[0] + d * p3[0],
						a * p0[1] + b * p1[1] + c * p2[1] + d * p3[1],
					]);
				}
			}
			[x, y] = [segment.x, segment.y];
		}
	}
	return out;
}

/** Widths of three box blurs whose sum is a Gaussian of deviation `sigma` (in pixels). */
function boxWidths(sigma: number): number[] {
	const ideal = Math.sqrt((12 * sigma * sigma) / 3 + 1);
	let lower = Math.floor(ideal);
	if (lower % 2 === 0) lower--;
	const upper = lower + 2;
	const m = Math.round(
		(12 * sigma * sigma - 3 * lower * lower - 12 * lower - 9) / (-4 * lower - 4),
	);
	return [0, 1, 2].map((index) => (index < m ? lower : upper));
}

function boxPass(
	source: Float32Array,
	target: Float32Array,
	width: number,
	height: number,
	radius: number,
	horizontal: boolean,
): void {
	const lines = horizontal ? height : width;
	const length = horizontal ? width : height;
	const stride = horizontal ? 1 : width;
	const scale = 1 / (2 * radius + 1);
	for (let line = 0; line < lines; line++) {
		const base = horizontal ? line * width : line;
		let sum = 0;
		for (let index = 0; index <= Math.min(radius, length - 1); index++)
			sum += source[base + index * stride]!;
		for (let index = 0; index < length; index++) {
			target[base + index * stride] = sum * scale;
			const enter = index + radius + 1;
			const leave = index - radius;
			if (enter < length) sum += source[base + enter * stride]!;
			if (leave >= 0) sum -= source[base + leave * stride]!;
		}
	}
}

export function gaussianBlur(
	data: Float32Array,
	width: number,
	height: number,
	sigma: number,
): void {
	if (!(sigma > 0.2)) return;
	const scratch = new Float32Array(data.length);
	for (const box of boxWidths(sigma)) {
		const radius = (box - 1) / 2;
		if (radius < 1) continue;
		boxPass(data, scratch, width, height, radius, true);
		boxPass(scratch, data, width, height, radius, false);
	}
}

/** Adds the anti-aliased coverage of every filled part (pixel units) to `coverage`. */
export function fillCoverage(
	coverage: Float32Array,
	width: number,
	height: number,
	polygons: readonly Point[][],
	evenOdd: boolean,
): void {
	const rows: { x0: number; y0: number; x1: number; y1: number }[][] = Array.from(
		{ length: height },
		() => [],
	);
	for (const polygon of polygons)
		for (let index = 0; index < polygon.length; index++) {
			const [x0, y0] = polygon[index]!;
			const [x1, y1] = polygon[(index + 1) % polygon.length]!;
			if (y0 === y1) continue;
			const top = Math.max(0, Math.floor(Math.min(y0, y1)));
			const bottom = Math.min(height - 1, Math.floor(Math.max(y0, y1)));
			for (let row = top; row <= bottom; row++) rows[row]!.push({ x0, y0, x1, y1 });
		}
	const crossings: { x: number; winding: number }[] = [];
	const row = new Float32Array(width);
	for (let line = 0; line < height; line++) {
		const edges = rows[line]!;
		if (!edges.length) continue;
		row.fill(0);
		for (let sub = 0; sub < SUB; sub++) {
			const sy = line + (sub + 0.5) / SUB;
			crossings.length = 0;
			for (const e of edges) {
				if (sy < Math.min(e.y0, e.y1) || sy >= Math.max(e.y0, e.y1)) continue;
				crossings.push({
					x: e.x0 + ((sy - e.y0) * (e.x1 - e.x0)) / (e.y1 - e.y0),
					winding: e.y1 > e.y0 ? 1 : -1,
				});
			}
			crossings.sort((a, b) => a.x - b.x);
			let winding = 0;
			for (let index = 0; index < crossings.length - 1; index++) {
				winding += evenOdd ? 1 : crossings[index]!.winding;
				const inside = evenOdd ? winding % 2 === 1 : winding !== 0;
				if (inside) addSpan(row, crossings[index]!.x, crossings[index + 1]!.x, 1 / SUB);
			}
		}
		const offset = line * width;
		for (let x = 0; x < width; x++) coverage[offset + x] = coverage[offset + x]! + row[x]!;
	}
}

function addSpan(row: Float32Array, from: number, to: number, weight: number): void {
	const a = Math.max(0, Math.min(row.length, from));
	const b = Math.max(0, Math.min(row.length, to));
	if (b <= a) return;
	const first = Math.floor(a);
	const last = Math.floor(b);
	if (first === last) {
		row[first] = row[first]! + (b - a) * weight;
		return;
	}
	row[first] = row[first]! + (first + 1 - a) * weight;
	for (let x = first + 1; x < last; x++) row[x] = row[x]! + weight;
	if (last < row.length) row[last] = row[last]! + (b - last) * weight;
}

/** The coverage of a stroke `half` pixels either side of each line, kept with `max`. */
export function strokeCoverage(
	coverage: Float32Array,
	width: number,
	height: number,
	line: Polyline,
	half: number,
): void {
	const points = line.closed ? [...line.points, line.points[0]!] : line.points;
	for (let index = 0; index + 1 < points.length; index++) {
		const [ax, ay] = points[index]!;
		const [bx, by] = points[index + 1]!;
		const dx = bx - ax;
		const dy = by - ay;
		const lengthSquared = dx * dx + dy * dy;
		const reach = half + 1;
		const x0 = Math.max(0, Math.floor(Math.min(ax, bx) - reach));
		const x1 = Math.min(width - 1, Math.ceil(Math.max(ax, bx) + reach));
		const y0 = Math.max(0, Math.floor(Math.min(ay, by) - reach));
		const y1 = Math.min(height - 1, Math.ceil(Math.max(ay, by) + reach));
		for (let y = y0; y <= y1; y++)
			for (let x = x0; x <= x1; x++) {
				const px = x + 0.5 - ax;
				const py = y + 0.5 - ay;
				const t = lengthSquared ? Math.max(0, Math.min(1, (px * dx + py * dy) / lengthSquared)) : 0;
				const distance = Math.hypot(px - t * dx, py - t * dy);
				const value = Math.max(0, Math.min(1, half + 0.5 - distance));
				const at = y * width + x;
				if (value > coverage[at]!) coverage[at] = value;
			}
	}
}
