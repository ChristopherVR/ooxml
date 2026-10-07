import { flattenSvgPath } from '../geometry/svg-path-flatten';
import { cross2, ensureCCW, signedArea, vec2Eq } from '../geometry/shape-boolean-types';
import { lineIntersection } from '../geometry/shape-boolean-clipping';
import { segmentIntersection } from '../geometry/shape-boolean-union';
import type { VisioFillGradient, VisioGeometry, VisioGradientPaint } from './model';

/** Native direction 13 follows the outline, unlike Visio's rectangular SVG fallback.
 * Reuse the evaluated geometry, shared polygon primitives and existing gradient paints.
 * Curves other than a canonical bounds-filling ellipse, holes and open outlines remain unsupported.
 */
export function pathFillGradient(
	geometry: readonly VisioGeometry[],
	width: number,
	height: number,
	stops: VisioGradientPaint['stops'],
): VisioFillGradient | undefined {
	const filled = geometry.filter((part) => part.fill);
	if (filled.length !== 1) return undefined;
	const path = filled[0]!.path;
	if (path.length > 30_000) return undefined;
	const commands = path.match(/[A-Za-z]/g)?.join('');
	if (commands === 'MAAZ') {
		// Canonical axis-aligned Ellipse output from geometry.ts; never flatten its curve.
		const values = path.match(/-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/g)?.map(Number);
		const expected = [
			width,
			height / 2,
			width / 2,
			height / 2,
			0,
			0,
			1,
			0,
			height / 2,
			width / 2,
			height / 2,
			0,
			0,
			1,
			width,
			height / 2,
		];
		if (
			!values ||
			values.length !== expected.length ||
			values.some((value, i) => Math.abs(value - expected[i]!) > 1e-9)
		)
			return undefined;
		return { type: 'radial', center: [0.5, 0.5], radius: 0.5, stops };
	}
	// geometry.ts emits absolute M/L paths, optionally closed by Z or a repeated endpoint.
	if (!commands || !/^ML{2,256}Z?$/.test(commands)) return undefined;
	const loops = flattenSvgPath(path);
	if (loops.length !== 1) return undefined;
	const points = loops[0]!.map((point) => ({ x: point.x / width, y: point.y / height }));
	if (!vec2Eq(points[0]!, points.at(-1)!)) return undefined;
	points.pop();
	// The normalized paint currently requires an outline filling the shape bounds.
	if (
		Math.min(...points.map((p) => p.x)) !== 0 ||
		Math.max(...points.map((p) => p.x)) !== 1 ||
		Math.min(...points.map((p) => p.y)) !== 0 ||
		Math.max(...points.map((p) => p.y)) !== 1
	)
		return undefined;
	if (
		points.length < 3 ||
		points.length > 256 ||
		Math.abs(signedArea(points)) < 1e-9 ||
		points.some(
			({ x, y }) => !Number.isFinite(x) || !Number.isFinite(y) || x < 0 || x > 1 || y < 0 || y > 1,
		)
	)
		return undefined;
	// Reject self crossings, repeated vertices, touches and overlapping collinear edges.
	for (let i = 0; i < points.length; i++) {
		const a = points[i]!,
			b = points[(i + 1) % points.length]!;
		for (let j = i + 1; j < points.length; j++) {
			const c = points[j]!,
				d = points[(j + 1) % points.length]!;
			if (vec2Eq(a, c)) return undefined;
			if (j === i + 1 || (i === 0 && j === points.length - 1)) continue;
			const touches = (p: typeof a, start: typeof a, end: typeof a) =>
				Math.abs(cross2(start, end, p)) < 1e-9 &&
				p.x >= Math.min(start.x, end.x) &&
				p.x <= Math.max(start.x, end.x) &&
				p.y >= Math.min(start.y, end.y) &&
				p.y <= Math.max(start.y, end.y);
			if (
				segmentIntersection(a, b, c, d) ||
				touches(a, c, d) ||
				touches(b, c, d) ||
				touches(c, a, b) ||
				touches(d, a, b)
			)
				return undefined;
		}
	}
	const outline = ensureCCW(points),
		center = { x: 0.5, y: 0.5 };
	const regions = [];
	for (let i = 0; i < outline.length; i++) {
		const a = outline[i]!,
			b = outline[(i + 1) % outline.length]!;
		// Reentrant edges facing away from the bounding-box center do not create a fan face.
		if (cross2(center, a, b) <= 1e-9) continue;
		const end = lineIntersection(
			center,
			{ x: center.x - (b.y - a.y), y: center.y + b.x - a.x },
			a,
			b,
		);
		if (!end) return undefined;
		regions.push({
			points: [
				[0.5, 0.5],
				[a.x, a.y],
				[b.x, b.y],
			] as const,
			start: [0.5, 0.5] as const,
			end: [end.x, end.y] as const,
			angle: 0,
		});
	}
	return regions.length ? { type: 'regions', coordinateSpace: 'shape', regions, stops } : undefined;
}
