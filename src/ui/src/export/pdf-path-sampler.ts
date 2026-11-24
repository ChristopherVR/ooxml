/**
 * Points along a path by arc length, for text set on a path (WordArt warps): the position and
 * direction at a distance from the start. Beyond either end the path is continued straight
 * along its end direction. DOM-free.
 */
import { flattenSegments } from './pdf-raster';
import type { PathSegment } from './pdf-path';

export interface PathPoint {
	x: number;
	y: number;
	/** Direction of travel in radians (y down: positive turns clockwise on screen). */
	angle: number;
}

export interface PathSampler {
	/** Total length. */
	length: number;
	/** The point `distance` along the path. */
	at(distance: number): PathPoint;
}

/** A sampler for the first subpath of `segments`; `undefined` when it has no length. */
export function pathSampler(
	segments: readonly PathSegment[],
	step = 0.25,
): PathSampler | undefined {
	const lines = flattenSegments(segments, undefined, step);
	const points: (readonly [number, number])[] = [];
	for (const line of lines) {
		for (const point of line.points) {
			const last = points[points.length - 1];
			if (!last || Math.hypot(point[0] - last[0], point[1] - last[1]) > 1e-9) points.push(point);
		}
		if (line.closed && line.points[0]) points.push(line.points[0]);
	}
	if (points.length < 2) return undefined;
	const lengths = [0];
	for (let index = 1; index < points.length; index++) {
		const [ax, ay] = points[index - 1]!;
		const [bx, by] = points[index]!;
		lengths.push(lengths[index - 1]! + Math.hypot(bx - ax, by - ay));
	}
	const length = lengths[lengths.length - 1]!;
	if (!(length > 0)) return undefined;
	const direction = (index: number) => {
		const [ax, ay] = points[index]!;
		const [bx, by] = points[index + 1]!;
		return Math.atan2(by - ay, bx - ax);
	};
	return {
		length,
		at(distance) {
			let lo = 0;
			let hi = lengths.length - 1;
			const target = Math.max(0, Math.min(length, distance));
			while (hi - lo > 1) {
				const mid = (lo + hi) >> 1;
				if (lengths[mid]! <= target) lo = mid;
				else hi = mid;
			}
			const span = lengths[hi]! - lengths[lo]!;
			const t = span > 0 ? (target - lengths[lo]!) / span : 0;
			const [ax, ay] = points[lo]!;
			const [bx, by] = points[hi]!;
			const angle = direction(lo);
			// Past an end, carry on straight.
			const beyond = distance - target;
			return {
				x: ax + (bx - ax) * t + Math.cos(angle) * beyond,
				y: ay + (by - ay) * t + Math.sin(angle) * beyond,
				angle,
			};
		},
	};
}
