/**
 * Convex polygon clipping (Sutherland-Hodgman) for the warped glyph outlines
 * of the closed-path presets, whose ink is limited to an ellipse (see
 * `text-warp-closed-paths.ts`).
 */

/** A point of a polygon. */
export interface ClipPoint {
	x: number;
	y: number;
}

/** Vertices of an ellipse in increasing angle, as a convex clip polygon. */
export function ellipsePolygon(
	cx: number,
	cy: number,
	rx: number,
	ry: number,
	vertices = 192,
): ClipPoint[] {
	const points: ClipPoint[] = [];
	for (let i = 0; i < vertices; i++) {
		const angle = (2 * Math.PI * i) / vertices;
		points.push({ x: cx + rx * Math.cos(angle), y: cy + ry * Math.sin(angle) });
	}
	return points;
}

function intersect(p: ClipPoint, q: ClipPoint, dp: number, dq: number): ClipPoint {
	const t = dp / (dp - dq);
	return { x: p.x + (q.x - p.x) * t, y: p.y + (q.y - p.y) * t };
}

/** One Sutherland-Hodgman pass: keep the points whose signed distance is not negative. */
function clipByHalfPlane(polygon: ClipPoint[], distance: (p: ClipPoint) => number): ClipPoint[] {
	const out: ClipPoint[] = [];
	for (let j = 0; j < polygon.length; j++) {
		const current = polygon[j]!;
		const previous = polygon[(j + polygon.length - 1) % polygon.length]!;
		const dc = distance(current);
		const dp = distance(previous);
		if (dc >= 0) {
			if (dp < 0) {
				out.push(intersect(previous, current, dp, dc));
			}
			out.push(current);
		} else if (dp >= 0) {
			out.push(intersect(previous, current, dp, dc));
		}
	}
	return out;
}

/**
 * `subject` clipped to the convex polygon `clip` (in the orientation
 * {@link ellipsePolygon} produces). A concave subject keeps its fill, with
 * zero-width bridges along the clip boundary where it is cut in several parts.
 */
export function clipPolygonConvex(subject: ClipPoint[], clip: ClipPoint[]): ClipPoint[] {
	let output = subject;
	for (let i = 0; i < clip.length && output.length > 0; i++) {
		const a = clip[i]!;
		const b = clip[(i + 1) % clip.length]!;
		output = clipByHalfPlane(output, (p) => (b.x - a.x) * (p.y - a.y) - (b.y - a.y) * (p.x - a.x));
	}
	return output;
}

/** The part of `polygon` with `x` below `limit` (`side` -1) or at or above it (`side` 1). */
export function clipPolygonByX(polygon: ClipPoint[], limit: number, side: -1 | 1): ClipPoint[] {
	return clipByHalfPlane(polygon, (p) => side * (p.x - limit));
}
