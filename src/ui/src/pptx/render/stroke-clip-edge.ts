/**
 * `stroke-clip-edge`: the outer edge of a widened outline clip, as path data
 * the outline overlay strokes once more so the fill never shows along it.
 *
 * A shape's fill is its wrapper's background, and the wrapper is clipped to
 * the geometry widened by half the outline (`stroke-clip-outset`), so the fill
 * reaches the clip's edge too. The outline ends exactly there, and wherever it
 * only partly covers an edge pixel (anti-aliasing, a mitred point) the fill
 * underneath showed through: issue #144's white map markers grew an orange
 * fringe, plainest at fit zoom where the slide is scaled down. Stroking this
 * edge a couple of device pixels wide in the outline's own paint puts solid
 * outline under every pixel the clip cuts through; the clip trims the half
 * that lies outside it, and the inner half lies on the outline it matches.
 *
 * @module render/stroke-clip-edge
 */
import type { Point, StrokeClipJoin } from './stroke-clip-outset';
import { flattenPath } from './stroke-clip-outset';

/** Points that stand in for a round join's arc. */
const ROUND_JOIN_STEPS = 8;

const fmt = (n: number): string => String(Math.round(n * 1000) / 1000);

function signedArea(pts: readonly Point[]): number {
	let area = 0;
	for (let k = 0; k < pts.length; k++) {
		const a = pts[k] as Point;
		const b = pts[(k + 1) % pts.length] as Point;
		area += a[0] * b[1] - b[0] * a[1];
	}
	return area;
}

/** One ring offset away from the shape's fill by `half`, joined as `join` joins. */
function offsetRing(
	pts: readonly Point[],
	half: number,
	join: StrokeClipJoin,
	away: 1 | -1,
): Point[] {
	let last: Point = [0, 0];
	// `away * normal` points away from the fill: the side the outline's outer half lies on.
	const normals: Point[] = pts.map((a, k) => {
		const b = pts[(k + 1) % pts.length] as Point;
		const len = Math.hypot(b[0] - a[0], b[1] - a[1]);
		if (len >= 1e-6) {
			last = [(away * -(b[1] - a[1])) / len, (away * (b[0] - a[0])) / len];
		}
		return last;
	});
	const ring: Point[] = [];
	for (let k = 0; k < pts.length; k++) {
		const v = pts[k] as Point;
		const n1 = normals[(k - 1 + pts.length) % pts.length] as Point;
		const n2 = normals[k] as Point;
		const p1: Point = [v[0] + n1[0] * half, v[1] + n1[1] * half];
		const p2: Point = [v[0] + n2[0] * half, v[1] + n2[1] * half];
		const cosHalf = Math.sqrt(Math.max(0, (1 + n1[0] * n2[0] + n1[1] * n2[1]) / 2));
		const bx = n1[0] + n2[0];
		const by = n1[1] + n2[1];
		const blen = Math.hypot(bx, by);
		const miter: Point | undefined =
			cosHalf > 1e-6 && blen > 1e-9
				? [v[0] + (bx / blen) * (half / cosHalf), v[1] + (by / blen) * (half / cosHalf)]
				: undefined;
		// The outer side of the turn is where the two offset edges part; inside a
		// turn they cross, and that crossing is the edge whatever the join.
		const turn = n1[0] * n2[1] - n1[1] * n2[0];
		// The normals point away from the fill, so the corner is convex (the two
		// offset edges part) exactly when the path turns towards the fill.
		const opensOutward = away * turn < 0;
		if (!opensOutward || Math.abs(turn) < 1e-9) {
			ring.push(miter ?? p2);
		} else if (join.join === 'miter' && miter && 1 / cosHalf <= join.miterLimit) {
			ring.push(miter);
		} else if (join.join === 'round') {
			const a1 = Math.atan2(n1[1], n1[0]);
			let sweep = Math.atan2(n2[1], n2[0]) - a1;
			while (sweep > Math.PI) sweep -= 2 * Math.PI;
			while (sweep < -Math.PI) sweep += 2 * Math.PI;
			for (let i = 0; i <= ROUND_JOIN_STEPS; i++) {
				const t = a1 + (sweep * i) / ROUND_JOIN_STEPS;
				ring.push([v[0] + Math.cos(t) * half, v[1] + Math.sin(t) * half]);
			}
		} else {
			ring.push(p1, p2);
		}
	}
	return ring;
}

/**
 * Path data along the outer edge of the clip `outsetClipPathForStroke` builds
 * for the same outline, or `undefined` when it cannot be traced.
 *
 * @param d           The outline's path data in element pixels.
 * @param strokeWidth The width the clip was widened for.
 * @param join        The outline's join, which shapes the clip's corners.
 */
export function outerStrokeEdgePath(
	d: string,
	strokeWidth: number,
	join: StrokeClipJoin,
): string | undefined {
	if (!(strokeWidth > 0)) {
		return undefined;
	}
	const subpaths = flattenPath(d);
	if (!subpaths || subpaths.length === 0) {
		return undefined;
	}
	// The first ring is the shape's outside; a ring wound the other way is a
	// hole, whose fill lies on its other side.
	const reference = Math.sign(signedArea(subpaths[0] as Point[])) || 1;
	return subpaths
		.filter((sub) => sub.length > 2)
		.map((sub) => {
			const sign = Math.sign(signedArea(sub)) || reference;
			// For a ring wound like a positive-area square in y-down space, the left
			// normal (-dy, dx) points into the ring.
			const away: 1 | -1 = sign > 0 ? -1 : 1;
			const ring = offsetRing(
				sub,
				strokeWidth / 2,
				join,
				sign === reference ? away : away === 1 ? -1 : 1,
			);
			return `M ${ring.map(([x, y]) => `${fmt(x)} ${fmt(y)}`).join(' L ')} Z`;
		})
		.join(' ');
}
