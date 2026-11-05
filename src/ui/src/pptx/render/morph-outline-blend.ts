/**
 * Two outlines of DIFFERENT structure turned into one structure, so a morph can
 * move their vertices instead of dissolving one shape into the other.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object):
 *
 *  - a freeform that gains two nodes between the slides (4 -> 6 vertices)
 *    morphs as ONE outline: the 4-gon grows extra vertices along its edges and
 *    every vertex walks to its place on the 6-gon, the outline never
 *    dissolves;
 *  - a `!!`-named rectangle that becomes an ellipse morphs the same way, and
 *    in PATH order: the rectangle's first corner (top left) walks to the
 *    ellipse's first point (its left end), so mid-morph the shape reads as a
 *    rounded quad turned about 45 degrees, its corners pointing up-right and
 *    down-left.
 *
 * So both outlines are walked from their own first point in their own order,
 * flattened to polylines (`flattenPath`), and parameterised by arc length;
 * every vertex of either ring is inserted into the other at the same
 * arc-length fraction. Each ring keeps its own corners exactly (so the morph
 * starts and ends on the real outlines), and the two rings pair up vertex for
 * vertex.
 *
 * @module render/morph-outline-blend
 */
import { serializeSvgPath } from './morph-svg-path';
import type { SvgPathCommand } from './morph-types';
import type { Point } from './stroke-clip-outset';
import { flattenPath } from './stroke-clip-outset';

/** Most points a blended ring may carry; beyond it the pair is not blended. */
const MAX_POINTS = 400;

/** The one closed ring an outline draws, or `undefined` for several sub-paths. */
function ringOf(commands: readonly SvgPathCommand[]): Point[] | undefined {
	const subpaths = flattenPath(serializeSvgPath([...commands]));
	const ring = subpaths?.length === 1 ? [...(subpaths[0] ?? [])] : undefined;
	const first = ring?.[0];
	const last = ring?.[ring.length - 1];
	if (ring && first && last && ring.length > 1 && first[0] === last[0] && first[1] === last[1]) {
		ring.pop();
	}
	return ring && ring.length >= 3 ? ring : undefined;
}

/** Each point's arc-length fraction around the closed ring, from 0. */
function fractions(ring: readonly Point[]): number[] | undefined {
	const lengths = [0];
	for (let i = 1; i <= ring.length; i++) {
		const a = ring[i - 1] ?? [0, 0];
		const b = ring[i % ring.length] ?? [0, 0];
		lengths.push((lengths[i - 1] ?? 0) + Math.hypot(b[0] - a[0], b[1] - a[1]));
	}
	const total = lengths[ring.length] ?? 0;
	return total > 0 ? lengths.slice(0, ring.length).map((l) => l / total) : undefined;
}

/** The ring's point at arc-length fraction `t`. */
function pointAt(ring: readonly Point[], at: readonly number[], t: number): Point {
	for (let i = 0; i < ring.length; i++) {
		const t0 = at[i] ?? 0;
		const t1 = i + 1 < ring.length ? (at[i + 1] ?? 1) : 1;
		if (t <= t1 || i === ring.length - 1) {
			const a = ring[i] ?? [0, 0];
			const b = ring[(i + 1) % ring.length] ?? a;
			const u = t1 > t0 ? (t - t0) / (t1 - t0) : 0;
			return [a[0] + (b[0] - a[0]) * u, a[1] + (b[1] - a[1]) * u];
		}
	}
	return ring[0] ?? [0, 0];
}

function ringCommands(points: readonly Point[]): SvgPathCommand[] {
	return [
		...points.map((p, i) => ({ type: i === 0 ? 'M' : 'L', values: [p[0], p[1]] })),
		{ type: 'Z', values: [] },
	];
}

/**
 * Two outlines as two rings of the same length whose points correspond, or
 * `undefined` when either is not one closed sub-path.
 */
export function blendOutlineCommands(
	a: readonly SvgPathCommand[],
	b: readonly SvgPathCommand[],
): [SvgPathCommand[], SvgPathCommand[]] | undefined {
	const ringA = ringOf(a);
	const ringB = ringOf(b);
	const atA = ringA && fractions(ringA);
	const atB = ringB && fractions(ringB);
	if (!ringA || !ringB || !atA || !atB) return undefined;
	const all = [...atA, ...atB].sort((x, y) => x - y);
	const merged = all.filter((t, i) => i === 0 || t - (all[i - 1] ?? 0) > 1e-4);
	if (merged.length > MAX_POINTS) return undefined;
	return [
		ringCommands(merged.map((t) => pointAt(ringA, atA, t))),
		ringCommands(merged.map((t) => pointAt(ringB, atB, t))),
	];
}
