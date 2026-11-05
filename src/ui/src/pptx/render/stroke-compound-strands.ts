/**
 * The parallel strands of a compound line (`a:ln/@cmpd`) drawn along the
 * outline itself, not shifted down the page.
 *
 * A compound strand used to be the outline's own path translated by its offset
 * on the y axis (`translate(0, offset)`), which is right only for a horizontal
 * line. Exported from PowerPoint 16 (`Slide.Export`, 1280 x 720), a rectangle
 * with a 12 pt `thickThin` line paints its thick strand outside and the thin
 * one inside on all four sides; the translated copies drew both strands on the
 * top and bottom edges only, shifted, and one smeared band on the sides.
 *
 * Each strand is instead its own path, the outline offset perpendicular to its
 * direction of travel by the strand's offset (negative to the left, as the
 * first strand of a horizontal line drawn left to right lies above it), with
 * the outline's own joins.
 *
 * @module render/stroke-compound-strands
 */
import type { Point, StrokeClipJoin } from './stroke-clip-outset';
import { flattenPath } from './stroke-clip-outset';
import { offsetRing } from './stroke-clip-edge';

const fmt = (n: number): string => String(Math.round(n * 1000) / 1000);

/** An open polyline offset by `distance` to the right of its travel (y down), mitred. */
function offsetPolyline(points: readonly Point[], distance: number): Point[] {
	const normals = points.slice(1).map((b, k): Point => {
		const a = points[k] ?? b;
		const len = Math.hypot(b[0] - a[0], b[1] - a[1]) || 1;
		return [-(b[1] - a[1]) / len, (b[0] - a[0]) / len];
	});
	return points.map((p, k): Point => {
		const n1 = normals[Math.max(0, k - 1)] ?? [0, 0];
		const n2 = normals[Math.min(normals.length - 1, k)] ?? n1;
		const bx = n1[0] + n2[0];
		const by = n1[1] + n2[1];
		const blen = Math.hypot(bx, by);
		const cos = blen / 2;
		const reach = cos > 0.25 ? distance / cos : distance;
		return blen > 1e-9 ? [p[0] + (bx / blen) * reach, p[1] + (by / blen) * reach] : p;
	});
}

/**
 * The path of one strand `offset` px off the outline `d`, or `undefined` when
 * the path cannot be traced.
 */
export function compoundStrandPath(
	d: string,
	offset: number,
	join: StrokeClipJoin,
): string | undefined {
	const subpaths = flattenPath(d);
	if (!subpaths || subpaths.length === 0) {
		return undefined;
	}
	const closedCount = (d.match(/[zZ]/gu) ?? []).length;
	return subpaths
		.map((sub, index) => {
			const closed = index < closedCount;
			const points = closed
				? offsetRing(sub, Math.abs(offset), join, offset < 0 ? -1 : 1)
				: offsetPolyline(sub, offset);
			const body = points.map(([x, y]) => `${fmt(x)} ${fmt(y)}`).join(' L ');
			return `M ${body}${closed ? ' Z' : ''}`;
		})
		.join(' ');
}

/** Strands whose offsets are traced as their own paths along `d` (offset `0` after). */
export function traceCompoundStrands<T extends { offset: number; d?: string }>(
	strands: readonly T[],
	d: string,
	join: StrokeClipJoin,
): T[] {
	if (strands.length < 2) {
		return [...strands];
	}
	return strands.map((strand) => {
		if (strand.offset === 0 || strand.d !== undefined) {
			return strand;
		}
		const traced = compoundStrandPath(d, strand.offset, join);
		return traced ? { ...strand, d: traced, offset: 0 } : strand;
	});
}
