/**
 * The strands of a compound outline (`a:ln/@cmpd`: `dbl`, `thickThin`,
 * `thinThick`, `tri`) carried through a resizing Morph.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object): a
 * 12 pt compound line on a rectangle (and a rounded rectangle) growing from
 * 120 x 80 pt to 480 x 240 pt keeps its band proportions on every frame (a
 * `thickThin` edge reads 10 + 3 px at the start, 15 + 4 across the vertical
 * edges at the widest, 10 + 3 again at the end) while the whole line widens
 * and narrows by the same `W(p) * lerp(w / W0, w / W1, p)` an ordinary line
 * follows. Each strand and the gaps between them scale together, so a strand
 * is its own path traced at its offset in the same path units the plain
 * outline carry uses, and the path's anisotropic scale does the rest.
 *
 * Every strand is its own `<path>` of the outline overlay, so each gets its own
 * baked track (the plan rules address them by position).
 *
 * @module render/morph-shape-carry-strands
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import { getCompoundLineOffsets } from './connector-style';
import { bakeStops } from './morph-shape-carry-path';
import { compoundStrandPath } from './stroke-compound-strands';
import type { StrokeClipJoin } from './stroke-clip-outset';
import { buildStrokeOutline } from './stroke-outline';

/** The offsets of an element's compound strands, when the outline overlay paints them as traced paths. */
export function tracedCompoundOffsets(element: PptxElement): number[] | undefined {
	const outline = buildStrokeOutline(element);
	if (!outline || outline.markers || !hasShapeProperties(element)) {
		return undefined;
	}
	const offsets = getCompoundLineOffsets(element.shapeStyle?.compoundLine, outline.strokeWidth);
	return offsets.length > 1 &&
		outline.strands.length === offsets.length &&
		outline.strands.every((strand) => strand.offset === 0)
		? offsets
		: undefined;
}

/** The compound line of a shape, `undefined` for a single line. */
export function compoundLineOf(element: PptxElement): string | undefined {
	const kind = hasShapeProperties(element) ? element.shapeStyle?.compoundLine : undefined;
	return kind && kind !== 'sng' ? kind : undefined;
}

/** How many path points a traced strand has: equal at every stop or CSS cannot tween it. */
function pointCount(d: string): number {
	return d.split('L').length;
}

/**
 * One keyframes body per strand, each stop carrying the strand's path traced
 * along the stop's outline (`userPathAt`, in the path's user units) and the
 * scale that undoes the journey's anisotropy. `undefined` when a strand cannot
 * be traced at some stop or changes its number of points between stops.
 */
export function bakeStrandStops(
	progress: readonly number[],
	offsets: readonly number[],
	join: StrokeClipJoin,
	userPathAt: (p: number) => { d: string; scale: string },
): string[] | undefined {
	const tracks = offsets.map(() => new Map<number, string[]>());
	const counts: number[] = [];
	for (const p of progress) {
		const { d, scale } = userPathAt(p);
		for (const [i, offset] of offsets.entries()) {
			const strand = offset === 0 ? d : compoundStrandPath(d, offset, join);
			if (!strand) {
				return undefined;
			}
			const points = pointCount(strand);
			if ((counts[i] ?? points) !== points) {
				return undefined;
			}
			counts[i] = points;
			tracks[i]?.set(p, [
				`d: path('${strand}')`,
				'transform-origin: 0px 0px',
				`transform: ${scale}`,
			]);
		}
	}
	return tracks.map((track) => bakeStops(progress, (p) => track.get(p) ?? []));
}
