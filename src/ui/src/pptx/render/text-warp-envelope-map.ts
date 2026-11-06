/**
 * PowerPoint's point mapping for the WordArt two-curve envelope presets
 * (`textInflate`/`textDeflate`/`textCanUp`/`textCanDown` and the rest of
 * {@link GLYPH_ENVELOPE_PRESETS}), derived from PowerPoint COM renders
 * (2026-09-24, re-measured 2026-09-25; see `docs/guide/visual-effects.md`).
 *
 * The whole text block (every paragraph, laid out unwarped) is normalised by
 * its own bounding box `[left, right] x [top, bottom]` into `(s, v)`, both in
 * `[0, 1]`. Then:
 *
 *   P(s, v) = (1 - v) * T(s) + v * B(s)
 *
 * where `T(s)` is the point at ARC-LENGTH fraction `s` along the preset's top
 * path and `B(s)` the point at arc-length fraction `s` along its bottom path,
 * each measured independently in the shape's own (aspect-correct) units. So
 * the horizontal position is not `s * width`: a steep stretch of a curve
 * (the ends of the `can` cylinder, the shoulders of a strong `deflate`)
 * consumes more arc length per unit of x, which compresses glyphs there and
 * pushes the rest outward. Measured on vertical-stem captions (24 `I` stems,
 * both edges of every stem, 38 preset/adj/aspect combinations) this law
 * lands within ~0.1% of box width on every family, where the previous
 * linear-`x` placement missed `can` by 5-18%.
 *
 * A later change briefly replaced this with linear `x` placement, citing
 * stem positions that were bit-identical between `textCanUp` adj 15000 and
 * adj 50000. That evidence was void: `textCanUp` pins `adj` to
 * [66667, 100000], so both values clamp to 66667 and draw the same shape.
 * A 2026-09-25 whole-outline ink-mask re-measurement against PowerPoint
 * (`scripts/measure-wordart-can-*.{ps1,mjs}`, in-range adj values) confirmed
 * arc length: at the deepest `can` (adj 66667 / 33333) ink IoU rose from
 * 0.35-0.39 under linear `x` to 0.94-0.95, and from 0.86-0.88 to 0.95-0.97 at
 * the shallow end, where the half-ellipse still turns vertical at both ends;
 * `textDeflate` adj 37500 went from 0.57 to 0.96. The residual left at some
 * middle `adj` values is PowerPoint ending a row short of its path end, a
 * non-monotonic flattening artifact this mapping does not model.
 */
import { closedWarpPaths } from './text-warp-closed-paths';
import { ellipsePolygon } from './text-warp-polygon-clip';
import type { ClipPoint } from './text-warp-polygon-clip';
import type { ClosedPathPoint } from './text-warp-closed-paths';
import { envelopeCurveAt, envelopeCurveXExtent } from './text-warp-envelope-curves';

/** The unwarped text block's bounding box, in layout coordinates. */
export interface EnvelopeBlockBox {
	left: number;
	top: number;
	right: number;
	bottom: number;
}

/**
 * Limits the ink of a warp to regions of the shape box: the glyph outlines are
 * cut at the layout `x` positions `splitsX` into consecutive parts, and part
 * `i` is clipped, once mapped, to the convex polygon `regions[i]` (not at all
 * when that is `undefined`).
 */
export interface EnvelopeClip {
	splitsX: number[];
	regions: (ClipPoint[] | undefined)[];
}

/** Maps an unwarped layout point onto the warped shape box. */
export interface EnvelopeWarp {
	map(x: number, y: number): { x: number; y: number };
	clip?: EnvelopeClip;
}

/** Samples per curve; the arc-length inverse is linear between them. */
const CURVE_SAMPLES = 512;

interface ArcTable {
	/** Curve x (box units) at each sample. */
	xs: Float64Array;
	/** Curve y (box units) at each sample. */
	ys: Float64Array;
	/** Cumulative arc length at each sample, normalised to `[0, 1]`. */
	s: Float64Array;
}

/** `startX..endX` is the curve's horizontal span, in box units. */
function buildArcTable(ys: Float64Array, startX: number, endX: number): ArcTable {
	const xs = new Float64Array(ys.length);
	for (let i = 0; i < ys.length; i++) {
		xs[i] = startX + (i / (ys.length - 1)) * (endX - startX);
	}
	return buildArcTableXY(xs, ys);
}

/** The arc-length table of a polyline given by its sample coordinates. */
function buildArcTableXY(xs: Float64Array, ys: Float64Array): ArcTable {
	const s = new Float64Array(ys.length);
	let total = 0;
	for (let i = 0; i < ys.length; i++) {
		if (i > 0) {
			total += Math.hypot(xs[i] - xs[i - 1], ys[i] - ys[i - 1]);
		}
		s[i] = total;
	}
	if (total > 0) {
		for (let i = 0; i < s.length; i++) {
			s[i] /= total;
		}
	} else {
		for (let i = 0; i < s.length; i++) {
			s[i] = i / (s.length - 1);
		}
	}
	return { xs, ys, s };
}

/** The point at arc-length fraction `t` (clamped to `[0, 1]`) along `table`. */
function pointAtArcFraction(table: ArcTable, t: number): { x: number; y: number } {
	const { xs, ys, s } = table;
	const clamped = Math.max(0, Math.min(1, t));
	let lo = 0;
	let hi = s.length - 1;
	while (hi - lo > 1) {
		const mid = (lo + hi) >> 1;
		if (s[mid] <= clamped) {
			lo = mid;
		} else {
			hi = mid;
		}
	}
	const span = s[hi] - s[lo];
	const f = span > 0 ? (clamped - s[lo]) / span : 0;
	return { x: xs[lo] + (xs[hi] - xs[lo]) * f, y: ys[lo] + (ys[hi] - ys[lo]) * f };
}

function tableOfPoints(points: ClosedPathPoint[]): ArcTable {
	return buildArcTableXY(
		Float64Array.from(points, (p) => p.x),
		Float64Array.from(points, (p) => p.y),
	);
}

/** The point at share `s` of a path made of equally weighted pieces. */
function pointAtPieces(pieces: ArcTable[], s: number): { x: number; y: number } {
	const clamped = Math.max(0, Math.min(1, s));
	const index = Math.min(pieces.length - 1, Math.floor(clamped * pieces.length));
	const piece = pieces[index];
	return piece ? pointAtArcFraction(piece, clamped * pieces.length - index) : { x: 0, y: 0 };
}

/**
 * Build the envelope mapping for `preset` in a `width` x `height` shape box,
 * normalising layout points by `block`. Returns `undefined` for a preset
 * outside the envelope family, a degenerate box, or an empty block.
 */
export function createEnvelopeWarp(
	preset: string,
	width: number,
	height: number,
	adj: number | undefined,
	adj2: number | undefined,
	block: EnvelopeBlockBox,
): EnvelopeWarp | undefined {
	const blockW = block.right - block.left;
	const blockH = block.bottom - block.top;
	if (!(width > 0) || !(height > 0) || !(blockW > 0) || !(blockH > 0)) {
		return undefined;
	}
	const closed = closedWarpPaths(preset, width, height, adj, adj2);
	if (closed) {
		const top = closed.top.map(tableOfPoints);
		const bottom = closed.bottom.map(tableOfPoints);
		const pieces = top.length;
		const clips = closed.clips;
		const clip: EnvelopeClip | undefined = clips
			? {
					splitsX: Array.from(
						{ length: pieces - 1 },
						(_, i) => block.left + (blockW * (i + 1)) / pieces,
					),
					regions: clips.map((e) => e && ellipsePolygon(e.cx, e.cy, e.rx, e.ry)),
				}
			: undefined;
		return {
			...(clip ? { clip } : {}),
			map(x: number, y: number) {
				const s = (x - block.left) / blockW;
				const v = (y - block.top) / blockH;
				const t = pointAtPieces(top, s);
				const b = pointAtPieces(bottom, s);
				return { x: t.x + (b.x - t.x) * v, y: t.y + (b.y - t.y) * v };
			},
		};
	}
	const topYs = new Float64Array(CURVE_SAMPLES + 1);
	const bottomYs = new Float64Array(CURVE_SAMPLES + 1);
	for (let i = 0; i <= CURVE_SAMPLES; i++) {
		const band = envelopeCurveAt(preset, i / CURVE_SAMPLES, adj, adj2);
		if (!band) {
			return undefined;
		}
		topYs[i] = band.top * height;
		bottomYs[i] = band.bottom * height;
	}
	const extent = envelopeCurveXExtent(preset, adj, adj2);
	const top = buildArcTable(topYs, extent.top[0] * width, extent.top[1] * width);
	const bottom = buildArcTable(bottomYs, extent.bottom[0] * width, extent.bottom[1] * width);
	return {
		map(x: number, y: number) {
			const s = (x - block.left) / blockW;
			const v = (y - block.top) / blockH;
			const t = pointAtArcFraction(top, s);
			const b = pointAtArcFraction(bottom, s);
			return { x: t.x + (b.x - t.x) * v, y: t.y + (b.y - t.y) * v };
		},
	};
}
