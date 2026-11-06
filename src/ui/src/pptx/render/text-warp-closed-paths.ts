/**
 * The envelope presets whose two paths are not functions of x: the two rings
 * and Circle Pour. Their text runs once around a closed loop, so the warp maps
 * the text block onto a pair of loops (see `text-warp-envelope-map.ts`).
 *
 * Measured on PowerPoint 16 (rows of full-block glyphs and letter rows, at
 * several adjusts, font sizes and text lengths):
 *
 * - Circle Pour: the top path is the box ellipse, the bottom path the same
 *   ellipse scaled by `adj2` (default 50%) around the centre. Both start at
 *   the angle `adj1` (default 180 degrees, the left tip) and run clockwise.
 *   Glyph tops point outward.
 * - Ring Inside: two congruent ellipses, one above the other, each as wide as
 *   the box and `a h` tall (`a` is `adj`, pinned to 50%..99%, default 60%),
 *   one touching the top of the box and one the bottom. The ink is where
 *   exactly one of them covers: the first half of the text runs along the top
 *   half of the upper ellipse left to right, its glyphs reaching down to the
 *   top half of the lower ellipse, and the second half runs back along the
 *   bottom half of the lower ellipse (upside down), reaching up to the bottom
 *   half of the upper one. The sweep between the two arcs overshoots the
 *   ellipse near the left and right tips, where PowerPoint draws nothing, so
 *   a glyph in the upper half is clipped to the upper ellipse and one in the
 *   lower half to the lower ellipse.
 * - Ring Outside: the same bands, but the text starts along the lower band
 *   left to right and glyph tops point at the middle of the box.
 *
 * Every piece covers an equal share of the text: the two arcs of a ring are
 * each half of the loop.
 */

/** A point in shape-box units. */
export interface ClosedPathPoint {
	x: number;
	y: number;
}

/** An axis-aligned ellipse in shape-box units. */
export interface ClosedPathEllipse {
	cx: number;
	cy: number;
	rx: number;
	ry: number;
}

/** A top and a bottom path, each as pieces covering equal shares of the text. */
export interface ClosedWarpPaths {
	top: ClosedPathPoint[][];
	bottom: ClosedPathPoint[][];
	/** Per piece, the ellipse the ink of that piece is clipped to (none for an unclipped piece). */
	clips?: (ClosedPathEllipse | undefined)[];
}

/** Presets drawn with {@link closedWarpPaths}. */
export const CLOSED_PATH_PRESETS: ReadonlySet<string> = new Set([
	'textRingInside',
	'textRingOutside',
	'textCirclePour',
]);

const SAMPLES = 256;
const DEFAULT_RING_ADJ = 60000;
const DEFAULT_POUR_ANGLE = 10800000;
const DEFAULT_POUR_SCALE = 50000;

function clamp(n: number, lo: number, hi: number): number {
	return Math.max(lo, Math.min(hi, n));
}

/** Top or bottom half of an ellipse, from the left tip to the right one (or back). */
function halfEllipse(
	e: ClosedPathEllipse,
	half: 'top' | 'bottom',
	reverse: boolean,
): ClosedPathPoint[] {
	const points: ClosedPathPoint[] = [];
	const sign = half === 'top' ? -1 : 1;
	for (let i = 0; i <= SAMPLES; i++) {
		const theta = (Math.PI * (reverse ? SAMPLES - i : i)) / SAMPLES;
		points.push({
			x: e.cx - e.rx * Math.cos(theta),
			y: e.cy + sign * e.ry * Math.sin(theta),
		});
	}
	return points;
}

function ringPaths(preset: string, w: number, h: number, adj: number | undefined): ClosedWarpPaths {
	const a = clamp((adj ?? DEFAULT_RING_ADJ) / 100000, 0.5, 0.99);
	const ry = (a * h) / 2;
	const upper: ClosedPathEllipse = { cx: w / 2, cy: ry, rx: w / 2, ry };
	const lower: ClosedPathEllipse = { cx: w / 2, cy: h - ry, rx: w / 2, ry };
	if (preset === 'textRingOutside') {
		return {
			top: [halfEllipse(upper, 'bottom', false), halfEllipse(lower, 'top', true)],
			bottom: [halfEllipse(lower, 'bottom', false), halfEllipse(upper, 'top', true)],
			clips: [lower, upper],
		};
	}
	return {
		top: [halfEllipse(upper, 'top', false), halfEllipse(lower, 'bottom', true)],
		bottom: [halfEllipse(lower, 'top', false), halfEllipse(upper, 'bottom', true)],
		clips: [upper, lower],
	};
}

function ellipseLoop(w: number, h: number, scale: number, startAngle: number): ClosedPathPoint[] {
	const points: ClosedPathPoint[] = [];
	for (let i = 0; i <= SAMPLES; i++) {
		const angle = startAngle + (2 * Math.PI * i) / SAMPLES;
		points.push({
			x: w / 2 + (w / 2) * scale * Math.cos(angle),
			y: h / 2 + (h / 2) * scale * Math.sin(angle),
		});
	}
	return points;
}

function pourPaths(
	w: number,
	h: number,
	adj: number | undefined,
	adj2: number | undefined,
): ClosedWarpPaths {
	const startAngle = (((adj ?? DEFAULT_POUR_ANGLE) % 21600000) / 60000) * (Math.PI / 180);
	const scale = clamp((adj2 ?? DEFAULT_POUR_SCALE) / 100000, 0, 0.99);
	return {
		top: [ellipseLoop(w, h, 1, startAngle)],
		bottom: [ellipseLoop(w, h, scale, startAngle)],
	};
}

/** The top and bottom loops of a closed-path preset, or `undefined` for any other preset. */
export function closedWarpPaths(
	preset: string,
	width: number,
	height: number,
	adj: number | undefined,
	adj2: number | undefined,
): ClosedWarpPaths | undefined {
	if (preset === 'textCirclePour') {
		return pourPaths(width, height, adj, adj2);
	}
	if (preset === 'textRingInside' || preset === 'textRingOutside') {
		return ringPaths(preset, width, height, adj);
	}
	return undefined;
}
