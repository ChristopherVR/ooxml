/**
 * The `box-shadow` list that cross-dissolves a morphing shape's two outer
 * shadows the way PowerPoint does: per pixel `(1 - p) A + p B`.
 *
 * Measured on PowerPoint 16 (`CreateVideo`, 60 fps, 720p, Morph by object,
 * 2 s) on a rectangle growing from 200 x 120 pt to 440 x 240 pt with a hard
 * black shadow 20 pt at 45 degrees: where the outgoing and the incoming shadow
 * overlap, the shadow reads alpha 1.00 at EVERY frame (0.99-1.00 with a fill
 * change as well), while each one alone reads `1 - p` and `p`. With the
 * shadow at 50 % transparency the overlap holds 0.52 throughout and each alone
 * fades from 0.52. That is a true cross-dissolve of two renders, not two fades
 * stacked: two layers of one `box-shadow` composite source-over, so the
 * overlap dipped to 0.75 at the midpoint.
 *
 * The overlap of two hard shadows of the same box is itself a rectangle, so an
 * extra layer on top of the two fades can paint exactly the missing alpha
 * there: `box-shadow` draws a box-shaped rectangle, a negative spread shrinks
 * it evenly, and a few of them side by side tile any narrower rectangle (tiles
 * that cross each other only composite right at alpha 1, so a translucent
 * overlap gets one centred tile and leaves two end strips to the plain fades).
 * Its alpha and colour are solved per keyframe so the stack composites to the
 * dissolve (`overlapColor`). Blurred shadows get the same layer with the
 * smaller blur: exact where both are solid, an approximation in their fringes
 * (in a 1-D model of a 300 px shadow, mean error 0.003-0.028 against 0.11-0.135
 * without it, for blur deviations of 5-26 px).
 *
 * Measured after (same deck, ours vs PowerPoint): the hard shadow's overlap
 * reads 1.00 at 10-90 % where it read 0.93, 0.83, 0.77, 0.77, 0.81, 0.86, and
 * the translucent one holds 0.50, its own still alpha (PowerPoint's is 0.52),
 * where it dipped to 0.45.
 *
 * @module render/morph-effect-shadow
 */

/** One parsed `box-shadow` layer, lengths in the half's local px. */
export interface ShadowLayer {
	inset: boolean;
	/** Offset x, offset y, then optional blur and spread. */
	lengths: number[];
	/** RGB 0-255 and alpha 0-1. */
	color: [number, number, number, number];
}

/** A rectangle in the element's local frame. */
interface Rect {
	readonly left: number;
	readonly top: number;
	readonly right: number;
	readonly bottom: number;
}

/** The most layers one overlap is tiled with before it is left out. */
const MAX_TILES = 6;

function round(value: number): number {
	return Math.round(value * 1000) / 1000;
}

function clampUnit(value: number): number {
	return Math.min(1, Math.max(0, value));
}

/** The rectangle an outer layer paints before its blur. */
function coreRect(layer: ShadowLayer, width: number, height: number): Rect {
	const [dx = 0, dy = 0, , spread = 0] = layer.lengths;
	return {
		left: dx - spread,
		top: dy - spread,
		right: width + dx + spread,
		bottom: height + dy + spread,
	};
}

/**
 * Box-shaped rectangles (offset and spread) that tile `rect` exactly, or
 * `undefined` when that takes more than `maxTiles`. With `maxTiles` 1 the one
 * rectangle is centred in `rect` and may leave strips of it at two ends.
 */
function tile(
	rect: Rect,
	width: number,
	height: number,
	maxTiles: number,
): Array<{ dx: number; dy: number; spread: number }> | undefined {
	const w = rect.right - rect.left;
	const h = rect.bottom - rect.top;
	if (w <= 0.5 || h <= 0.5) {
		return [];
	}
	const spread = Math.min((w - width) / 2, (h - height) / 2);
	const tw = width + 2 * spread;
	const th = height + 2 * spread;
	if (tw <= 0.5 || th <= 0.5) {
		return undefined;
	}
	const nx = maxTiles === 1 ? 1 : Math.max(1, Math.ceil(w / tw - 1e-6));
	const ny = maxTiles === 1 ? 1 : Math.max(1, Math.ceil(h / th - 1e-6));
	if (nx * ny > maxTiles) {
		return undefined;
	}
	const tiles: Array<{ dx: number; dy: number; spread: number }> = [];
	for (let i = 0; i < nx; i++) {
		for (let j = 0; j < ny; j++) {
			const cx =
				nx === 1 ? (rect.left + rect.right) / 2 : rect.left + tw / 2 + ((w - tw) * i) / (nx - 1);
			const cy =
				ny === 1 ? (rect.top + rect.bottom) / 2 : rect.top + th / 2 + ((h - th) * j) / (ny - 1);
			tiles.push({ dx: cx - width / 2, dy: cy - height / 2, spread });
		}
	}
	return tiles;
}

/** Where two outer layers overlap, as layers to paint there, or `undefined` when none can. */
export interface ShadowOverlap {
	readonly tiles: ReadonlyArray<{ dx: number; dy: number; spread: number }>;
	readonly blur: number;
}

/**
 * The overlap of outgoing layer `a` and incoming layer `b`, both already in
 * the half's local frame of `width` x `height`.
 */
export function shadowOverlap(
	a: ShadowLayer,
	b: ShadowLayer,
	width: number,
	height: number,
): ShadowOverlap | undefined {
	if (a.inset || b.inset) {
		return undefined;
	}
	const ra = coreRect(a, width, height);
	const rb = coreRect(b, width, height);
	const tiles = tile(
		{
			left: Math.max(ra.left, rb.left),
			top: Math.max(ra.top, rb.top),
			right: Math.min(ra.right, rb.right),
			bottom: Math.min(ra.bottom, rb.bottom),
		},
		width,
		height,
		// Tiles overlap each other, which only composites right at alpha 1: a
		// translucent overlap is painted once, centred, leaving its two end
		// strips to the plain fades.
		a.color[3] >= 1 && b.color[3] >= 1 ? MAX_TILES : 1,
	);
	if (!tiles || tiles.length === 0) {
		return undefined;
	}
	return { tiles, blur: Math.min(a.lengths[2] ?? 0, b.lengths[2] ?? 0) };
}

type Rgba = [number, number, number, number];

/**
 * The overlap layer's colour at progress `p`, painted over the fades below it.
 *
 * `alone`: this half paints both fades itself, `b` at `p` over `a` at `1 - p`,
 * and the stack has to composite to `(1 - p) A + p B`. `underGhost`: the
 * outgoing ghost paints `A` at `1 - p` OVER this half (`b` at `p` only), so
 * this half's own stack has to be whatever that composites to the dissolve.
 */
export function overlapColor(
	a: Rgba,
	b: Rgba,
	p: number,
	mode: 'alone' | 'underGhost',
): Rgba | undefined {
	const qa = (1 - p) * a[3];
	const pb = p * b[3];
	if (mode === 'underGhost') {
		if (qa >= 1) {
			return undefined;
		}
		const want = clampUnit(pb / (1 - qa));
		const alpha = pb >= 1 ? 0 : clampUnit((want - pb) / (1 - pb));
		return [b[0], b[1], b[2], alpha];
	}
	const ua = pb + qa * (1 - pb);
	if (ua >= 1) {
		return [b[0], b[1], b[2], 0];
	}
	const alpha = clampUnit((qa * pb) / (1 - ua));
	if (alpha <= 0) {
		return [b[0], b[1], b[2], 0];
	}
	const channel = (i: 0 | 1 | 2) => {
		const target = qa * a[i] + pb * b[i];
		const under = pb * b[i] + qa * a[i] * (1 - pb);
		return Math.min(255, Math.max(0, (target - under * (1 - alpha)) / alpha));
	};
	return [channel(0), channel(1), channel(2), alpha];
}

/** One layer as `box-shadow` text, with its colour's alpha multiplied by `alpha`. */
export function serializeShadow(layer: ShadowLayer, alpha: number): string {
	const lengths = layer.lengths.map((value) => `${round(value)}px`);
	const [r, g, b, a] = layer.color;
	return `${layer.inset ? 'inset ' : ''}${lengths.join(' ')} rgba(${round(r)}, ${round(g)}, ${round(b)}, ${round(a * alpha)})`;
}

/** The overlap's tiles painted in `color`. */
export function serializeOverlap(overlap: ShadowOverlap, color: Rgba): string[] {
	return overlap.tiles.map((t) =>
		serializeShadow({ inset: false, lengths: [t.dx, t.dy, overlap.blur, t.spread], color }, 1),
	);
}
