import { archEllipsePath, circleEllipsePath, warpLineOutward } from './text-warp-arch';
import {
	clamp4,
	archUpPath,
	archDownPath,
	circlePath,
	wave1Path,
	inflatePath,
	deflatePath,
	curveUpPath,
	curveDownPath,
	wave2Path,
	wave4Path,
	doubleWave1Path,
	canUpPath,
	canDownPath,
	buttonPath,
	ringInsidePath,
	ringOutsidePath,
} from './text-warp-paths';

/** Produces an SVG path `d` attribute for a single text line at vertical
 *  position `t` (0 = top, 1 = bottom). Optional adj/adj2 are raw OOXML units. */
export type WarpPathGenerator = (
	w: number,
	h: number,
	t: number,
	adj?: number,
	adj2?: number,
) => string;

/** Cascading up: lines tilt from lower-left to upper-right. */
export function cascadeUpPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 44444 : 1;
	const tilt = 0.2 * clamp4(adjFactor);
	const yMid = h * (0.2 + t * 0.6);
	const yStart = yMid + (h * tilt) / 2;
	const yEnd = yMid - (h * tilt) / 2;
	return `M 0,${yStart} L ${w},${yEnd}`;
}

/** Cascading down: lines tilt from upper-left to lower-right. */
export function cascadeDownPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 44444 : 1;
	const tilt = 0.2 * clamp4(adjFactor);
	const yMid = h * (0.2 + t * 0.6);
	const yStart = yMid - (h * tilt) / 2;
	const yEnd = yMid + (h * tilt) / 2;
	return `M 0,${yStart} L ${w},${yEnd}`;
}

// ── Priority 3 ──────────────────────────────────────────────────────────

/** Triangle / trapezoid: top line narrow, bottom line full width. */
function trianglePath(w: number, h: number, t: number, adj?: number): string {
	const adjRatio = adj !== undefined ? adj / 100000 : 0.5;
	const narrowW = w * (1 - Math.max(0, Math.min(adjRatio, 1))) * 0.3;
	const lineW = narrowW + t * (w - narrowW);
	const xStart = (w - lineW) / 2;
	const yBase = h * (0.1 + t * 0.8);
	return `M ${xStart},${yBase} L ${xStart + lineW},${yBase}`;
}

/** Inverted triangle: top line full width, bottom line narrow. */
function triangleInvertedPath(w: number, h: number, t: number, adj?: number): string {
	const adjRatio = adj !== undefined ? adj / 100000 : 0.5;
	const narrowW = w * (1 - Math.max(0, Math.min(adjRatio, 1))) * 0.3;
	const lineW = w - t * (w - narrowW);
	const xStart = (w - lineW) / 2;
	const yBase = h * (0.1 + t * 0.8);
	return `M ${xStart},${yBase} L ${xStart + lineW},${yBase}`;
}

/** Stop / octagon: lines narrow at top and bottom, widest in centre. */
function stopPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 25000 : 1;
	const insetScale = clamp4(adjFactor);
	const inset = w * 0.15 * (1 - (1 - 2 * Math.abs(t - 0.5)) ** 2) * insetScale;
	const yBase = h * (0.1 + t * 0.8);
	return `M ${inset},${yBase} L ${w - inset},${yBase}`;
}

/** Chevron: V-shape pointing down. */
function chevronPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const adjFactor = adj !== undefined ? adj / 25000 : 1;
	const dip = h * 0.2 * (1 - t) * clamp4(adjFactor);
	return `M 0,${yBase} L ${w / 2},${yBase + dip} L ${w},${yBase}`;
}

/** Inverted chevron: V-shape pointing up. */
function chevronInvertedPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const adjFactor = adj !== undefined ? adj / 25000 : 1;
	const rise = h * 0.2 * t * clamp4(adjFactor);
	return `M 0,${yBase} L ${w / 2},${yBase - rise} L ${w},${yBase}`;
}

/** Inflate bottom only. */
function inflateBottomPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const bulge = t > 0.4 ? h * 0.25 * ((t - 0.4) / 0.6) * clamp4(adjFactor) : 0;
	return `M 0,${yBase} Q ${w / 2},${yBase + bulge} ${w},${yBase}`;
}

/** Inflate top only. */
function inflateTopPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const bulge = t < 0.6 ? h * 0.25 * ((0.6 - t) / 0.6) * clamp4(adjFactor) : 0;
	return `M 0,${yBase} Q ${w / 2},${yBase - bulge} ${w},${yBase}`;
}

/** Deflate bottom only. */
function deflateBottomPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const pinch = t > 0.4 ? h * 0.2 * ((t - 0.4) / 0.6) * clamp4(adjFactor) : 0;
	return `M 0,${yBase} Q ${w / 2},${yBase - pinch} ${w},${yBase}`;
}

/** Deflate top only. */
function deflateTopPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const pinch = t < 0.6 ? h * 0.2 * ((0.6 - t) / 0.6) * clamp4(adjFactor) : 0;
	return `M 0,${yBase} Q ${w / 2},${yBase + pinch} ${w},${yBase}`;
}

// ── Priority 4 ──────────────────────────────────────────────────────────

/** Slant up: baseline rises from left to right. */
function slantUpPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 55000 : 1;
	const slant = 0.25 * clamp4(adjFactor);
	const yMid = h * (0.175 + t * 0.55);
	const yStart = yMid + (h * slant) / 2;
	const yEnd = yMid - (h * slant) / 2;
	return `M 0,${yStart} L ${w},${yEnd}`;
}

/** Slant down: baseline falls from left to right. */
function slantDownPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 55000 : 1;
	const slant = 0.25 * clamp4(adjFactor);
	const yMid = h * (0.175 + t * 0.55);
	const yStart = yMid - (h * slant) / 2;
	const yEnd = yMid + (h * slant) / 2;
	return `M 0,${yStart} L ${w},${yEnd}`;
}

/**
 * Fade right: baseline tilts down towards the right edge, hinting the
 * shrink-to-the-right taper. The squeeze magnitude no longer depends on `t`
 * (the original `(1 - 2 * t)` factor went to zero at `t=0.5`, so single-line
 * Fade Right rendered as a flat horizontal baseline with no fade at all).
 */
function fadeRightPath(w: number, h: number, t: number, adj?: number): string {
	const intensity = adj !== undefined ? clamp4(adj / 50000) : 1;
	const yLeft = h * (0.1 + t * 0.8);
	const squeeze = 0.32 * intensity;
	const yRight = h * (0.5 + squeeze * 0.4);
	return `M 0,${yLeft} L ${w},${yRight}`;
}

/** Fade left: mirror of {@link fadeRightPath}, tilts towards the left edge. */
function fadeLeftPath(w: number, h: number, t: number, adj?: number): string {
	const intensity = adj !== undefined ? clamp4(adj / 50000) : 1;
	const squeeze = 0.32 * intensity;
	const yLeft = h * (0.5 + squeeze * 0.4);
	const yRight = h * (0.1 + t * 0.8);
	return `M 0,${yLeft} L ${w},${yRight}`;
}

/** Fade up: text narrows towards the top. */
function fadeUpPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 50000 : 1;
	const taperScale = clamp4(adjFactor);
	const narrowFraction = 1 - 0.7 * taperScale;
	const narrowW = w * Math.max(0, narrowFraction);
	const lineW = narrowW + t * (w - narrowW);
	const xStart = (w - lineW) / 2;
	const yBase = h * (0.1 + t * 0.8);
	return `M ${xStart},${yBase} L ${xStart + lineW},${yBase}`;
}

/** Fade down: text narrows towards the bottom. */
function fadeDownPath(w: number, h: number, t: number, adj?: number): string {
	const adjFactor = adj !== undefined ? adj / 50000 : 1;
	const taperScale = clamp4(adjFactor);
	const narrowFraction = 1 - 0.7 * taperScale;
	const narrowW = w * Math.max(0, narrowFraction);
	const lineW = w - t * (w - narrowW);
	const xStart = (w - lineW) / 2;
	const yBase = h * (0.1 + t * 0.8);
	return `M ${xStart},${yBase} L ${xStart + lineW},${yBase}`;
}

/** Arch up pour: hollowed arch upward. */
function archUpPourPath(w: number, h: number, t: number, adj?: number): string {
	const adjNorm = adj !== undefined ? Math.max(0, Math.min(adj / 21600000, 1)) : 0.5;
	const maxArch = (0.7 * adjNorm) / 0.5;
	const archH = h * Math.max(0, maxArch - t * 0.5);
	if (archH < 1) {
		return `M 0,${h} L ${w},${h}`;
	}
	return `M 0,${h} A ${w / 2},${archH} 0 0,1 ${w},${h}`;
}

/** Arch down pour: hollowed arch downward. */
function archDownPourPath(w: number, h: number, t: number, adj?: number): string {
	const adjNorm = adj !== undefined ? Math.max(0, Math.min(adj / 21600000, 1)) : 0.5;
	const baseDepth = (0.2 * adjNorm) / 0.5;
	const archH = h * (baseDepth + t * 0.5);
	if (archH < 1) {
		return `M 0,0 L ${w},0`;
	}
	return `M 0,0 A ${w / 2},${archH} 0 0,0 ${w},0`;
}

/** Circle pour: concentric ellipses with an inner gap. */
function circlePourPath(w: number, h: number, t: number, adj?: number): string {
	const cx = w / 2;
	const cy = h / 2;
	const adjNorm = adj !== undefined ? Math.max(0, Math.min(adj / 21600000, 1)) : 0.5;
	const baseScale = 0.35 + Number(adjNorm);
	const scale = Math.min(1, baseScale) - t * 0.45;
	const rx = Math.max(1, (w / 2) * scale);
	const ry = Math.max(1, (h / 2) * scale);
	return (
		`M ${cx},${cy - ry} ` +
		`A ${rx},${ry} 0 1,1 ${cx},${cy + ry} ` +
		`A ${rx},${ry} 0 1,1 ${cx},${cy - ry}`
	);
}

/**
 * Button pour: convex top / concave bottom with larger margins. Base+extra
 * amplitude (see {@link inflatePath}) so single-line text still bows.
 */
function buttonPourPath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const intensity = adj !== undefined ? clamp4(adj / 18750) : 1;
	const bulge = h * 0.09 * intensity + h * 0.06 * intensity * (0.5 - t);
	return `M 0,${yBase} Q ${w / 2},${yBase - bulge} ${w},${yBase}`;
}

/** Deflate-inflate: pinched in centre top/bottom, expanded at edges. */
function deflateInflatePath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const adjFactor = adj !== undefined ? adj / 18750 : 1;
	const factor = Math.sin(t * Math.PI);
	const bulge = h * 0.2 * (factor - 0.5) * clamp4(adjFactor);
	return `M 0,${yBase} Q ${w / 2},${yBase - bulge} ${w},${yBase}`;
}

/**
 * Deflate-inflate-deflate: triple oscillation across the width. Base+extra
 * amplitude (see {@link inflatePath}): the original `sin(t * 2*PI)` factor is
 * zero at `t=0.5`, which made a single-line block render as a dead-flat
 * baseline (the two Bezier segments collapsed to one straight line).
 */
function deflateInflateDeflatePath(w: number, h: number, t: number, adj?: number): string {
	const yBase = h * (0.15 + t * 0.7);
	const ampScale = adj !== undefined ? clamp4(adj / 18750) : 1;
	const bulge = h * 0.15 * ampScale + h * 0.05 * ampScale * Math.sin(t * Math.PI * 2);
	const q1 = w / 3;
	const q2 = (2 * w) / 3;
	return (
		`M 0,${yBase} ` +
		`Q ${q1},${yBase - bulge} ${w / 2},${yBase} ` +
		`Q ${q2},${yBase + bulge} ${w},${yBase}`
	);
}

/**
 * Plain: a straight baseline. Only a fallback; every binding renders
 * `textPlain` through the glyph envelope, which stretches the block to the box.
 */
function plainPath(w: number, h: number, t: number): string {
	const y = h * (0.2 + t * 0.6);
	return `M 0,${y} L ${w},${y}`;
}

// ── Generator look-up table ─────────────────────────────────────────────

export const WARP_PATH_GENERATORS: Readonly<Record<string, WarpPathGenerator>> = {
	textArchUp: archUpPath,
	textArchDown: archDownPath,
	textCircle: circlePath,
	textWave1: wave1Path,
	textInflate: inflatePath,
	textDeflate: deflatePath,
	textCurveUp: curveUpPath,
	textCurveDown: curveDownPath,
	textWave2: wave2Path,
	textWave4: wave4Path,
	textDoubleWave1: doubleWave1Path,
	textCanUp: canUpPath,
	textCanDown: canDownPath,
	textButton: buttonPath,
	textRingInside: ringInsidePath,
	textRingOutside: ringOutsidePath,
	textCascadeUp: cascadeUpPath,
	textCascadeDown: cascadeDownPath,
	textTriangle: trianglePath,
	textTriangleInverted: triangleInvertedPath,
	textStop: stopPath,
	textChevron: chevronPath,
	textChevronInverted: chevronInvertedPath,
	textInflateBottom: inflateBottomPath,
	textInflateTop: inflateTopPath,
	textDeflateBottom: deflateBottomPath,
	textDeflateTop: deflateTopPath,
	textSlantUp: slantUpPath,
	textSlantDown: slantDownPath,
	textFadeRight: fadeRightPath,
	textFadeLeft: fadeLeftPath,
	textFadeUp: fadeUpPath,
	textFadeDown: fadeDownPath,
	textArchUpPour: archUpPourPath,
	textArchDownPour: archDownPourPath,
	textCirclePour: circlePourPath,
	textButtonPour: buttonPourPath,
	textDeflateInflate: deflateInflatePath,
	textDeflateInflateDeflate: deflateInflateDeflatePath,
	textPlain: plainPath,
};

/**
 * Build the SVG path `d` attribute for a warp preset at a given line position.
 *
 * @param preset    OOXML `prstTxWarp` preset name.
 * @param width     Box width in px.
 * @param height    Box height in px.
 * @param lineIndex Zero-based index of the line/paragraph.
 * @param lineCount Total number of lines/paragraphs.
 * @param adj       Optional primary adjustment (raw OOXML 1/60000th units).
 * @param adj2      Optional secondary adjustment (raw OOXML 1/60000th units).
 * @param lineGap   Arch Up, Arch Down, Button and Circle: px between the paths of two lines
 *                  (see `warpLineOutward`); 0 puts every line on the base path.
 */
export function buildWarpPath(
	preset: string,
	width: number,
	height: number,
	lineIndex: number,
	lineCount: number,
	adj?: number,
	adj2?: number,
	lineGap = 0,
): string {
	const t = lineCount <= 1 ? 0.5 : lineIndex / (lineCount - 1);
	const outward = warpLineOutward(preset, lineIndex, lineCount, lineGap);
	if (preset === 'textArchUp' || preset === 'textButton') {
		return archEllipsePath(width, height, outward, adj, 'up');
	}
	if (preset === 'textArchDown') {
		return archEllipsePath(width, height, outward, adj, 'down');
	}
	if (preset === 'textCircle') {
		return circleEllipsePath(width, height, outward, adj);
	}
	const generator = WARP_PATH_GENERATORS[preset];
	if (generator) {
		return generator(width, height, t, adj, adj2);
	}
	const y = height * (0.2 + t * 0.6);
	return `M 0,${y} L ${width},${y}`;
}
