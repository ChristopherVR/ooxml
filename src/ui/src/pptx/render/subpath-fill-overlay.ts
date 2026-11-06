/**
 * Per-sub-path FILL overlay: the shared mechanism behind two structurally
 * identical bugs.
 *
 * Both custom geometry (`a:custGeom`) and preset geometry (`a:prstGeom`) can
 * carry several sub-paths, each with its own `@fill` mode
 * (`norm`/`lighten`/`lightenLess`/`darken`/`darkenLess`/`none`) and `@stroke`
 * flag. Every binding paints a shape as ONE box: a single CSS
 * `background-color` clipped to a single merged `clip-path` built by
 * concatenating every sub-path's `d` together (`getResolvedShapeClipPath`,
 * `buildCustomGeometryClipPath`). That merge is lossy two ways:
 *
 *  - It discards each sub-path's own `@fill`, so a preset's shading/bevel
 *    sub-paths (`lighten`/`darken`, e.g. every `actionButton*`, curved arrows,
 *    `bevel`, `foldedCorner`) paint as flat instead of shaded, and a `fill="none"`
 *    sub-path (`smileyFace`'s eyes/mouth) paints FILLED instead of as an open
 *    stroke - because `clip-path` auto-closes every sub-path.
 *  - It cannot vary the fill AT ALL across sub-paths, so it is architecturally
 *    incapable of the above regardless of which colour is chosen.
 *
 * The fix is the same for both: paint the affected element as layered SVG
 * `<path>`s, each carrying its own resolved fill, instead of one CSS box. This
 * module decides WHICH elements need that (an element needs it only when at
 * least one sub-path's mode is not `norm`/unset, or it opts out of its stroke)
 * and builds the paints via the shared `./vector-subpath-paint`; a binding's
 * `ShapeEffectOverlay` renders the result as an `<svg>` sibling to the shape
 * box (mirroring `buildStrokeOutline`), and `getComputedFillStyle`
 * ({@link suppressesCssFill}) drops the CSS background so the flat colour does
 * not show underneath.
 *
 * A solid (or absent) fill is painted entirely by the overlay. A gradient,
 * pattern or picture fill keeps the single merged clip-path box (a sub-path
 * cannot carry its own paint server) and gets only the shading veils on top
 * (`buildShadingOverlay`).
 */
import type { PptxElement } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import { DEFAULT_FILL_COLOR } from '../constants';
import { silhouetteFillsOverlay } from './silhouette-effects';
import { getSubpathGeometry } from './subpath-geometry';
import type { SubpathPaint } from './vector-subpath-paint';
import { buildSubpathPaints } from './vector-subpath-paint';

/** Everything a binding needs to paint an element's per-sub-path fill overlay. */
export interface SubpathFillOverlay {
	/** One paint per sub-path, in authoring order. */
	readonly paints: readonly SubpathPaint[];
	/** `viewBox` width, in the same coordinate space as every `paints[].d`. */
	readonly viewBoxWidth: number;
	/** `viewBox` height, in the same coordinate space as every `paints[].d`. */
	readonly viewBoxHeight: number;
	/**
	 * Set when the overlay paints the fill of a shape that is not clipped (its
	 * shadow and glow are a silhouette filter, see `silhouette-effects`), as
	 * opposed to the shading sub-paths of a shape that is. A binding that paints
	 * the whole fill itself (React's custom geometry and wedge callouts) skips it.
	 */
	readonly silhouette?: boolean;
}

/**
 * Whether an element's fill mode is one this overlay can reproduce (a solid
 * colour or none at all). A gradient/pattern/image fill keeps the existing
 * single merged clip-path box, since a sub-path can only carry a solid,
 * mode-shifted colour.
 */
function hasReproducibleFillMode(element: PptxElement): boolean {
	if (!hasShapeProperties(element)) {
		return false;
	}
	const mode = element.shapeStyle?.fillMode;
	return mode === undefined || mode === 'solid' || mode === 'none';
}

/**
 * Whether a binding must drop its CSS `background-color` / `background-image`
 * for this element because {@link buildSubpathFillOverlay} is painting the
 * fill instead. Mirrors `suppressesCssBorder`'s role for the stroke overlay;
 * consumed by shared `getComputedFillStyle`.
 */
export function suppressesCssFill(element: PptxElement): boolean {
	return (
		hasReproducibleFillMode(element) &&
		getSubpathGeometry(element, silhouetteFillsOverlay(element)) !== undefined
	);
}

/** The black or white veil that shifts a fill by a sub-path's `@fill` mode. */
function shadeVeil(mode: string | undefined, opacity: number): string | undefined {
	const strength =
		mode === 'lighten' || mode === 'darken'
			? 0.4
			: mode === 'lightenLess' || mode === 'darkenLess'
				? 0.2
				: 0;
	if (strength === 0) {
		return undefined;
	}
	const channel = mode?.startsWith('lighten') ? 255 : 0;
	return `rgba(${channel},${channel},${channel},${Math.round(strength * opacity * 1000) / 1000})`;
}

/**
 * A gradient, pattern or picture fill cannot be painted per sub-path (the CSS
 * box keeps painting it), but PowerPoint still shades the `lighten`/`darken`
 * sub-paths of such a shape (a can's lid, a cube's faces, a bevel, the eyes of a
 * smiley, a folded corner): each pixel moves toward white or black by the same
 * 40% / 20% as for a solid colour. That is a translucent white or black veil
 * over just those sub-paths, painted above the unchanged fill. Returns
 * `undefined` when no sub-path shades.
 */
function buildShadingOverlay(element: PptxElement): SubpathFillOverlay | undefined {
	if (!hasShapeProperties(element)) {
		return undefined;
	}
	const ss = element.shapeStyle;
	if (ss?.fillMode !== 'gradient' && ss?.fillMode !== 'pattern' && ss?.fillMode !== 'image') {
		return undefined;
	}
	const geometry = getSubpathGeometry(element);
	if (!geometry) {
		return undefined;
	}
	const opacity = ss.fillOpacity ?? 1;
	const paints = geometry.subpaths.map((subpath) => ({
		d: subpath.d,
		fill: shadeVeil(subpath.fillMode, opacity) ?? 'none',
		stroked: subpath.stroke !== false,
	}));
	if (paints.every((paint) => paint.fill === 'none')) {
		return undefined;
	}
	return { paints, viewBoxWidth: geometry.width, viewBoxHeight: geometry.height };
}

/**
 * Build the per-sub-path fill overlay for an element, or `undefined` when a
 * single merged fill/clip-path is correct (the ordinary case).
 */
export function buildSubpathFillOverlay(element: PptxElement): SubpathFillOverlay | undefined {
	if (!hasReproducibleFillMode(element)) {
		return buildShadingOverlay(element);
	}
	const silhouette = silhouetteFillsOverlay(element);
	const geometry = getSubpathGeometry(element, silhouette);
	if (!geometry) {
		return undefined;
	}
	const ss = hasShapeProperties(element) ? element.shapeStyle : undefined;
	const hasFill =
		Boolean(ss?.fillColor) && ss?.fillColor !== 'transparent' && ss?.fillMode !== 'none';
	const fillColor = ss?.fillColor ?? DEFAULT_FILL_COLOR;
	return {
		paints: buildSubpathPaints(geometry.subpaths, hasFill, fillColor, ss?.fillOpacity),
		viewBoxWidth: geometry.width,
		viewBoxHeight: geometry.height,
		...(silhouette ? { silhouette: true } : {}),
	};
}
