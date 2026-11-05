/**
 * The per-sub-path geometry of a shape: custom geometry with structured paths,
 * or a preset's evaluated sub-paths. Shared by the per-sub-path fill overlay
 * (`subpath-fill-overlay`) and the silhouette effects (`silhouette-effects`),
 * which paint a shape's fill as SVG paths when a CSS `clip-path` cannot.
 *
 * @module render/subpath-geometry
 */
import type {
	CustomGeometryPath,
	CustomGeometryRawData,
	CustomGeometrySubpathSvg,
	PptxElement,
} from 'ooxml-core/pptx';
import {
	customGeometryPathsToSvgSubpaths,
	evaluateCustomGeometryPaths,
	evaluatePresetShape,
} from 'ooxml-core/pptx';

/**
 * The box a preset is evaluated at: the element's authored extent, floored at
 * zero. Matches `getContainerStyle`'s `paintedElementSize` (not imported
 * directly, to avoid a `fill-style.ts` <-> `element-style.ts` <->
 * `subpath-fill-overlay.ts` import cycle: `fill-style.ts` consumes
 * {@link suppressesCssFill} and `element-style.ts` consumes `fill-style.ts`).
 * This used to pad out to `MIN_ELEMENT_SIZE`, which stretched a degenerate
 * shape's per-sub-path fill overlay past its authored extent in read-only
 * rendering (issue #285); any interaction-only padding now lives solely in
 * `elementHitTargetStyle`, which never touches paint.
 */
export function paintedSize(element: PptxElement): { width: number; height: number } {
	return {
		width: Math.max(element.width, 0),
		height: Math.max(element.height, 0),
	};
}

/**
 * Whether any sub-path needs individual paint handling: a non-`norm` fill mode
 * or an explicit stroke opt-out. Mirrors the threshold the React custom-geometry
 * renderer already used, now shared with preset geometry too. A shape whose
 * sub-paths are all plain (unset/`norm` fill, stroked) renders identically
 * either way, so declining here keeps ordinary shapes on the cheaper CSS path.
 */
function needsPerSubpathPaint(subpaths: readonly CustomGeometrySubpathSvg[]): boolean {
	return subpaths.some(
		(sp) => (sp.fillMode !== undefined && sp.fillMode !== 'norm') || sp.stroke === false,
	);
}

/** The custom-geometry / preset slots read off an element. */
interface SubpathCapableElement {
	pathData?: string;
	pathWidth?: number;
	pathHeight?: number;
	customGeometryPaths?: CustomGeometryPath[];
	customGeometryRawData?: CustomGeometryRawData;
	shapeType?: string;
	shapeAdjustments?: Record<string, number>;
}

/** Resolved per-sub-path geometry needing individual paint, or `undefined`. */
export interface SubpathGeometry {
	subpaths: CustomGeometrySubpathSvg[];
	width: number;
	height: number;
}

/** Sub-path-capable element types: `PptxCustomPathProperties` lives only on these. */
function asSubpathCapable(element: PptxElement): SubpathCapableElement | undefined {
	if (element.type === 'shape' || element.type === 'image' || element.type === 'picture') {
		return element;
	}
	return undefined;
}

/**
 * Resolve an element's per-sub-path geometry when, and only when, it genuinely
 * needs individual paint handling - custom geometry with structured sub-paths,
 * or a preset whose evaluated sub-paths vary. Returns `undefined` for every
 * ordinary shape so callers keep their existing single merged fill/clip.
 */
export function getSubpathGeometry(
	element: PptxElement,
	force = false,
): SubpathGeometry | undefined {
	const slots = asSubpathCapable(element);
	if (!slots) {
		return undefined;
	}

	// Custom geometry: only the STRUCTURED sub-paths (from `a:custGeom/a:pathLst`)
	// carry per-sub-path `@fill`/`@stroke`; a freeform with only aggregate
	// `pathData` (no structured paths) has no per-sub-path intent to honour.
	if (slots.pathData && slots.pathWidth && slots.pathHeight) {
		// Live-evaluate the raw `a:pathLst` against the current adjust values so a
		// mid-drag `a:ahXY`/`a:ahPolar` reshape moves the per-sub-path overlay
		// with the clip path; falls back to the parse-time paths when the raw
		// XML is missing (older data) or the guide list cannot be evaluated.
		const structured =
			evaluateCustomGeometryPaths(
				slots.customGeometryRawData,
				slots.pathWidth,
				slots.pathHeight,
				slots.shapeAdjustments,
			) ?? slots.customGeometryPaths;
		if (!structured || structured.length === 0) {
			return undefined;
		}
		const subpaths = customGeometryPathsToSvgSubpaths(
			structured,
			slots.pathWidth,
			slots.pathHeight,
		);
		return force || needsPerSubpathPaint(subpaths)
			? { subpaths, width: slots.pathWidth, height: slots.pathHeight }
			: undefined;
	}

	// Preset geometry: evaluate at the same PAINTED size the container box uses
	// (padded for degenerate shapes), so the overlay's viewBox lines up with it.
	if (!slots.shapeType) {
		return undefined;
	}
	const { width, height } = paintedSize(element);
	const result = evaluatePresetShape(slots.shapeType, width, height, slots.shapeAdjustments);
	// `fillNone` (every sub-path opts out of fill, e.g. `arc`) is an OPEN preset:
	// `./stroke-only-preset` already routes it through the stroke overlay with no
	// fill at all, so it must not also be treated as a multi-sub-path FILL here.
	if (!result || result.fillNone || result.paths.length === 0) {
		return undefined;
	}
	const subpaths: CustomGeometrySubpathSvg[] = result.paths.map((p) => ({
		d: p.d,
		fillMode: p.fill as CustomGeometrySubpathSvg['fillMode'],
		stroke: p.stroke,
	}));
	return force || needsPerSubpathPaint(subpaths) ? { subpaths, width, height } : undefined;
}
