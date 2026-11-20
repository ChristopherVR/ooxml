/**
 * `rasterizeElement()`, but the requested scale is clamped so the result is
 * always a single canvas (`kind: 'canvas'`) - never tiled `png-bytes`.
 *
 * For a caller that genuinely cannot consume tiled output (JPEG has no way
 * to concatenate tiles the way PNG's row-band encoder does; the notes-PDF
 * layout draws exactly one image into a fixed page region alongside
 * wrapped notes text, not a tile grid), downscaling above the cap is the
 * documented, honest fallback: fidelity (`foreignObject`, preserving
 * `backdrop-filter`/custom properties/3D transforms) is still preserved,
 * only the requested resolution is not, and only once the natural size
 * already exceeds what a single canvas can hold at the requested scale (in
 * practice: a source many thousands of CSS pixels wide, far beyond any real
 * slide).
 */
import { maxCanvasAreaFor } from '../render/canvas-area-cap';
import { getMaxCanvasDimension } from '../render/canvas-size-probe';
import type { RasterizeElementCanvasResult, RasterizeElementOptions } from './rasterize-element';
import { rasterizeElement } from './rasterize-element';

/** Same result as the common (non-tiled) `rasterizeElement()` case, plus whether the scale had to be reduced. */
export interface RasterizeElementClampedResult extends RasterizeElementCanvasResult {
	/** `true` when `options.scale` was reduced to fit the canvas cap. */
	clamped: boolean;
	/** The scale actually used (equals `options.scale` unless `clamped`). */
	effectiveScale: number;
}

/** The tile planner rounds to whole pixels; this is the size it will plan for `scale`. */
function plannedPixels(natural: number, scale: number): number {
	return Math.max(1, Math.round(natural * scale));
}

/**
 * The largest scale <= `requestedScale` at which a `naturalWidth` x `naturalHeight` raster fits one
 * canvas: no edge over `maxDim` and no more than `maxArea` pixels (iOS caps the area at 4096 x 4096
 * even where a 16384-pixel edge works). Both axes shrink by the same factor, so the aspect ratio holds.
 */
export function clampScaleToCanvas(
	naturalWidth: number,
	naturalHeight: number,
	requestedScale: number,
	maxDim: number,
	maxArea: number = Number.POSITIVE_INFINITY,
): number {
	const fullWidth = naturalWidth * requestedScale;
	const fullHeight = naturalHeight * requestedScale;
	const largestEdge = Math.max(fullWidth, fullHeight);
	const edgeFactor = largestEdge > maxDim ? maxDim / largestEdge : 1;
	const area = fullWidth * fullHeight;
	const areaFactor = area > maxArea ? Math.sqrt(maxArea / area) : 1;
	let scale = requestedScale * Math.min(edgeFactor, areaFactor);
	// Rounding to whole pixels can push the planned raster one pixel over a cap; step down until
	// the planner keeps it in one canvas.
	for (let step = 0; step < 64; step++) {
		const width = plannedPixels(naturalWidth, scale);
		const height = plannedPixels(naturalHeight, scale);
		if (width <= maxDim && height <= maxDim && width * height <= maxArea) break;
		scale *= 0.999;
	}
	return scale;
}

/**
 * Rasterise `element`, reducing `options.scale` (preserving aspect ratio) so
 * the full raster never exceeds the browser's canvas-dimension or area cap,
 * and therefore never tiles.
 */
export async function rasterizeElementClampedToCanvas(
	element: HTMLElement,
	naturalWidth: number,
	naturalHeight: number,
	doc: Document,
	options: RasterizeElementOptions,
): Promise<RasterizeElementClampedResult> {
	const requestedScale = options.scale && options.scale > 0 ? options.scale : 1;
	const maxDim = options.maxCanvasDim ?? getMaxCanvasDimension();
	const maxArea = options.maxCanvasArea ?? maxCanvasAreaFor(maxDim);
	const effectiveScale = clampScaleToCanvas(
		naturalWidth,
		naturalHeight,
		requestedScale,
		maxDim,
		maxArea,
	);

	const result = await rasterizeElement(element, naturalWidth, naturalHeight, doc, {
		...options,
		scale: effectiveScale,
		maxCanvasDim: maxDim,
		maxCanvasArea: maxArea,
	});

	if (result.kind !== 'canvas') {
		// Should not happen (the clamp above guarantees the full raster fits
		// under `maxDim`), but a defensive, honest failure beats a silent
		// `png-bytes` value the caller does not expect.
		throw new Error(
			'rasterizeElementClampedToCanvas: rasterizeElement still tiled after clamping the scale',
		);
	}

	return { ...result, clamped: effectiveScale < requestedScale, effectiveScale };
}
