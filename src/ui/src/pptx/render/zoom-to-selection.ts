/**
 * Zoom to Selection: the zoom that fits the selected objects in the editing viewport, and the
 * scroll position that centres them once that zoom is drawn.
 *
 * The canvas zoom itself (`viewer-zoom-store`), the Ctrl+wheel and pinch gestures
 * (`attachViewportNavigation`) and the per-slide scroll reset (`resetViewportScroll`) are already
 * shared by every binding; this was the one zoom computation still written inside a binding
 * (React's viewport hook). Pure and DOM-free.
 *
 * @module render/zoom-to-selection
 */
import { clampZoomScale } from './zoom-step';

/** The smallest extent an object counts for (a zero-width line still has a box to fit). */
export const ZOOM_TO_SELECTION_MIN_SIZE = 12;

/** Room kept around the selection, in viewport pixels. */
export const ZOOM_TO_SELECTION_PADDING = 96;

/** An axis-aligned box in slide coordinates. */
export interface SelectionZoomBounds {
	readonly minX: number;
	readonly minY: number;
	readonly maxX: number;
	readonly maxY: number;
}

/** The box around `elements`, each at least the minimum size; null for none. */
export function selectionZoomBounds(
	elements: readonly { x: number; y: number; width: number; height: number }[],
): SelectionZoomBounds | null {
	if (elements.length === 0) {
		return null;
	}
	let minX = Number.POSITIVE_INFINITY;
	let minY = Number.POSITIVE_INFINITY;
	let maxX = Number.NEGATIVE_INFINITY;
	let maxY = Number.NEGATIVE_INFINITY;
	for (const element of elements) {
		minX = Math.min(minX, element.x);
		minY = Math.min(minY, element.y);
		maxX = Math.max(maxX, element.x + Math.max(element.width, ZOOM_TO_SELECTION_MIN_SIZE));
		maxY = Math.max(maxY, element.y + Math.max(element.height, ZOOM_TO_SELECTION_MIN_SIZE));
	}
	return [minX, minY, maxX, maxY].every(Number.isFinite) ? { minX, minY, maxX, maxY } : null;
}

/**
 * The user zoom (1 is "fit", see `viewer-zoom-store`) that fits `bounds` in a viewport of
 * `viewport` pixels, given the viewport's fit scale; clamped to the zoom range.
 */
export function zoomToSelectionScale(
	bounds: SelectionZoomBounds,
	viewport: { width: number; height: number },
	fitScale: number,
): number {
	const min = ZOOM_TO_SELECTION_MIN_SIZE;
	const boundsWidth = Math.max(bounds.maxX - bounds.minX, min);
	const boundsHeight = Math.max(bounds.maxY - bounds.minY, min);
	const availableWidth = Math.max(viewport.width - ZOOM_TO_SELECTION_PADDING, min);
	const availableHeight = Math.max(viewport.height - ZOOM_TO_SELECTION_PADDING, min);
	const target = Math.min(availableWidth / boundsWidth, availableHeight / boundsHeight);
	return clampZoomScale(target / Math.max(fitScale, Number.EPSILON));
}

/**
 * The scroll position that centres `bounds` in the scrolling wrapper once the slide is drawn at
 * `editorScale` (fit scale times user zoom), with the slide's canvas at `canvasOffset` inside it.
 */
export function selectionScrollTarget(
	bounds: SelectionZoomBounds,
	editorScale: number,
	canvasOffset: { left: number; top: number },
	wrapper: { width: number; height: number },
): { left: number; top: number } {
	const centerX = (bounds.minX + bounds.maxX) / 2;
	const centerY = (bounds.minY + bounds.maxY) / 2;
	return {
		left: Math.max(canvasOffset.left + centerX * editorScale - wrapper.width / 2, 0),
		top: Math.max(canvasOffset.top + centerY * editorScale - wrapper.height / 2, 0),
	};
}
