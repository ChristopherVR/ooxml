import { MOBILE_BREAKPOINT } from '../mobile-viewport';
import type { ViewportFitOptions } from '../viewport-fit';

/** Ordinary editor layout, based on the React binding. Hosts may override fit. */
export const EDITOR_VIEWPORT_FIT: Required<ViewportFitOptions> = {
	fitPadding: { horizontal: 4, vertical: 16 },
	maxFitScale: 1,
};

/**
 * The editor fit policy for a viewport of this width. A phone insets the slide 12px from each
 * edge, as PowerPoint mobile does, instead of the desktop 4px that read as "touching the
 * edge". Pass the result as the `defaults` of `calculateViewportFit` / `resolveViewportFitOptions`;
 * a host's own `fitPadding` still wins over it.
 */
export function editorViewportFitFor(viewportWidth: number): Required<ViewportFitOptions> {
	return viewportWidth > 0 && viewportWidth < MOBILE_BREAKPOINT
		? { ...EDITOR_VIEWPORT_FIT, fitPadding: { horizontal: 12, vertical: 12 } }
		: EDITOR_VIEWPORT_FIT;
}

export const EDITOR_SLIDE_RAIL_WIDTH = 180;
export const EDITOR_THUMBNAIL_WIDTH = 132;
export const EDITOR_THUMBNAIL_ROW_CHROME = 6;
export const EDITOR_THUMBNAIL_GAP = 4;
export const EDITOR_THUMBNAIL_NUMBER_HEIGHT = 15;

/** Reserve the border, number column, gap and padding before sizing the preview. */
export function editorThumbnailWidth(railWidth = EDITOR_SLIDE_RAIL_WIDTH): number {
	return Math.max(1, railWidth - (EDITOR_SLIDE_RAIL_WIDTH - EDITOR_THUMBNAIL_WIDTH));
}

export function editorThumbnailHeight(
	canvasWidth: number,
	canvasHeight: number,
	width = EDITOR_THUMBNAIL_WIDTH,
): number {
	return Math.max(1, canvasHeight) * (width / Math.max(1, canvasWidth));
}

export function editorThumbnailStep(
	canvasWidth: number,
	canvasHeight: number,
	width = EDITOR_THUMBNAIL_WIDTH,
): number {
	// A very wide slide can be shorter than its number's line box. Include the
	// taller content, two 2px row paddings and the gap in the virtualized step.
	const contentHeight = Math.max(
		editorThumbnailHeight(canvasWidth, canvasHeight, width) + 2,
		EDITOR_THUMBNAIL_NUMBER_HEIGHT,
	);
	return contentHeight + (EDITOR_THUMBNAIL_ROW_CHROME - 2) + EDITOR_THUMBNAIL_GAP;
}
