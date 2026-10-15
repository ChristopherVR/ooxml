/** Smallest zoom the automatic fit may choose; the status bar slider reaches it as well. */
export const AUTO_FIT_MIN = 25;

/**
 * The zoom Word for the web and mobile apply on their own: a page wider than the space the canvas
 * offers is scaled to that width, anything else stays at 100%. Returns undefined when the canvas has
 * not been laid out yet (zero width), so callers keep the current zoom.
 */
export function autoFitZoom(availableWidth: number, pageWidth: number): number | undefined {
	if (!(availableWidth > 0) || !(pageWidth > 0)) return undefined;
	if (availableWidth >= pageWidth) return 100;
	return Math.max(AUTO_FIT_MIN, Math.floor((availableWidth / pageWidth) * 100));
}

/** Pinch scale applied to the zoom at the gesture start, rounded and kept inside the slider range. */
export function pinchZoom(startPercent: number, scale: number, min: number, max: number): number {
	if (!(scale > 0)) return startPercent;
	return Math.min(max, Math.max(min, Math.round(startPercent * scale)));
}
