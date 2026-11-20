/**
 * The browser's maximum canvas AREA, beside the per-axis cap `canvas-size-probe.ts` measures.
 *
 * WebKit on iOS and iPadOS limits a canvas by its pixel count, not only its width or height: a
 * canvas over 16,777,216 pixels (4096 x 4096) gets no backing store, so drawing into it silently
 * yields a blank image, even though a 16384 x 1 strip (what the per-axis probe draws) works. Probing
 * the area directly would allocate the very canvas we are trying to avoid, so the documented WebKit
 * limit is applied when the browser is iOS or iPadOS WebKit, and the square of the per-axis cap
 * elsewhere (which is what that cap already implies for every other engine).
 */

import type { WebKitEnvironment } from './webkit-engine';
import { isIosWebKit } from './webkit-engine';

/** WebKit's documented canvas area limit on iOS and iPadOS, in pixels (4096 x 4096). */
export const IOS_CANVAS_MAX_AREA = 16_777_216;

/** The largest canvas area (width x height) the browser can back, for a probed per-axis cap. */
export function maxCanvasAreaFor(
	maxDim: number,
	env: WebKitEnvironment | undefined = globalThis.navigator,
): number {
	const square = Math.max(1, Math.floor(maxDim)) ** 2;
	return isIosWebKit(env) ? Math.min(square, IOS_CANVAS_MAX_AREA) : square;
}

/**
 * Free a canvas's backing store now instead of at garbage collection. iOS counts every live
 * canvas against a total canvas memory budget and refuses new ones once it is spent, so a long
 * export (many slides or tiles) must give each canvas back as soon as its pixels are read.
 */
export function releaseCanvas(canvas: { width: number; height: number } | null | undefined): void {
	if (!canvas) return;
	try {
		canvas.width = 0;
		canvas.height = 0;
	} catch {
		// A detached or read-only stand-in: nothing to release.
	}
}
