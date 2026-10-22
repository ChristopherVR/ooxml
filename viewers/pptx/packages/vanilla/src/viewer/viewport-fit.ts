import {
	calculateViewportFit,
	editorViewportFitFor,
	measureViewportClientSize,
	resolveViewportFitOptions,
} from 'ooxml-ui/pptx';
import type { CanvasSize, ViewportClientSize, ViewportFitOptions } from 'ooxml-ui/pptx';

/**
 * The last fractional measurement of a viewport, reused until its owner clears it (a stage
 * render, a viewport resize). The fit scale is read on every pointer event and selection
 * change, and a full measurement (computed style and bounding rect) on each of those slowed a
 * tap enough to miss the double-tap window.
 */
export interface ViewportMeasureCache {
	/** The viewport measured; another element (a rebuilt chrome) starts afresh. */
	viewport?: HTMLElement | undefined;
	/** The padding last written to it; writing a different one discards `client`. */
	padding?: string | undefined;
	client?: ViewportClientSize | undefined;
}

/** Apply explicit host padding to both the CSS box and its fit measurement. */
export function fitViewerViewport(
	viewport: HTMLElement,
	canvasSize: CanvasSize,
	options: ViewportFitOptions,
	presenting: boolean,
	cache: ViewportMeasureCache = {},
): number {
	const policy = presenting ? { fitPadding: 0, maxFitScale: null } : options;
	const defaults = editorViewportFitFor(viewport.clientWidth);
	{
		const { horizontal, vertical } = resolveViewportFitOptions(policy, defaults).fitPadding;
		const padding = `${vertical}px ${horizontal}px`;
		if (cache.viewport !== viewport || cache.padding !== padding) {
			viewport.style.padding = padding;
			cache.viewport = viewport;
			cache.padding = padding;
			cache.client = undefined;
		}
	}
	// Fractional, not clientWidth: a whole-pixel size rounds up by up to half a pixel and the
	// fitted slide would overflow the viewport by that much.
	const client = (cache.client ??= measureViewportClientSize(viewport));
	return calculateViewportFit(
		{
			viewportWidth: client.width,
			viewportHeight: client.height,
			canvasWidth: Math.max(canvasSize.width, 1),
			canvasHeight: Math.max(canvasSize.height, 1),
			...policy,
		},
		defaults,
	).scale;
}
