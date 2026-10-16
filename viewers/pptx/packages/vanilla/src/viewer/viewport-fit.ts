import {
	calculateViewportFit,
	editorViewportFitFor,
	resolveViewportFitOptions,
} from 'ooxml-ui/pptx';
import type { CanvasSize, ViewportFitOptions } from 'ooxml-ui/pptx';

/** Apply explicit host padding to both the CSS box and its fit measurement. */
export function fitViewerViewport(
	viewport: HTMLElement,
	canvasSize: CanvasSize,
	options: ViewportFitOptions,
	presenting: boolean,
): number {
	const policy = presenting ? { fitPadding: 0, maxFitScale: null } : options;
	const defaults = editorViewportFitFor(viewport.clientWidth);
	{
		const { horizontal, vertical } = resolveViewportFitOptions(policy, defaults).fitPadding;
		viewport.style.padding = `${vertical}px ${horizontal}px`;
	}
	return calculateViewportFit(
		{
			viewportWidth: viewport.clientWidth,
			viewportHeight: viewport.clientHeight,
			canvasWidth: Math.max(canvasSize.width, 1),
			canvasHeight: Math.max(canvasSize.height, 1),
			...policy,
		},
		defaults,
	).scale;
}
