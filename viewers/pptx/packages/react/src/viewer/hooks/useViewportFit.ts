import {
	calculateViewportFit,
	editorViewportFitFor,
	observeViewportClientSize,
	resolveViewportFitOptions,
} from 'ooxml-ui/pptx';
import type { ViewportFitOptions } from 'ooxml-ui/pptx';
import { useEffect, useLayoutEffect, useState } from 'react';
import type { RefObject } from 'react';

import type { CanvasSize } from '../types';

/**
 * Measure before paint in the browser, so the first painted frame is already fitted instead
 * of showing the authored size for a frame. Neither effect runs on a server; `useEffect` there
 * keeps React 18 from warning about a layout effect during server rendering.
 */
const useBeforePaintEffect = typeof window === 'undefined' ? useEffect : useLayoutEffect;

/** Keep measured fit and available selection area on the same host layout policy. */
export function useViewportFit(
	canvasSize: CanvasSize,
	viewportRef: RefObject<HTMLDivElement | null>,
	options: ViewportFitOptions,
	viewportNode?: HTMLDivElement | null,
) {
	const [editorDimensions, setEditorDimensions] = useState<CanvasSize | null>(null);
	const { maxFitScale } = resolveViewportFitOptions(options, editorViewportFitFor(Infinity));
	const hostPadding = options.fitPadding;
	const effectiveEditorDimensions = editorDimensions ?? canvasSize;
	const { scale: fitScale } = calculateViewportFit({
		viewportWidth: effectiveEditorDimensions.width,
		viewportHeight: effectiveEditorDimensions.height,
		canvasWidth: canvasSize.width,
		canvasHeight: canvasSize.height,
		maxFitScale,
	});

	// Preserve the stock decorative allowance unless the host supplies one.
	// Rulers and the host's actual toolbar/sidebar footprint are separate layout.
	useBeforePaintEffect(() => {
		if (viewportNode === null) {
			return;
		}
		let stopObserving: (() => void) | null = null;
		let raf = 0;
		const measure = (client: { width: number; height: number }) => {
			const { availableWidth: width, availableHeight: height } = calculateViewportFit({
				viewportWidth: client.width,
				viewportHeight: client.height,
				canvasWidth: canvasSize.width,
				canvasHeight: canvasSize.height,
				// The host's padding when it gave one, else the width-aware editor default.
				...resolveViewportFitOptions(
					{ fitPadding: hostPadding },
					editorViewportFitFor(client.width),
				),
			});
			// A hidden or not-yet-sized viewport must not reset a valid measurement.
			if (width > 0 && height > 0) {
				setEditorDimensions({ width, height });
			}
		};
		const attach = () => {
			const element = viewportNode ?? viewportRef.current;
			if (!element) {
				// Only an object-ref caller can get here: its node is assigned after this commit.
				raf = requestAnimationFrame(attach);
				return;
			}
			// The fractional client size, not the whole-pixel clientWidth (which rounds up by up
			// to half a pixel), reported now and on every change, padding-only ones included.
			stopObserving = observeViewportClientSize(element, measure);
		};
		attach();
		return () => {
			cancelAnimationFrame(raf);
			stopObserving?.();
		};
	}, [viewportRef, viewportNode, canvasSize.width, canvasSize.height, hostPadding]);

	return { editorDimensions, setEditorDimensions, effectiveEditorDimensions, fitScale };
}
