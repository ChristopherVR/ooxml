/**
 * PNG slide export utilities.
 *
 * PDF and notes-PDF export live in `export-pdf.ts` (split out once this file
 * grew past the 300-LOC limit with every PDF variant included).
 */
import React from 'react';

import type { PngExportOptions, SlideCaptureOptions } from './export-helpers';
import {
	downloadBlob,
	renderElementToTiledCanvas,
	renderElementToRaster,
	rasterResultToPngBlob,
	waitForRender,
} from './export-helpers';

/* ------------------------------------------------------------------ */
/*  PNG Export                                                        */
/* ------------------------------------------------------------------ */

/**
 * Export a single slide element to a PNG Blob.
 *
 * @param slideElement - The DOM element representing the slide stage
 *                       (typically `canvasStageRef.current`).
 * @param options      - Scale, background colour, etc.
 * @returns            A PNG Blob ready for download or clipboard.
 */
export async function exportSlideToPngBlob(
	slideElement: HTMLElement,
	options: PngExportOptions = {},
): Promise<Blob> {
	const { scale = 2, backgroundColor } = options;

	const result = await renderElementToRaster(slideElement, scale, backgroundColor);
	return rasterResultToPngBlob(result);
}

/**
 * Export the current slide as a PNG and trigger a browser download.
 *
 * @param slideElement   - The slide stage DOM element.
 * @param slideIndex     - Zero-based slide index (used in filename).
 * @param options        - Scale, background colour, etc.
 */
export async function exportSlideAsPng(
	slideElement: HTMLElement,
	slideIndex: number,
	options: PngExportOptions = {},
): Promise<void> {
	const blob = await exportSlideToPngBlob(slideElement, options);
	downloadBlob(blob, `slide-${slideIndex + 1}.png`);
}

/* ------------------------------------------------------------------ */
/*  Copy slide to clipboard                                           */
/* ------------------------------------------------------------------ */

/**
 * Render the current slide as a PNG and copy it to the system clipboard.
 *
 * @param slideElement - The slide stage DOM element.
 * @param options      - Scale, background colour, etc.
 */
export async function copySlideToClipboard(
	slideElement: HTMLElement,
	options: PngExportOptions = {},
): Promise<void> {
	const blob = await exportSlideToPngBlob(slideElement, options);
	await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
}

/** One slide captured by {@link captureSlidesAsPngDataUrls}. */
export interface CapturedSlidePng {
	/** Zero-based index of the captured slide. */
	slideIndex: number;
	/** The slide as a PNG `data:` URL. */
	dataUrl: string;
}

/**
 * Capture the given slides, in the given order, as PNG data URLs.
 *
 * Each slide is shown on the live stage in turn (`setActiveSlide`, then a
 * render wait) and captured; the slide on screen before is restored at the
 * end. Goes through the shared single-canvas raster path
 * (`renderElementToTiledCanvas`, tiled and stitched past the browser canvas
 * cap rather than downscaled) rather than calling `html2canvas-pro` directly:
 * a print layout draws one image per slide, so it needs one canvas, and this
 * preserves the same `backdrop-filter`/custom-property/3D-transform fidelity
 * every other export format gets, at full resolution. A slide whose stage is
 * missing is left out of the result.
 */
export async function captureSlidesAsPngDataUrls(
	slideStageRef: React.RefObject<HTMLElement | null>,
	slideIndices: readonly number[],
	setActiveSlide: (index: number) => void,
	currentSlideIndex: number,
	options: SlideCaptureOptions = {},
): Promise<CapturedSlidePng[]> {
	const { scale = 2, onProgress } = options;
	const total = slideIndices.length;
	const captures: CapturedSlidePng[] = [];

	for (const [position, slideIndex] of slideIndices.entries()) {
		onProgress?.(position, total);
		setActiveSlide(slideIndex);
		await waitForRender(150);

		const stageEl = slideStageRef.current;
		if (!stageEl) {
			console.warn(`[export] Could not find slide stage element for slide ${slideIndex}`);
			continue;
		}

		const { canvas } = await renderElementToTiledCanvas(stageEl, scale);
		// Extract data URL immediately so the canvas can be GC'd
		captures.push({ slideIndex, dataUrl: canvas.toDataURL('image/png') });
	}

	onProgress?.(total, total);
	setActiveSlide(currentSlideIndex);
	return captures;
}

/**
 * Capture every slide of the deck as PNG data URLs (see
 * {@link captureSlidesAsPngDataUrls}, which captures only the slides asked for).
 */
export async function captureAllSlidesAsPngDataUrls(
	slideStageRef: React.RefObject<HTMLElement | null>,
	totalSlides: number,
	setActiveSlide: (index: number) => void,
	currentSlideIndex: number,
	options: SlideCaptureOptions = {},
): Promise<string[]> {
	const indices = Array.from({ length: totalSlides }, (_, i) => i);
	const captures = await captureSlidesAsPngDataUrls(
		slideStageRef,
		indices,
		setActiveSlide,
		currentSlideIndex,
		options,
	);
	return captures.map((capture) => capture.dataUrl);
}
