/**
 * Tile-aware rasterisation + tiled-PDF assembly, split out of
 * `export.service.ts` to keep that file under the file-size budget.
 *
 * Every raster entry point here goes through the shared `foreignObject`
 * fidelity pipeline (`pptx-viewer-shared`'s `rasterizeElement` /
 * `rasterizeElementTiledToCanvas` / `rasterizeElementTiles`), with
 * `html2canvas-pro` kept only as the documented fallback driver. This is the
 * one place in the Angular binding that touches that driver directly.
 */
import { jsPDF } from 'jspdf';
import {
	canvasToJpegData,
	isExportIgnoredElement,
	measureExportRoot,
	placeTileOnPage,
	rasterizeElement,
	rasterizeElementTiledToCanvas,
	rasterizeElementTiles,
	rasterResultToPngBlob,
	rasterResultToPngDataUrl,
	sanitizeDownloadFilename,
	untransformExportClone,
} from 'ooxml-ui/pptx';
import type {
	RasterizeElementOptions,
	RasterOutputSize,
	RasterSourceRect,
	RasterizeElementTilesResult,
} from 'ooxml-ui/pptx';

import { renderToCanvas } from '../lib/canvas-export';
import { pdfPageSize } from './export-helpers';

/**
 * The `html2canvasFallback` driver every shared rasterize call needs, bound
 * to one element, so the html2canvas-pro adapter is written once, not
 * hand-copied per raster entry point.
 */
export function html2canvasFallbackFor(
	el: HTMLElement,
	untransformedRoot = false,
): (sourceRect: RasterSourceRect, outputSize: RasterOutputSize) => Promise<HTMLCanvasElement> {
	return (sourceRect, outputSize) =>
		renderToCanvas(el, {
			ignoreElements: isExportIgnoredElement,
			scale: outputSize.width / (sourceRect.width || 1),
			x: sourceRect.x,
			y: sourceRect.y,
			width: sourceRect.width,
			height: sourceRect.height,
			...(untransformedRoot && {
				onclone: (_doc: Document, clone: HTMLElement) => untransformExportClone(clone),
			}),
		});
}

/**
 * The size to rasterise `el` at and the matching options: a zoomed stage is drawn at its layout
 * size without its zoom (`measureExportRoot`), so the export does not follow the window.
 */
function rasterRoot(
	el: HTMLElement,
	scale: number,
): { width: number; height: number; options: RasterizeElementOptions } {
	const root = measureExportRoot(el);
	return {
		width: root.width,
		height: root.height,
		options: {
			scale,
			untransformedRoot: root.untransformedRoot,
			html2canvasFallback: html2canvasFallbackFor(el, root.untransformedRoot),
		},
	};
}

/**
 * Rasterize `el` to a PNG `Blob` via the shared `foreignObject` -> vector-SVG
 * -> html2canvas fallback chain (preserves backdrop-filter, CSS custom
 * properties and 3D transforms html2canvas cannot), tiling transparently
 * for a resolution beyond the browser's canvas cap - for PNG download and
 * "copy slide as image", whose PNG encoder can consume tiled output.
 */
export async function renderElementPngBlob(el: HTMLElement, scale: number = 2): Promise<Blob> {
	const { width, height, options } = rasterRoot(el, scale);
	const result = await rasterizeElement(el, width, height, el.ownerDocument, options);
	return rasterResultToPngBlob(result);
}

/**
 * `renderElementPngBlob`'s data-URL sibling: for a caller (Paste Special's
 * "Picture" format) that needs the pixels embedded as a data URL, not a
 * `Blob` to download.
 */
export async function renderElementPngDataUrl(el: HTMLElement, scale: number = 2): Promise<string> {
	const { width, height, options } = rasterRoot(el, scale);
	const result = await rasterizeElement(el, width, height, el.ownerDocument, options);
	return rasterResultToPngDataUrl(result);
}

/**
 * Rasterize a single element to a single full-resolution canvas via the
 * shared `foreignObject` fidelity pipeline, tiling and stitching
 * transparently (never reducing the requested scale) so the result is always
 * one canvas - for GIF/video/print, which can only consume one image per
 * frame/page. Capture each slide's canvas *while that slide is the live
 * DOM*: the viewer reuses one stage node, so a deferred capture would yield
 * the same (last) slide for every page.
 */
export async function renderElementTiled(
	el: HTMLElement,
	scale: number = 2,
): Promise<HTMLCanvasElement> {
	const { width, height, options } = rasterRoot(el, scale);
	const result = await rasterizeElementTiledToCanvas(el, width, height, el.ownerDocument, options);
	return result.canvas;
}

/**
 * Rasterize a single element to its raw per-tile canvases via the shared
 * `foreignObject` fidelity pipeline, tiling transparently when the requested
 * resolution exceeds the browser's canvas cap - for PDF pages, which have no
 * canvas-size limit of their own and so can place several small tile images
 * instead of needing one oversized canvas.
 */
export async function renderElementTilesRaster(
	el: HTMLElement,
	scale: number = 2,
): Promise<RasterizeElementTilesResult> {
	const { width, height, options } = rasterRoot(el, scale);
	return rasterizeElementTiles(el, width, height, el.ownerDocument, options);
}

/**
 * Assemble a multi-page PDF from pre-rendered per-slide tile sets (one page
 * per entry) and trigger a download. Each tile is placed at its proportional
 * position on the page (`placeTileOnPage`); a single-tile slide (the
 * overwhelming majority) degrades to exactly one full-page image per page,
 * unchanged from before tiling existed. Escapes the browser canvas cap:
 * every tile is individually small, and a PDF page has no canvas-size limit
 * of its own.
 *
 * @param pages        - One entry per slide, in order, each captured while
 *                       its slide was the live stage (`renderElementTilesRaster`).
 * @param canvasWidth  - Slide canvas width in pixels (for aspect ratio).
 * @param canvasHeight - Slide canvas height in pixels (for aspect ratio).
 * @param fileName     - Suggested download file name (unsafe chars stripped).
 */
export function buildTiledPdf(
	pages: RasterizeElementTilesResult[],
	canvasWidth: number,
	canvasHeight: number,
	fileName: string,
): void {
	if (pages.length === 0) {
		throw new Error('[ExportService] No slide pages provided for PDF export');
	}

	const { width: pageW, height: pageH, orientation } = pdfPageSize(canvasWidth, canvasHeight);
	const doc = new jsPDF({ orientation, unit: 'pt', format: [pageW, pageH] });

	for (let i = 0; i < pages.length; i++) {
		if (i > 0) {
			doc.addPage([pageW, pageH], orientation);
		}
		const { fullWidth, fullHeight, tiles } = pages[i];
		for (const tile of tiles) {
			const jpegBytes = canvasToJpegData(tile.canvas, 0.92).bytes;
			const placement = placeTileOnPage(tile, fullWidth, fullHeight, pageW, pageH);
			doc.addImage(jpegBytes, 'JPEG', placement.x, placement.y, placement.width, placement.height);
		}
	}

	doc.save(sanitizeDownloadFilename(fileName));
}
