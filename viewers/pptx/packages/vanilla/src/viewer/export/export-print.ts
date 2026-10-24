import type { CanvasSize, PrintJobCapture, PrintJobPlan, PrintSettings } from 'ooxml-ui/pptx';
import {
	buildOutlinePrintJobDocument,
	buildRasterPrintJobDocument,
	buildSaveSlides,
	buildVectorPrintJobDocument,
	exportAbortError,
	finishPrintWindow as finishPendingWindow,
	openPendingPrintWindow as openPendingWindow,
	openPrintWindow as defaultOpenPrintWindow,
	planPrintJob,
	printJobRasterFallback,
} from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

import { createTranslator } from '../i18n';
import { exportSlideToSvg } from './export-svg';
import type { ExportCaptureDeps, ExportProgress } from './export-types';

/**
 * Print for the vanilla binding. The shared `planPrintJob` decides what prints
 * (range, current slide, hidden slides), through which path and under which
 * window title, the shared `build*PrintJobDocument` functions assemble the
 * DOMPurify-hardened document, and the shared print-window helpers keep the
 * window clear of popup blocking. Only the capture and the opener live here:
 * full-page slides go through the core SVG exporter (captured instead when it
 * throws), notes pages and handouts are rasterised slide by slide.
 */

/**
 * Open a print window for a complete HTML document and trigger printing.
 * Returns `false` when the window could not be opened (popup blocker).
 */
export type OpenPrintWindow = (htmlDocument: string) => boolean;

/**
 * Options for `print`: any subset of the shared `PrintSettings` (unspecified
 * fields fall back to `DEFAULT_PRINT_SETTINGS`, i.e. all slides, landscape,
 * full colour) plus progress/abort and a print-window override.
 */
export interface PrintOptions extends Partial<PrintSettings> {
	/** Rasterisation progress callback: `(currentSlide, totalSlidesToPrint)`. */
	onProgress?: ExportProgress;
	/** Abort before the window opens; checked between slide captures. */
	signal?: AbortSignal;
	/**
	 * Override how the assembled document is opened (e.g. write it into a
	 * hidden iframe). Popup-blocker caveat: the default opener uses
	 * `window.open`, which browsers typically only allow inside a user
	 * gesture, so call `print()` from a click handler; when the popup is
	 * blocked the returned promise resolves `false`.
	 */
	openPrintWindow?: OpenPrintWindow;
}

/**
 * Assemble the printable HTML document for the (validated) settings and open
 * it in a print window. Resolves `true` when the window opened, `false` when
 * it was blocked or nothing matched the slide range. Throws the shared
 * `AbortError` when `signal` aborts between slide captures.
 */
export async function runPrint(
	deps: ExportCaptureDeps,
	options: PrintOptions = {},
): Promise<boolean> {
	const { onProgress, signal, openPrintWindow, ...settings } = options;
	const openWindow = openPrintWindow ?? defaultOpenPrintWindow;
	const state = deps.store.get();
	// The store keeps layout and master elements apart from each slide; the
	// vector export and the outline need the complete slides.
	const slides = buildSaveSlides(state.slides, state.templateElementsBySlideId);
	const plan = planPrintJob(settings, {
		slides,
		activeSlideIndex: state.currentSlide,
		includeHiddenSlides: deps.getIncludeHiddenSlides?.() ?? false,
	});
	if (!plan.ok) {
		return false;
	}
	const t = deps.getTranslator?.() ?? createTranslator();
	const title = (job: PrintJobPlan): string => t(job.title.key, job.title.params);

	// Outline is text-only: no rasterisation needed.
	if (plan.output === 'outline') {
		return openWindow(buildOutlinePrintJobDocument(plan, title(plan), slides));
	}

	if (plan.output === 'vector') {
		const html = vectorDocument(plan, title(plan), slides, state.canvasSize, options);
		if (html !== undefined) {
			return openWindow(html);
		}
	}

	// Rasterising each slide below `await`s, so the default opener must grab
	// the window NOW, before that first await, or it gets popup-blocked. A
	// custom `openPrintWindow` owns its own timing and gets the finished
	// document in a single call, as documented.
	const job = printJobRasterFallback(plan);
	const pendingWindow = openPrintWindow
		? undefined
		: openPendingWindow(t('pptx.print.preparingToPrint'));
	if (!openPrintWindow && !pendingWindow) {
		return false;
	}

	// Options > Advanced > "High quality" raster scale for the capture path,
	// composed on top of the host's own baseline (2x * Options > Advanced >
	// Image Size/Quality) scale.
	const printScaleMultiplier = deps.getPrintHighQuality?.() ? 2 : 1;
	const total = job.slideIndices.length;
	const captures: PrintJobCapture[] = [];
	for (const [position, slideIndex] of job.slideIndices.entries()) {
		if (signal?.aborted) {
			pendingWindow?.close();
			throw exportAbortError();
		}
		onProgress?.(position, total);
		const canvas = await deps.rasterizeSlide(slideIndex, printScaleMultiplier);
		captures.push({ slideIndex, dataUrl: canvas.toDataURL('image/png') });
	}
	onProgress?.(total, total);

	const html = buildRasterPrintJobDocument(job, title(job), captures, slides, state.handoutMaster);
	if (pendingWindow) {
		finishPendingWindow(pendingWindow, html);
		return true;
	}
	return openWindow(html);
}

/**
 * The vector document of a `vector` plan, or `undefined` when the SVG export
 * threw. Reports progress per exported slide and honours an abort between them.
 */
function vectorDocument(
	plan: PrintJobPlan,
	title: string,
	slides: PptxSlide[],
	size: CanvasSize,
	{ onProgress, signal }: Pick<PrintOptions, 'onProgress' | 'signal'>,
): string | undefined {
	const total = plan.slideIndices.length;
	const svgs: string[] = [];
	try {
		for (const [position, slideIndex] of plan.slideIndices.entries()) {
			if (signal?.aborted) {
				throw exportAbortError();
			}
			onProgress?.(position, total);
			svgs.push(exportSlideToSvg(slides[slideIndex], size.width, size.height));
		}
	} catch (err) {
		if (signal?.aborted) {
			throw err;
		}
		console.warn('[PowerPointViewer] SVG print failed, falling back to raster:', err);
		return undefined;
	}
	onProgress?.(total, total);
	return buildVectorPrintJobDocument(plan, title, svgs, size);
}
