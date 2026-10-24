import type { CanvasSize, PrintJobCapture, PrintJobPlan, PrintSettings } from 'ooxml-ui/pptx';
import {
	buildOutlinePrintJobDocument,
	buildRasterPrintJobDocument,
	buildVectorPrintJobDocument,
	planPrintJob,
	printJobRasterFallback,
} from 'ooxml-ui/pptx';
import type { PptxHandoutMaster, PptxSlide } from 'pptx-viewer-core';

import { createTranslator } from '../../i18n/translator';
import type { Translator } from '../../i18n/translator';
import type { RasterizeSlide } from './export-controller.svelte';
import { exportSlideToSvg } from './export-svg';

/**
 * Print flow: the shared `planPrintJob` decides what prints (range, current
 * slide, hidden slides), through which path and under which window title, and
 * the shared `build*PrintJobDocument` functions assemble the DOMPurify-hardened
 * document. This module (the Svelte counterpart of Vue's `usePrint`) keeps
 * only its capture and its opener: it exports direct slide pages as vector SVG
 * (captured instead when the export throws), rasterises notes/handout
 * thumbnails, and opens the print surface.
 *
 * **Print surface / popup-blocker caveats:** the default opener renders the
 * document into a hidden same-origin `<iframe srcdoc>` and calls
 * `contentWindow.print()`, so no popup window is involved and popup blockers
 * cannot interfere (`document.write` is avoided too). A host that prefers a
 * visible print window can inject its own {@link OpenPrintWindow} via the
 * controller deps; note that a `window.open`-based opener only succeeds inside
 * a user gesture (a click handler) and returns `null` under popup blockers,
 * in which case it should report `false` so the flow resolves `false` and the
 * host can surface "allow popups for this site" guidance.
 */

/** Options for the print flow: any subset of the shared print settings. */
export type PrintOptions = Partial<PrintSettings>;

/**
 * Render a complete HTML document into a print surface and trigger printing.
 * Returns `false` when the surface could not be opened (e.g. a custom
 * `window.open`-based opener hit a popup blocker).
 */
export type OpenPrintWindow = (htmlDocument: string) => boolean;

/** Injected dependencies (kept DOM-light for unit tests). */
export interface PrintDeps {
	/** Live slide list; read fresh on every call. */
	getSlides(): PptxSlide[];
	/** Active slide index (0-based); target of the `current` slide range. */
	getCurrent(): number;
	/** Live slide dimensions used by the core SVG exporter. */
	getCanvasSize(): CanvasSize;
	rasterizeSlide: RasterizeSlide;
	/** Print-surface opener override (test seam / host popup handling). */
	openPrintWindow?: OpenPrintWindow;
	/** Options > Advanced > "Print hidden slides". Defaults to `false` (excluded), matching PowerPoint. */
	getIncludeHiddenSlides?(): boolean;
	/** Options > Advanced > "High quality" raster scale for the print fallback path. */
	getPrintHighQuality?(): boolean;
	/**
	 * The deck's handout master, consumed for its background, header/footer/
	 * date/page-number placeholders, and (when authored) positioned slide
	 * rects, in the `handouts` print mode.
	 */
	getHandoutMaster?(): PptxHandoutMaster | undefined;
	/** Live translator for the print window title. English when omitted. */
	getTranslator?(): Translator;
}

/** How long the hidden print iframe survives if `afterprint` never fires. */
const PRINT_FRAME_FALLBACK_TTL_MS = 60_000;

/**
 * Default opener: mount the document in a hidden same-origin iframe (via
 * `srcdoc`, no `document.write`) and print its content window. The frame is
 * removed after `afterprint`, with a TTL fallback for browsers that never
 * fire it on iframe windows.
 */
export function defaultOpenPrintWindow(htmlDocument: string): boolean {
	const frame = document.createElement('iframe');
	frame.setAttribute('aria-hidden', 'true');
	Object.assign(frame.style, {
		position: 'fixed',
		right: '0',
		bottom: '0',
		width: '0',
		height: '0',
		border: '0',
		visibility: 'hidden',
	});
	const cleanup = (): void => {
		setTimeout(() => frame.remove(), 500);
	};
	frame.addEventListener('load', () => {
		const win = frame.contentWindow;
		if (!win) {
			frame.remove();
			return;
		}
		win.addEventListener('afterprint', cleanup, { once: true });
		win.focus();
		win.print();
	});
	frame.srcdoc = htmlDocument;
	document.body.appendChild(frame);
	setTimeout(() => frame.remove(), PRINT_FRAME_FALLBACK_TTL_MS);
	return true;
}

/** The vector document of a `vector` plan, or `undefined` when the SVG export threw. */
function vectorDocument(
	plan: PrintJobPlan,
	title: string,
	slides: PptxSlide[],
	size: CanvasSize,
): string | undefined {
	try {
		const svgs = plan.slideIndices.map((index) =>
			exportSlideToSvg(slides[index] as PptxSlide, size.width, size.height),
		);
		return buildVectorPrintJobDocument(plan, title, svgs, size);
	} catch (err) {
		console.warn('[PowerPointViewer] SVG print failed, falling back to raster:', err);
		return undefined;
	}
}

/**
 * Run the print flow for the given (partial) settings. Resolves `true` when
 * the print surface opened, `false` when there was nothing to print or the
 * opener reported failure (see the popup-blocker caveats above).
 */
export async function printSlides(deps: PrintDeps, options: PrintOptions = {}): Promise<boolean> {
	const slides = deps.getSlides();
	const plan = planPrintJob(options, {
		slides,
		activeSlideIndex: deps.getCurrent(),
		includeHiddenSlides: deps.getIncludeHiddenSlides?.() ?? false,
	});
	if (!plan.ok) {
		return false;
	}
	const openWindow = deps.openPrintWindow ?? defaultOpenPrintWindow;
	const t = deps.getTranslator?.() ?? createTranslator(() => 'en');
	const title = (job: PrintJobPlan): string => t(job.title.key, job.title.params);

	// Outline is text-only: no rasterisation needed.
	if (plan.output === 'outline') {
		return openWindow(buildOutlinePrintJobDocument(plan, title(plan), slides));
	}

	// Full-page slides use the core data-model SVG exporter. This preserves
	// rich vector content such as chart marks and avoids DOM raster capture.
	if (plan.output === 'vector') {
		const html = vectorDocument(plan, title(plan), slides, deps.getCanvasSize());
		if (html !== undefined) {
			return openWindow(html);
		}
	}

	// Options > Advanced > "High quality" raster scale for the capture path,
	// composed on top of the host's own baseline (2x * Options > Advanced >
	// Image Size/Quality) scale.
	const job = printJobRasterFallback(plan);
	const printScaleMultiplier = deps.getPrintHighQuality?.() ? 2 : 1;
	const captures: PrintJobCapture[] = [];
	for (const slideIndex of job.slideIndices) {
		const canvas = await deps.rasterizeSlide(slideIndex, printScaleMultiplier);
		captures.push({ slideIndex, dataUrl: canvas.toDataURL('image/png') });
	}
	return openWindow(
		buildRasterPrintJobDocument(job, title(job), captures, slides, deps.getHandoutMaster?.()),
	);
}
