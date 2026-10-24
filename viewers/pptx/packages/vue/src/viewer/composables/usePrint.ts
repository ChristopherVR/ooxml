import {
	buildOutlinePrintJobDocument,
	buildRasterPrintJobDocument,
	buildVectorPrintJobDocument,
	DEFAULT_VIEWER_OPTIONS,
	finishPrintWindow as finishPendingPrintWindow,
	openPendingPrintWindow,
	openPrintWindow as sharedOpenPrintWindow,
	planPrintJob,
	printJobRasterFallback,
} from 'ooxml-ui/pptx';
import type { PrintJobCapture, PrintJobPlan } from 'ooxml-ui/pptx';
import type { PptxHandoutMaster, PptxSlide } from 'pptx-viewer-core';
import { inject, ref } from 'vue';
import type { Ref } from 'vue';

import type { PrintSettings } from '../components/print-dialog-types';
import { exportSlideToSvg } from '../export-svg';
import type { CanvasSize } from '../types';
import { useSafeTranslate } from './useSafeTranslate';
import { ViewerOptionsKey } from './useViewerOptionsStore';

/**
 * usePrint: print-dialog state + the print-with-settings flow for the Vue
 * viewer. What prints, through which path and under which (translated) window
 * title is decided by the shared `planPrintJob`; this composable keeps only
 * its capture and its opener. Direct slide printing uses core SVG output
 * (captured instead when the export throws); notes and handouts use the DOM
 * raster path.
 *
 * The DOM-touching pieces are injected so the composable is unit-testable with
 * mocks, exactly like `useExport` injects `rasterizeSlide`:
 *  - `rasterizeSlide(index)` rasterises one slide to a canvas (the host owns the
 *    off-screen `SlideStage` + `html2canvas-pro` integration).
 *  - `openPrintWindow(html)` opens a print window for a full HTML document. A
 *    default implementation (`window.open` → write → print) is supplied.
 */

/**
 * Rasterise the slide at `index` to a canvas. Host-supplied (DOM-coupled).
 * `scaleMultiplier` (default 1) is an extra factor the Print dialog's
 * notes/handouts raster path applies on top of the host's own baseline scale
 * when Options > Advanced > "High quality" is on.
 */
export type RasterizeSlide = (
	index: number,
	scaleMultiplier?: number,
) => Promise<HTMLCanvasElement>;

/**
 * Open a print window for a complete HTML document and trigger printing.
 * Returns `false` if the window was blocked (e.g. popup blocker).
 */
export type OpenPrintWindow = (htmlDocument: string) => boolean;

export interface UsePrintOptions {
	slides: Ref<PptxSlide[]>;
	activeSlideIndex: Ref<number>;
	rasterizeSlide: RasterizeSlide;
	/** Native presentation dimensions used for vector slide printing. */
	slideSize?: Ref<CanvasSize>;
	/** Override the print-window opener (defaults to a `window.open` impl). */
	openPrintWindow?: OpenPrintWindow;
	/**
	 * The deck's handout master, consumed for its background, header/footer/
	 * date/page-number placeholders, and (when authored) positioned slide
	 * rects, in the `handouts` print mode.
	 */
	handoutMaster?: Ref<PptxHandoutMaster | undefined>;
	/** False while the host removed the Print dialog (UI customisation); opening is a no-op. */
	canOpen?: () => boolean;
}

export interface UsePrintResult {
	/** Whether the print dialog is open. */
	isPrintDialogOpen: Ref<boolean>;
	/** Open the print dialog (the toolbar print button). */
	openPrintDialog: () => void;
	/** Close the print dialog without printing. */
	closePrintDialog: () => void;
	/** Run the print flow for the confirmed settings (closes the dialog first). */
	print: (settings: PrintSettings) => Promise<void>;
}

/**
 * Default print-window opener: writes a full HTML doc and calls `print()`.
 * Delegates to the shared `openPrintWindow` (this composable used to carry
 * its own copy that passed `noopener` to `window.open`, which made Chrome
 * return `null` from the call even though the tab genuinely opened, so this
 * always read as "popup blocked" and left a stray blank tab with nothing
 * ever written into it - see `pptx-viewer-shared`'s `print-window.ts` for
 * why `noopener` doesn't belong on a same-origin window the app fills
 * itself).
 */
function defaultOpenPrintWindow(htmlDocument: string): boolean {
	return sharedOpenPrintWindow(htmlDocument);
}

export function usePrint(options: UsePrintOptions): UsePrintResult {
	const { slides, activeSlideIndex, rasterizeSlide } = options;
	// A custom opener gets the complete document in one call, as documented;
	// only the default `window.open` path gets the early-open popup-blocking
	// fix (see `openPendingPrintWindow`), since it owns the window's timing.
	const customOpenWindow = options.openPrintWindow;
	const slideSize = options.slideSize ?? ref<CanvasSize>({ width: 960, height: 540 });
	const t = useSafeTranslate();
	const injectedViewerOptions = inject(ViewerOptionsKey, undefined);

	const isPrintDialogOpen = ref(false);

	function openPrintDialog(): void {
		if (options.canOpen?.() === false) {
			return;
		}
		isPrintDialogOpen.value = true;
	}

	function closePrintDialog(): void {
		isPrintDialogOpen.value = false;
	}

	const title = (plan: PrintJobPlan): string => t(plan.title.key, plan.title.params);

	/**
	 * Capture each planned slide and hand the document to the opener. The
	 * default opener grabs its window before the first await (see
	 * `openPendingPrintWindow`), or the browser blocks it as a popup; a
	 * host-supplied opener owns its own timing and gets the finished document.
	 */
	async function printCaptured(plan: PrintJobPlan, slideList: PptxSlide[]): Promise<void> {
		const pendingWindow = customOpenWindow
			? undefined
			: openPendingPrintWindow(t('pptx.print.preparingToPrint'));
		if (!customOpenWindow && !pendingWindow) {
			console.warn(
				'[PowerPointViewer] Print window was blocked by the browser. Allow popups for this site to print.',
			);
			return;
		}
		try {
			const viewerOptions = injectedViewerOptions?.value ?? DEFAULT_VIEWER_OPTIONS;
			// Options > Advanced > "High quality" raster scale, composed on top of
			// the host's own baseline (2x * Options > Advanced > Image Size/Quality).
			const printScaleMultiplier = viewerOptions.advanced.printHighQuality ? 2 : 1;
			const captures: PrintJobCapture[] = [];
			for (const slideIndex of plan.slideIndices) {
				const canvas = await rasterizeSlide(slideIndex, printScaleMultiplier);
				captures.push({ slideIndex, dataUrl: canvas.toDataURL('image/png') });
			}
			const html = buildRasterPrintJobDocument(
				plan,
				title(plan),
				captures,
				slideList,
				options.handoutMaster?.value,
			);
			if (pendingWindow) {
				finishPendingPrintWindow(pendingWindow, html);
			} else {
				customOpenWindow?.(html);
			}
		} catch (err) {
			console.error('[PowerPointViewer] Print layout failed:', err);
			pendingWindow?.close();
		}
	}

	async function print(settings: PrintSettings): Promise<void> {
		isPrintDialogOpen.value = false;

		const slideList = slides.value;
		const viewerOptions = injectedViewerOptions?.value ?? DEFAULT_VIEWER_OPTIONS;
		const plan = planPrintJob(settings, {
			slides: slideList,
			activeSlideIndex: activeSlideIndex.value,
			includeHiddenSlides: viewerOptions.advanced.printHiddenSlides,
		});
		if (!plan.ok) {
			return;
		}
		const openWindow = customOpenWindow ?? defaultOpenPrintWindow;

		if (plan.output === 'outline') {
			openWindow(buildOutlinePrintJobDocument(plan, title(plan), slideList));
			return;
		}

		if (plan.output === 'vector') {
			try {
				const size = slideSize.value;
				const svgs = plan.slideIndices.map((index) =>
					exportSlideToSvg(slideList[index], size.width, size.height),
				);
				openWindow(buildVectorPrintJobDocument(plan, title(plan), svgs, size));
				return;
			} catch (err) {
				console.warn('[PowerPointViewer] SVG print failed, falling back to raster:', err);
			}
		}

		await printCaptured(printJobRasterFallback(plan), slideList);
	}

	return { isPrintDialogOpen, openPrintDialog, closePrintDialog, print };
}
