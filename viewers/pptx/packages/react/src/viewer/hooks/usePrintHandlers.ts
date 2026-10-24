/**
 * usePrintHandlers -- Print dialog and print-with-settings logic for
 * slides, notes, handouts, and outline layouts.
 *
 * What prints, through which path and under which window title is decided
 * once for every binding by the shared `planPrintJob`; this hook keeps only
 * its capture and its opener:
 * 1. **Vector slides**: each planned slide goes through the core SVG
 *    exporter, producing resolution-independent print output that stays sharp
 *    at any DPI, with no slide flicker. Falls back to the raster path when
 *    the export throws.
 * 2. **Raster** (notes pages, handouts, and the vector fallback): each planned
 *    slide is shown on the live stage and captured via html2canvas as a PNG
 *    data URL.
 * 3. **Outline**: text only.
 */
import {
	buildOutlinePrintJobDocument,
	buildRasterPrintJobDocument,
	buildVectorPrintJobDocument,
	finishPrintWindow,
	openPendingPrintWindow,
	openPrintWindow,
	planPrintJob,
	printJobRasterFallback,
	resolveImageResolutionScale,
} from 'ooxml-ui/pptx';
import type { PrintJobPlan } from 'ooxml-ui/pptx';
import type { PptxData, PptxElement, PptxHandoutMaster, PptxSlide } from 'pptx-viewer-core';
import { useState } from 'react';
import type { RefObject } from 'react';

import type { PrintSettings } from '../components/print-dialog-types';
import { useViewerOptionsContext } from '../components/viewer-options-context';
import { useTranslation } from '../i18n/use-translation';
import type { CanvasSize } from '../types';
import { captureSlidesAsPngDataUrls } from '../utils/export';
import { exportSlideToSvg } from '../utils/export-svg';
import { buildSaveSlides } from '../utils/template-editing';

/* ------------------------------------------------------------------ */
/*  Types                                                              */
/* ------------------------------------------------------------------ */

export interface UsePrintHandlersInput {
	/** The editable slides (template elements partitioned out, as the viewer state holds them). */
	slides: PptxSlide[];
	activeSlideIndex: number;
	canvasStageRef: RefObject<HTMLDivElement | null>;
	setActiveSlideIndex: React.Dispatch<React.SetStateAction<number>>;
	/** Slide size in CSS pixels: the vector slides are exported at it. */
	canvasSize: CanvasSize;
	/**
	 * The layout and master elements partitioned out of each slide, merged back
	 * for the vector export and the outline (`buildSaveSlides`).
	 */
	templateElementsBySlideId?: Record<string, PptxElement[]>;
	/** The deck's handout master, painted on every handouts page. */
	handoutMaster?: PptxHandoutMaster;
	/** @deprecated Pass `canvasSize` and `handoutMaster`; only its `handoutMaster` is still read. */
	pptxData?: PptxData;
}

export interface PrintHandlersResult {
	handlePrint: () => void;
	handlePrintWithSettings: (settings: PrintSettings) => Promise<void>;
	handlePrintSvg: (settings: PrintSettings) => Promise<void>;
	isPrintDialogOpen: boolean;
	setIsPrintDialogOpen: React.Dispatch<React.SetStateAction<boolean>>;
}

const BLOCKED_WARNING =
	'[PowerPointViewer] Print window was blocked by the browser. Allow popups for this site to print.';

/* ------------------------------------------------------------------ */
/*  Hook                                                               */
/* ------------------------------------------------------------------ */

export function usePrintHandlers(input: UsePrintHandlersInput): PrintHandlersResult {
	const { slides, activeSlideIndex, canvasStageRef, setActiveSlideIndex, canvasSize } = input;
	const handoutMaster = input.handoutMaster ?? input.pptxData?.handoutMaster;
	const [isPrintDialogOpen, setIsPrintDialogOpen] = useState(false);
	const { t } = useTranslation();
	const viewerOptions = useViewerOptionsContext();
	// Options > Advanced > "High quality" raster scale for the capture path,
	// composed with Options > Advanced > Image Size/Quality's own resolution
	// scale so the two settings multiply rather than fight each other.
	const printRasterScale =
		(viewerOptions.advanced.printHighQuality ? 4 : 3) * resolveImageResolutionScale(viewerOptions);

	const handlePrint = () => {
		setIsPrintDialogOpen(true);
	};

	/** The complete slides (template elements merged back) the plan prints from. */
	const deckSlides = (): PptxSlide[] =>
		input.templateElementsBySlideId
			? buildSaveSlides(slides, input.templateElementsBySlideId)
			: slides;

	const plan = (settings: PrintSettings, deck: PptxSlide[]) =>
		planPrintJob(settings, {
			slides: deck,
			activeSlideIndex,
			// Options > Advanced > "Print hidden slides".
			includeHiddenSlides: viewerOptions.advanced.printHiddenSlides,
		});

	const title = (job: PrintJobPlan): string => t(job.title.key, job.title.params);

	/**
	 * Captures each planned slide on the live stage and writes the document
	 * into `printWindow` (opened before the first await, see
	 * `openPendingPrintWindow`). Not flicker-free: each slide is shown in turn,
	 * and the slide on screen before is restored.
	 */
	const runRasterPrint = async (printWindow: Window, job: PrintJobPlan, deck: PptxSlide[]) => {
		try {
			if (!canvasStageRef.current) {
				printWindow.close();
				return;
			}
			const captures = await captureSlidesAsPngDataUrls(
				canvasStageRef,
				job.slideIndices,
				setActiveSlideIndex,
				activeSlideIndex,
				{ scale: printRasterScale },
			);
			if (captures.length === 0) {
				printWindow.close();
				return;
			}
			finishPrintWindow(
				printWindow,
				buildRasterPrintJobDocument(job, title(job), captures, deck, handoutMaster),
			);
		} catch (err) {
			console.error('[PowerPointViewer] Print layout failed:', err);
			printWindow.close();
		}
	};

	/** Opens a finished document in one call (no await happened since the click). */
	const openNow = (html: string) => {
		if (!openPrintWindow(html)) {
			console.warn(BLOCKED_WARNING);
		}
	};

	/** Opens the print window now, still inside the click, then captures. */
	const startRasterPrint = async (job: PrintJobPlan, deck: PptxSlide[]) => {
		const printWindow = openPendingPrintWindow(t('pptx.print.preparingToPrint'));
		if (!printWindow) {
			console.warn(BLOCKED_WARNING);
			return;
		}
		await runRasterPrint(printWindow, job, deck);
	};

	/** The vector document of a `vector` job, or `undefined` when the SVG export threw. */
	const vectorDocument = (job: PrintJobPlan, deck: PptxSlide[]): string | undefined => {
		try {
			const svgs = job.slideIndices.map((index) =>
				exportSlideToSvg(deck[index], canvasSize.width, canvasSize.height),
			);
			return buildVectorPrintJobDocument(job, title(job), svgs, canvasSize);
		} catch (err) {
			console.warn('[PowerPointViewer] SVG print path failed, falling back to raster:', err);
			return undefined;
		}
	};

	/**
	 * Runs a planned job. The outline and vector slides are assembled
	 * synchronously, inside the click, and opened in one call; captured pages
	 * open a placeholder window first. A failed vector export falls back to
	 * captures: nothing was awaited yet, so that window can still open.
	 */
	const runPrintJob = async (settings: PrintSettings, vector: boolean) => {
		setIsPrintDialogOpen(false);
		const deck = deckSlides();
		const job = plan(settings, deck);
		if (!job.ok) {
			return;
		}
		if (job.output === 'outline') {
			openNow(buildOutlinePrintJobDocument(job, title(job), deck));
			return;
		}
		const html = vector && job.output === 'vector' ? vectorDocument(job, deck) : undefined;
		if (html !== undefined) {
			openNow(html);
			return;
		}
		await startRasterPrint(printJobRasterFallback(job), deck);
	};

	/** Single entry point for the "Print" button inside `PrintDialog`. */
	const handlePrintSvg = (settings: PrintSettings) => runPrintJob(settings, true);

	/**
	 * Public entry point for the raster path on its own (kept for API
	 * back-compat; `PrintDialog` itself goes through `handlePrintSvg`). Full-page
	 * slides are captured here too, never exported as SVG.
	 */
	const handlePrintWithSettings = (settings: PrintSettings) => runPrintJob(settings, false);

	return {
		handlePrint,
		handlePrintWithSettings,
		handlePrintSvg,
		isPrintDialogOpen,
		setIsPrintDialogOpen,
	};
}
