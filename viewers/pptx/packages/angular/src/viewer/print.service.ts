/**
 * PrintService: print orchestration for the Angular viewer.
 *
 * The shared print job plan (`planPrintJob` in `ooxml-ui/pptx`) decides what
 * prints, through which path and under which window title, and assembles the
 * document. Direct slide printing uses core SVG output; notes and handouts
 * retain the DOM raster path.
 *
 * Slide capture is decoupled from rendering: the host viewer passes a
 * `captureSlide(index)` callback that flips the live stage to `index` and
 * rasterises it (typically via {@link ExportService.renderElement} +
 * `canvas.toDataURL`). The viewer owns the one reused stage node, so capture
 * must be sequential, exactly as the React raster path does.
 *
 * Dialog open/close state and the active settings are exposed as signals so
 * the toolbar and dialog component can bind to them directly.
 *
 * Provide at the component level so its lifetime tracks the host viewer:
 * `@Component({ providers: [PrintService] })`.
 */

import { Injectable, inject, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import {
	DEFAULT_CANVAS_HEIGHT,
	DEFAULT_CANVAS_WIDTH,
	buildOutlinePrintJobDocument,
	buildRasterPrintJobDocument,
	buildVectorPrintJobDocument,
	finishPrintWindow,
	openPendingPrintWindow,
	openPrintWindow,
	planPrintJob,
	printJobRasterFallback,
} from 'ooxml-ui/pptx';
import type { PrintJobCapture, PrintJobPlan, PrintJobSlideSize } from 'ooxml-ui/pptx';
import type { PptxHandoutMaster, PptxSlide } from 'pptx-viewer-core';

import { exportSlideToSvg } from './export-svg';
import { DEFAULT_PRINT_SETTINGS, validatePrintSettings } from './print-helpers';
import type { PrintSettings } from './print-helpers';
import { ViewerCustomizationService } from './viewer-customization.service';

/**
 * Captures the slide at `index` (zero-based) to a PNG `data:` URL. The viewer
 * supplies this: it flips the live stage to `index` and rasterises it.
 * Returning `null`/empty for a slide skips it.
 */
export type CaptureSlideFn = (index: number) => Promise<string | null>;

@Injectable()
export class PrintService {
	private readonly translate = inject(TranslateService);
	private readonly customization = inject(ViewerCustomizationService, { optional: true });

	/** Whether the print dialog is currently open. */
	readonly isDialogOpen = signal(false);

	/** The current (validated) print settings. */
	readonly settings = signal<PrintSettings>({ ...DEFAULT_PRINT_SETTINGS });

	/** Open the print dialog (a no-op when the host removed the `print` dialog). */
	openDialog(): void {
		if (this.customization?.dialogAvailable('print') !== false) {
			this.isDialogOpen.set(true);
		}
	}

	/** Close the print dialog. */
	closeDialog(): void {
		this.isDialogOpen.set(false);
	}

	/**
	 * Validate + store settings against the slide count. Returns the resolved
	 * settings so callers can use them immediately.
	 */
	updateSettings(partial: Partial<PrintSettings>, slideCount: number): PrintSettings {
		const resolved = validatePrintSettings({ ...this.settings(), ...partial }, slideCount);
		this.settings.set(resolved);
		return resolved;
	}

	/**
	 * Run the full print flow for the given settings. What prints, through
	 * which path and under which window title is decided by the shared
	 * `planPrintJob`; this service keeps only its capture and its opener.
	 *
	 * Outline and full-page slides need no rasterisation (the slides are core
	 * SVG; they are captured instead when the SVG export throws). Notes and
	 * handouts call `captureSlide` sequentially and skip slides that fail.
	 *
	 * @returns `true` if a print window was opened, `false` if blocked (popup
	 *          blocker) or there was nothing to print.
	 */
	async print(
		rawSettings: PrintSettings,
		slides: PptxSlide[],
		activeSlideIndex: number,
		captureSlide: CaptureSlideFn,
		slideSize: Readonly<{ width: number; height: number }> = {
			width: DEFAULT_CANVAS_WIDTH,
			height: DEFAULT_CANVAS_HEIGHT,
		},
		/** Options > Advanced > "Print hidden slides". Defaults to PowerPoint's own default (excluded). */
		includeHiddenSlides = false,
		/**
		 * The deck's handout master, consumed for its background, header/footer/
		 * date/page-number placeholders, and (when authored) positioned slide
		 * rects, in the `handouts` print mode. Optional and defaulted to
		 * `undefined` so existing callers keep printing the bare grid.
		 */
		handoutMaster?: PptxHandoutMaster,
	): Promise<boolean> {
		this.closeDialog();

		const plan = planPrintJob(rawSettings, { slides, activeSlideIndex, includeHiddenSlides });
		if (!plan.ok) {
			return false;
		}
		const title = this.translate.instant(plan.title.key, plan.title.params);

		if (plan.output === 'outline') {
			return this._open(buildOutlinePrintJobDocument(plan, title, slides));
		}
		if (plan.output === 'vector') {
			const html = this._vectorDocument(plan, title, slides, slideSize);
			if (html !== undefined) {
				return this._open(html);
			}
		}
		return this._printCaptured(
			printJobRasterFallback(plan),
			title,
			slides,
			captureSlide,
			handoutMaster,
		);
	}

	/** The vector document of a `vector` plan, or `undefined` when the SVG export threw. */
	private _vectorDocument(
		plan: PrintJobPlan,
		title: string,
		slides: PptxSlide[],
		slideSize: PrintJobSlideSize,
	): string | undefined {
		try {
			const svgs = plan.slideIndices.map((index) =>
				exportSlideToSvg(slides[index], slideSize.width, slideSize.height),
			);
			return buildVectorPrintJobDocument(plan, title, svgs, slideSize);
		} catch (err) {
			console.warn('[PowerPointViewer] SVG print failed, falling back to raster:', err);
			return undefined;
		}
	}

	/**
	 * Capture the plan's slides to PNG data URLs and print them. `captureSlide`
	 * awaits per slide, so the print window has to be opened NOW, before that
	 * first await, or the browser silently blocks it as a popup the instant the
	 * call stack leaves the click's user gesture (see `openPendingPrintWindow`).
	 */
	private async _printCaptured(
		plan: PrintJobPlan,
		title: string,
		slides: PptxSlide[],
		captureSlide: CaptureSlideFn,
		handoutMaster: PptxHandoutMaster | undefined,
	): Promise<boolean> {
		const pendingWindow = openPendingPrintWindow(
			this.translate.instant('pptx.print.preparingToPrint'),
		);
		if (!pendingWindow) {
			return false;
		}

		const captures: PrintJobCapture[] = [];
		for (const slideIndex of plan.slideIndices) {
			const dataUrl = await captureSlide(slideIndex);
			if (dataUrl) {
				captures.push({ slideIndex, dataUrl });
			}
		}

		if (captures.length === 0) {
			pendingWindow.close();
			return false;
		}
		finishPrintWindow(
			pendingWindow,
			buildRasterPrintJobDocument(plan, title, captures, slides, handoutMaster),
		);
		return true;
	}

	/* ---------------------------------------------------------------- */
	/*  Internal: print-window side effects (DOM only)                  */
	/* ---------------------------------------------------------------- */

	/**
	 * Open a print window, write the document, focus, and trigger printing, all
	 * in one synchronous call. Used by the outline/slides paths, which never
	 * `await` before this point so the window is never at risk of being popup
	 * blocked -- no need for the placeholder-then-real-document split `print`'s
	 * notes/handouts branch needs (see `openPendingPrintWindow`/`finishPrintWindow`).
	 */
	private _open(htmlDocument: string): boolean {
		return openPrintWindow(htmlDocument);
	}
}
