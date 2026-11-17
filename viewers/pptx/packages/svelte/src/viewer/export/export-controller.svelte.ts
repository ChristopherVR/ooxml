import type {
	CanvasSize,
	CompatibilityWarningToast,
	MountedPdfSlide,
	PdfExportSettings,
	VectorPdfExportResult,
	RasterizeElementResult,
	RasterizeElementTilesResult,
} from 'ooxml-ui/pptx';
import {
	canCopyImageToClipboard,
	copyPngToClipboard,
	downloadBlob,
	downloadDataUrl,
	exportAbortError,
	exportDeckJson,
	rasterResultToPngBlob,
	rasterResultToPngDataUrl,
	resolveExportBaseName,
	runPdfExport,
	videoFileName,
} from 'ooxml-ui/pptx';
import type { PptxData, PptxSlide } from 'pptx-viewer-core';

import type { Translator } from '../../i18n/translator';
import type { ExportGifOptions } from './export-gif';
import { exportSlidesToGifBlob } from './export-gif';
import type { OpenPrintWindow, PrintOptions } from './export-print';
import { printSlides } from './export-print';
import type { ExportVideoOptions } from './export-video';
import { exportSlidesToWebmBlob } from './export-video';
import { addTiledPageImages } from './pdf-tile-page';

/** Rasterise the slide at `index` to an `HTMLCanvasElement`. Injected so the
 * controller stays DOM-capture-free and unit-testable. */
export type RasterizeSlide = (
	index: number,
	scaleMultiplier?: number,
) => Promise<HTMLCanvasElement>;

/** Per-slide progress callback: `(currentSlideIndex, totalSlides)`. */
export type ExportProgress = (current: number, total: number) => void;

/** Options for the multi-slide PDF export (progress + cooperative cancel). */
export interface ExportPdfOptions {
	/** Capture-phase progress callback: `(currentSlide, totalSlides)`. */
	onProgress?: ExportProgress;
	/** Abort the export early; the loop checks this between slides. */
	signal?: AbortSignal;
	/** Include hidden slides (vector export; default from Options > Advanced > Save as PDF). */
	includeHiddenSlides?: boolean;
	/** 1-based slide range (vector export). */
	from?: number;
	to?: number;
	/** Explicit 0-based slides (the current slide); overrides `from`/`to` (vector export). */
	slideIndexes?: readonly number[];
	/** `pictures` writes one image per slide (the compatibility export); default from the options. */
	mode?: PdfExportSettings['mode'];
	/** Told what the vector PDF approximated (effects drawn as pictures, substituted fonts). */
	onReport?: (result: VectorPdfExportResult) => void;
}

export interface ExportControllerDeps {
	/** Live slide count; read fresh on every call. */
	getSlideCount(): number;
	/** Active slide index (0-based); the PNG export's default target. */
	getCurrent(): number;
	/** Live slide canvas size (px), for the PDF page size. */
	getCanvasSize(): CanvasSize;
	/** Live slide list (print needs slide notes/titles, not just the count). */
	getSlides(): PptxSlide[];
	rasterizeSlide: RasterizeSlide;
	/**
	 * PNG-export / "copy slide as image" only: returns the full
	 * `RasterizeElementResult` (tiled `png-bytes` included) instead of a
	 * plain canvas. Optional so existing test fixtures that only configure
	 * `rasterizeSlide` keep working; `exportSlidePng`/`copySlideAsImage` fall
	 * back to wrapping `rasterizeSlide`'s canvas when omitted.
	 */
	rasterizeSlideToRaster?: (
		index: number,
		scaleMultiplier?: number,
	) => Promise<RasterizeElementResult>;
	/**
	 * PDF export only: returns the raw per-tile canvases (no PNG stitching)
	 * instead of a plain canvas. Optional so existing test fixtures that only
	 * configure `rasterizeSlide` keep working; `exportPdf` falls back to
	 * placing a single full-page image (today's behaviour) when omitted.
	 */
	rasterizeSlideToTiles?: (
		index: number,
		scaleMultiplier?: number,
	) => Promise<RasterizeElementTilesResult>;
	/** Print-surface opener override; see `export-print.ts` for the default. */
	openPrintWindow?: OpenPrintWindow;
	/** Base file name (without extension) for downloads. Defaults to `presentation`. */
	fileName?: string;
	/** Live presentation data for the deck-JSON export; undefined before a load. */
	getDeckData?(): PptxData | undefined;
	/** Source file name for the deck-JSON download (`deck.pptx` -> `deck.json`). */
	getFileName?(): string | undefined;
	/** Options > Advanced > "Print hidden slides". Defaults to `false` (excluded). */
	getIncludeHiddenSlides?(): boolean;
	/** Options > Advanced > "High quality" raster scale for the print fallback path. */
	getPrintHighQuality?(): boolean;
	/** Live translator, for the print window title. English when omitted. */
	getTranslator?(): Translator;
	/** Mount a slide on the capture stage, for the vector PDF export. */
	mountSlide?: (index: number) => Promise<MountedPdfSlide>;
	/** Options > Advanced > Save as PDF. Defaults to vector pages without hidden slides. */
	getPdfExportSettings?(): PdfExportSettings | undefined;
	/** Options > Advanced image resolution multiplier, for effects drawn as pictures. */
	getImageResolutionScale?(): number;
	/** Show the PDF approximation notice in the viewer's toast stack. */
	notify?(toast: CompatibilityWarningToast): void;
	/** The host's `onpdfexportreport`, told what every vector PDF approximated. */
	onPdfExportReport?(result: VectorPdfExportResult): void;
}

/**
 * Export controller (runes class): render slides to PNG / PDF / animated GIF /
 * WebM video, plus the print flow. Started as a Svelte port of Vue's
 * `useExport` composable (`packages/vue/src/viewer/composables/useExport.ts`)
 * and now also covers React's `useExportHandlers` GIF/video surface (via
 * `export-gif.ts` / `export-video.ts`) and Vue's `usePrint`
 * (`export-print.ts`), reshaped as a `$state`-backed class to match this
 * package's runes-class convention (see `EditorState` in
 * `../editor/editor-state.svelte.ts`) rather than Vue `Ref`s.
 *
 * Rasterisation is delegated to the injected `rasterizeSlide` (the host owns
 * the off-screen `SlideStage` mount + `html2canvas-pro`, see
 * `rasterize-slide.ts`). `jspdf` is loaded lazily so it stays out of the main
 * bundle.
 */
export class ExportController {
	/** True while an export is running (disable UI / show a spinner). */
	exporting = $state(false);

	readonly #deps: ExportControllerDeps;

	constructor(deps: ExportControllerDeps) {
		this.#deps = deps;
	}

	/** `rasterizeSlideToRaster` when supplied, else `rasterizeSlide`'s canvas wrapped as a result. */
	async #captureRaster(index: number): Promise<RasterizeElementResult> {
		if (this.#deps.rasterizeSlideToRaster) {
			return this.#deps.rasterizeSlideToRaster(index);
		}
		const canvas = await this.#deps.rasterizeSlide(index);
		return {
			kind: 'canvas',
			canvas,
			strategy: 'html2canvas',
			width: canvas.width,
			height: canvas.height,
		};
	}

	/** Export a single slide as a PNG download. Defaults to the current slide. */
	async exportSlidePng(index?: number): Promise<void> {
		const targetIndex = index ?? this.#deps.getCurrent();
		if (this.exporting || targetIndex < 0 || targetIndex >= this.#deps.getSlideCount()) {
			return;
		}
		this.exporting = true;
		try {
			const result = await this.#captureRaster(targetIndex);
			downloadDataUrl(
				await rasterResultToPngDataUrl(result),
				`${resolveExportBaseName(this.#deps.fileName)}-slide-${targetIndex + 1}.png`,
			);
		} finally {
			this.exporting = false;
		}
	}

	/** Copy a slide to the system clipboard as a PNG image. */
	async copySlideAsImage(index?: number): Promise<void> {
		const targetIndex = index ?? this.#deps.getCurrent();
		if (
			this.exporting ||
			targetIndex < 0 ||
			targetIndex >= this.#deps.getSlideCount() ||
			!canCopyImageToClipboard()
		) {
			return;
		}
		this.exporting = true;
		try {
			// The clipboard write starts inside the click; Safari rejects it after an await.
			await copyPngToClipboard(
				this.#captureRaster(targetIndex).then((result) => rasterResultToPngBlob(result)),
			);
		} finally {
			this.exporting = false;
		}
	}

	/**
	 * Serialize the live deck to `pptx-viewer-json` and trigger the download.
	 * Pure data serialization (shared `exportDeckJson`): synchronous, no
	 * rasterization pipeline, no progress modal, no `exporting` toggle.
	 */
	exportJson(): void {
		const data = this.#deps.getDeckData?.();
		if (!data) {
			return;
		}
		exportDeckJson(data, this.#deps.getFileName?.() ?? this.#deps.fileName ?? null);
	}

	/** Export every slide as a multi-page PDF download (one slide per page). */
	async exportPdf(options: ExportPdfOptions = {}): Promise<void> {
		const total = this.#deps.getSlideCount();
		if (this.exporting || total === 0) {
			return;
		}
		// eslint-disable-next-line one-var -- separated from `total` above by a guard clause
		const { onProgress, signal } = options;
		const defaults = this.#deps.getPdfExportSettings?.() ?? {
			mode: 'vector',
			includeHiddenSlides: false,
		};
		const mountSlide = this.#deps.mountSlide;
		const mode = mountSlide ? (options.mode ?? defaults.mode) : 'pictures';
		if (mode === 'vector' && mountSlide) {
			this.exporting = true;
			try {
				const baseName = resolveExportBaseName(this.#deps.fileName);
				await runPdfExport({
					settings: {
						mode,
						includeHiddenSlides: options.includeHiddenSlides ?? defaults.includeHiddenSlides,
					},
					vector: {
						document,
						slides: this.#deps.getSlides(),
						canvasSize: this.#deps.getCanvasSize(),
						mountSlide,
						rasterScale: 2 * (this.#deps.getImageResolutionScale?.() ?? 1),
						title: baseName,
						...(options.from !== undefined ? { from: options.from } : {}),
						...(options.to !== undefined ? { to: options.to } : {}),
						...(options.slideIndexes ? { slideIndexes: options.slideIndexes } : {}),
						...(signal ? { signal } : {}),
						...(onProgress ? { onProgress } : {}),
					},
					pictures: () => Promise.resolve(),
					save: (blob) => downloadBlob(blob, `${baseName}.pdf`),
					onApproximated: (result) => {
						options.onReport?.(result);
						this.#deps.onPdfExportReport?.(result);
					},
					...(this.#deps.notify ? { notify: (toast) => this.#deps.notify?.(toast) } : {}),
				});
			} finally {
				this.exporting = false;
			}
			return;
		}
		this.exporting = true;
		try {
			const { jsPDF } = await import('jspdf'),
				{ width, height } = this.#deps.getCanvasSize(),
				orientation = width >= height ? 'landscape' : 'portrait',
				pdf = new jsPDF({ orientation, unit: 'px', format: [width, height], compress: true }),
				rasterizeSlideToTiles = this.#deps.rasterizeSlideToTiles;
			for (let i = 0; i < total; i++) {
				if (signal?.aborted) {
					throw exportAbortError();
				}
				onProgress?.(i, total);
				if (i > 0) {
					pdf.addPage([width, height], orientation);
				}
				if (rasterizeSlideToTiles) {
					// Escapes the browser canvas cap the same way PNG export does,
					// since a PDF page has no canvas-size limit of its own.
					addTiledPageImages(pdf, await rasterizeSlideToTiles(i), width, height);
				} else {
					const canvas = await this.#deps.rasterizeSlide(i);
					pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, width, height);
				}
			}
			pdf.save(`${resolveExportBaseName(this.#deps.fileName)}.pdf`);
		} finally {
			this.exporting = false;
		}
	}

	/**
	 * Export every slide as an animated GIF download. Frame delays come from
	 * the shared `planGifFrames` plan (default duration + per-slide overrides);
	 * see `export-gif.ts`. Supports progress + AbortSignal like `exportPdf`.
	 */
	async exportGif(options: ExportGifOptions = {}): Promise<void> {
		if (this.exporting || this.#deps.getSlideCount() === 0) {
			return;
		}
		this.exporting = true;
		try {
			const blob = await exportSlidesToGifBlob(
				{
					getSlideCount: () => this.#deps.getSlideCount(),
					rasterizeSlide: (index) => this.#deps.rasterizeSlide(index),
				},
				options,
			);
			downloadBlob(blob, `${resolveExportBaseName(this.#deps.fileName)}.gif`);
		} finally {
			this.exporting = false;
		}
	}

	/**
	 * Export every slide as a video download (WebM, or MP4 on Safari) (MediaRecorder over a canvas
	 * capture stream, timing from the shared `planVideoSegments` plan); see
	 * `export-video.ts`. Supports progress + AbortSignal like `exportPdf`.
	 */
	async exportVideo(options: ExportVideoOptions = {}): Promise<void> {
		if (this.exporting || this.#deps.getSlideCount() === 0) {
			return;
		}
		this.exporting = true;
		try {
			const blob = await exportSlidesToWebmBlob(
				{
					getSlideCount: () => this.#deps.getSlideCount(),
					rasterizeSlide: (index) => this.#deps.rasterizeSlide(index),
				},
				options,
			);
			downloadBlob(blob, videoFileName(resolveExportBaseName(this.#deps.fileName), blob));
		} finally {
			this.exporting = false;
		}
	}

	/**
	 * Assemble the shared print document for the given (partial) settings and
	 * open it in the print surface (hidden iframe by default; see
	 * `export-print.ts` for the popup-blocker caveats of custom openers).
	 * Resolves `true` when the print surface opened.
	 */
	async print(options: PrintOptions = {}): Promise<boolean> {
		if (this.exporting || this.#deps.getSlides().length === 0) {
			return false;
		}
		this.exporting = true;
		try {
			return await printSlides(
				{
					getSlides: () => this.#deps.getSlides(),
					getCurrent: () => this.#deps.getCurrent(),
					getCanvasSize: () => this.#deps.getCanvasSize(),
					rasterizeSlide: (index, scaleMultiplier) =>
						this.#deps.rasterizeSlide(index, scaleMultiplier),
					openPrintWindow: this.#deps.openPrintWindow,
					getIncludeHiddenSlides: this.#deps.getIncludeHiddenSlides,
					getPrintHighQuality: this.#deps.getPrintHighQuality,
					getHandoutMaster: () => this.#deps.getDeckData?.()?.handoutMaster,
					getTranslator: this.#deps.getTranslator,
				},
				options,
			);
		} finally {
			this.exporting = false;
		}
	}
}
