/**
 * File > Export > Create PDF: vector pages by default (the shared `runPdfExport`), reading the
 * live slide stage after each slide switch, or the picture export (`exportAllSlidesAsPdf`)
 * when Options > Advanced > "Save slides as pictures" is on.
 */
import { measureExportRoot, runPdfExport, untransformExportClone } from 'ooxml-ui/pptx';
import type {
	CanvasSize,
	CompatibilityWarningToast,
	PdfExportSettings,
	PdfSlideSelection,
	SlideSizeEmu,
	VectorPdfExportResult,
} from 'ooxml-ui/pptx';
import type React from 'react';

import { renderToCanvas } from '../../lib/canvas-export';
import { downloadBlob } from './dom-helpers';
import type { PdfExportOptions } from './export-helpers';
import { ignoreExportOverlayElements, waitForRender } from './export-helpers';
import { exportAllSlidesAsPdf } from './export-pdf';

export interface PresentationPdfOptions extends PdfExportOptions {
	settings: PdfExportSettings;
	slides: readonly { hidden?: boolean }[];
	canvasSize: CanvasSize;
	slideSize?: SlideSizeEmu | undefined;
	/** Told what the vector PDF approximated. */
	onApproximated?: ((result: VectorPdfExportResult) => void) | undefined;
	/** The range the PDF options step chose (`choosePdfExportSelection`); every slide when omitted. */
	selection?: PdfSlideSelection | undefined;
	/** Show the approximation notice in the viewer's toast stack. */
	notify?: ((toast: CompatibilityWarningToast) => void) | undefined;
	/** Dismiss an earlier PDF approximation notice when a later export is clean. */
	dismissNotice?: ((id: string) => void) | undefined;
}

/** Export the presentation as a PDF download (see the module comment). */
export async function exportPresentationPdf(
	slideStageRef: React.RefObject<HTMLElement | null>,
	setActiveSlide: (index: number) => void,
	currentSlideIndex: number,
	filename: string,
	options: PresentationPdfOptions,
): Promise<VectorPdfExportResult | undefined> {
	const { scale = 2, onProgress, signal } = options;
	try {
		return await runPdfExport({
			settings: options.settings,
			vector: {
				document,
				slides: options.slides,
				canvasSize: options.canvasSize,
				...(options.slideSize ? { slideSize: options.slideSize } : {}),
				...options.selection,
				mountSlide: async (index) => {
					setActiveSlide(index);
					await waitForRender(150);
					return slideStageRef.current;
				},
				html2canvasFallback: (sourceRect, outputSize) => {
					const stage = slideStageRef.current;
					if (!stage) {
						return Promise.reject(new Error('No slide stage.'));
					}
					return renderToCanvas(stage, {
						scale: outputSize.width / (sourceRect.width || 1),
						x: sourceRect.x,
						y: sourceRect.y,
						width: sourceRect.width,
						height: sourceRect.height,
						useCORS: true,
						allowTaint: true,
						backgroundColor: null,
						logging: false,
						ignoreElements: ignoreExportOverlayElements,
						// The vector pass draws the stage without its zoom; the fallback must match.
						...(measureExportRoot(stage).untransformedRoot && {
							onclone: (_doc: Document, clone: HTMLElement) => untransformExportClone(clone),
						}),
					});
				},
				ignoreElement: ignoreExportOverlayElements,
				rasterScale: scale,
				title: filename.replace(/\.pdf$/iu, ''),
				...(signal ? { signal } : {}),
				...(onProgress ? { onProgress } : {}),
			},
			pictures: () =>
				exportAllSlidesAsPdf(
					slideStageRef,
					options.slides.length,
					setActiveSlide,
					currentSlideIndex,
					filename,
					{ scale, ...(onProgress ? { onProgress } : {}), ...(signal ? { signal } : {}) },
				),
			save: (blob) => downloadBlob(blob, filename),
			...(options.onApproximated ? { onApproximated: options.onApproximated } : {}),
			...(options.notify ? { notify: options.notify } : {}),
			...(options.dismissNotice ? { dismissNotice: options.dismissNotice } : {}),
		});
	} finally {
		setActiveSlide(currentSlideIndex);
	}
}
