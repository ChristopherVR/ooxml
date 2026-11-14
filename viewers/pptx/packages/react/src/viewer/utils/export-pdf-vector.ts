/**
 * File > Export > Create PDF: vector pages by default (the shared `runPdfExport`), reading the
 * live slide stage after each slide switch, or the picture export (`exportAllSlidesAsPdf`)
 * when Options > Advanced > "Save slides as pictures" is on.
 */
import { runPdfExport } from 'ooxml-ui/pptx';
import type {
	CanvasSize,
	PdfExportSettings,
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
	onApproximated?: (result: VectorPdfExportResult) => void;
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
		});
	} finally {
		setActiveSlide(currentSlideIndex);
	}
}
