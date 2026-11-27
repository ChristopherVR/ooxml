import type { DocumentModel } from 'ooxml-core/docx';
import type { DocxPdfOptions, DocxPdfResult } from 'ooxml-core/docx/pdf-export';
import { createCanvasMeasurer } from '../canvas-measurer';

/**
 * Saves a Word document as a PDF with real text and vector graphics, on the pages Print Layout
 * shows. The writer is `ooxml-core/docx/pdf-export` (DOM-free, also usable in Node without the
 * editor); this wrapper loads it on first use, so importing the editor does not download it, and
 * measures with the canvas Print Layout uses when the page has a DOM and no measurer is given.
 */
export async function exportDocxPdf(
	model: DocumentModel,
	options: DocxPdfOptions = {},
): Promise<DocxPdfResult> {
	const { exportDocxPdf: run } = await import('./export-pdf-load');
	if (options.measurer || typeof document === 'undefined') return run(model, options);
	return run(model, { ...options, measurer: createCanvasMeasurer() });
}
