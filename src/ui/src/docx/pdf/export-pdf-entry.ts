import type { DocumentModel } from 'ooxml-core/docx';
import type { DocxPdfOptions, DocxPdfResult } from './export-pdf';

/**
 * Saves a Word document as a PDF with real text and vector graphics, on the pages Print Layout
 * shows (see `export-pdf.ts`). The PDF writer is loaded on first use, so importing the editor does
 * not download it.
 */
export async function exportDocxPdf(
	model: DocumentModel,
	options?: DocxPdfOptions,
): Promise<DocxPdfResult> {
	const { exportDocxPdf: run } = await import('./export-pdf');
	return run(model, options);
}
