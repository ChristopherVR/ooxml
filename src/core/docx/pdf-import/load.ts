import type { DocumentModel, LoadedDocument } from '../model';
import { loadDocx } from '../parse';
import { convertPdf, type ImportPdfOptions } from './import';
import type { PdfImportReport } from './report';

export interface ImportPdfResult {
	/** The converted document: `save()` writes it as a .docx. */
	document: LoadedDocument;
	model: DocumentModel;
	report: PdfImportReport;
}

/** Converts `bytes` (a PDF) to an editable Word document. */
export async function importPdf(
	bytes: Uint8Array,
	options: ImportPdfOptions = {},
): Promise<ImportPdfResult> {
	const { docx, report, warnings } = await convertPdf(bytes, options);
	const document = await loadDocx(docx);
	document.model.warnings.unshift(...warnings);
	return { document, model: document.model, report };
}
