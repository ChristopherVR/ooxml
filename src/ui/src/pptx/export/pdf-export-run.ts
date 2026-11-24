/**
 * File > Save as PDF as every binding runs it: vector pages (`exportSlidesVectorPdf`) unless
 * Options > Advanced > "Save slides as pictures" asks for the earlier one-picture-per-slide
 * export, which the binding keeps as `pictures`. Framework-free: the binding supplies its own
 * capture stage, picture export, download and toast stack.
 */
import type { PdfExportSettings } from '../render/options/viewer-options-apply';
import type { CompatibilityWarningToast } from '../render/compatibility-warning-toasts';
import { pdfExportNoticeToast } from './pdf-export-notice';
import {
	exportSlidesVectorPdf,
	type VectorPdfExportOptions,
	type VectorPdfExportResult,
} from './pdf-vector-export-entry';

export interface RunPdfExportOptions {
	settings: PdfExportSettings;
	/** Everything the vector export needs; hidden slides follow `settings` unless set here. */
	vector: VectorPdfExportOptions;
	/** The binding's picture export (the compatibility option). */
	pictures: () => Promise<void>;
	/** Save the finished PDF (a download, or the host's storage). */
	save: (blob: Blob) => void | Promise<void>;
	/** Told what the vector export approximated, when it approximated anything. */
	onApproximated?: (result: VectorPdfExportResult) => void;
	/**
	 * Show the binding's one non-blocking notice (its compatibility toast stack) summarising what
	 * the vector export approximated. Not called when nothing was.
	 */
	notify?: (toast: CompatibilityWarningToast) => void;
}

/** The prefix of the console note listing what a vector PDF approximated. */
export const PDF_EXPORT_REPORT_PREFIX = '[pptx-viewer] PDF export:';

/** Run the PDF export the settings ask for; resolves to the vector result, if one was made. */
export async function runPdfExport(
	options: RunPdfExportOptions,
): Promise<VectorPdfExportResult | undefined> {
	if (options.settings.mode === 'pictures') {
		await options.pictures();
		return undefined;
	}
	const result = await exportSlidesVectorPdf({
		includeHiddenSlides: options.settings.includeHiddenSlides,
		...options.vector,
	});
	await options.save(result.blob);
	if (result.diagnostics.length) {
		// oxlint-disable-next-line no-console -- the honest list of what the file approximates.
		console.info(PDF_EXPORT_REPORT_PREFIX, result.diagnostics);
		options.onApproximated?.(result);
	}
	const toast = pdfExportNoticeToast(result);
	if (toast) options.notify?.(toast);
	return result;
}
