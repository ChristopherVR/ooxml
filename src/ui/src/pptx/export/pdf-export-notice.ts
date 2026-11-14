/**
 * The one notice every binding shows after a vector PDF export that approximated something,
 * like the summary PowerPoint gives: counts first ("3 elements were saved as pictures; 1 font
 * was substituted"), each distinct approximation under Details. Built as a compatibility toast
 * (`CompatibilityWarningToast`) so it goes through the binding's existing toast stack, in the
 * viewer's language. No notice when nothing was approximated. Framework-free.
 */
import type {
	CompatibilityWarningToast,
	ToastMessagePart,
} from '../render/compatibility-warning-toasts';
import {
	pdfDiagnosticCategory,
	pdfDiagnosticMessageKey,
	type PdfExportDiagnostic,
} from './pdf-export-diagnostics';

/** The parts of a vector export result the notice reads. */
export interface PdfExportReportSource {
	readonly slides: readonly {
		readonly mode: 'vector' | 'picture';
		readonly pictures?: number;
		readonly diagnostics: readonly PdfExportDiagnostic[];
	}[];
	readonly diagnostics: readonly PdfExportDiagnostic[];
}

/** What a vector PDF approximated, counted. */
export interface PdfExportReportSummary {
	/** Elements drawn as pictures of just themselves. */
	pictures: number;
	/** Slides drawn as one picture each. */
	slidesAsPictures: number;
	/** Distinct font families written with another font. */
	fonts: number;
	/** Things left out of the file, per slide. */
	omitted: number;
	/** Effects drawn simpler than on the slide, per slide. */
	simplified: number;
}

/** Toast id of the PDF notice: a new export replaces the previous notice. */
export const PDF_EXPORT_NOTICE_ID = 'pdf-export-report';
/** Toast code of the PDF notice (mirrored to `data-code`). */
export const PDF_EXPORT_NOTICE_CODE = 'PDF_EXPORT_APPROXIMATED';

/** Count what `result` approximated. */
export function summarizePdfExport(result: PdfExportReportSource): PdfExportReportSummary {
	const summary: PdfExportReportSummary = {
		pictures: 0,
		slidesAsPictures: 0,
		fonts: 0,
		omitted: 0,
		simplified: 0,
	};
	const families = new Set<string>();
	for (const slide of result.slides) {
		if (slide.mode === 'picture') summary.slidesAsPictures++;
		else summary.pictures += slide.pictures ?? 0;
		for (const diagnostic of slide.diagnostics) {
			const category = pdfDiagnosticCategory(diagnostic.code);
			if (category === 'font') families.add(diagnostic.params?.family ?? '');
			else if (category === 'omitted') summary.omitted++;
			else if (category === 'simplified') summary.simplified++;
		}
	}
	summary.fonts = families.size;
	return summary;
}

/** Whether `summary` has anything to tell the user. */
export function hasPdfApproximations(summary: PdfExportReportSummary): boolean {
	return Object.values(summary).some((count) => count > 0);
}

/** `pptx.pdfExport.summary.<name>.one` or `.other`, with `{{count}}`. */
function countPart(name: keyof PdfExportReportSummary, count: number): ToastMessagePart {
	return {
		key: `pptx.pdfExport.summary.${name}.${count === 1 ? 'one' : 'other'}`,
		params: { count: String(count) },
	};
}

const SUMMARY_ORDER: readonly (keyof PdfExportReportSummary)[] = [
	'pictures',
	'slidesAsPictures',
	'fonts',
	'omitted',
	'simplified',
];

/** The translatable count sentences of `summary`, in notice order (non-zero counts only). */
export function pdfExportSummaryParts(summary: PdfExportReportSummary): ToastMessagePart[] {
	return SUMMARY_ORDER.filter((name) => summary[name] > 0).map((name) =>
		countPart(name, summary[name]),
	);
}

/**
 * The translatable Details line of one diagnostic. A font diagnostic without a family is the
 * default font (`...DefaultFamily` keys).
 */
export function pdfDiagnosticMessage(diagnostic: PdfExportDiagnostic): ToastMessagePart {
	const key = pdfDiagnosticMessageKey(diagnostic.code);
	if (pdfDiagnosticCategory(diagnostic.code) === 'font' && !diagnostic.params?.family)
		return {
			key: `${key}DefaultFamily`,
			params: { substitute: diagnostic.params?.substitute ?? '' },
		};
	return diagnostic.params ? { key, params: diagnostic.params } : { key };
}

/**
 * The toast a binding shows after a vector PDF export, or `undefined` when nothing was
 * approximated (no notice then).
 */
export function pdfExportNoticeToast(
	result: PdfExportReportSource | undefined,
): CompatibilityWarningToast | undefined {
	if (!result) return undefined;
	const summary = summarizePdfExport(result);
	if (!hasPdfApproximations(summary)) return undefined;
	return {
		id: PDF_EXPORT_NOTICE_ID,
		code: PDF_EXPORT_NOTICE_CODE,
		severity: 'info',
		messageKey: 'pptx.pdfExport.summary.message',
		parts: pdfExportSummaryParts(summary),
		details: result.diagnostics.map(pdfDiagnosticMessage),
	};
}
