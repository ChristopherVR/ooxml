/**
 * The choices of File > Export > Create a PDF's Options step (Word's "Publish as PDF" Options
 * dialog) and the export options they stand for. DOM-free, so the mapping is tested on its own.
 */
import type { EditorPdfOptions } from './editor-pdf';
import type { DocxPdfMarkupBalloons } from 'ooxml-core/docx/pdf-export';

export type PdfRangeChoice = 'all' | 'current' | 'selection' | 'pages';

export interface PdfOptionsChoice {
	range: PdfRangeChoice;
	/** The Page(s) From and To fields, as typed. */
	from: string;
	to: string;
	/** Publish what: Document showing markup (else Document). */
	markup: boolean;
	/** Word's Review > Show Markup > Balloons, for the markup shown. */
	balloons: DocxPdfMarkupBalloons;
	/** Create bookmarks using: off writes no outline. */
	createBookmarks: boolean;
	bookmarkSource: 'headings' | 'wordBookmarks';
	documentProperties: boolean;
	taggedPdf: boolean;
	pdfA: boolean;
}

/** What the editor knows when the step opens. */
export interface PdfOptionsContext {
	/** Print Layout's current page and page count; undefined in Draft view. */
	pages?: { current: number; total: number } | undefined;
	/** Whether body text is selected. */
	hasSelection: boolean;
}

/** The choices that give exactly what `exportPdf()` writes without options. */
export function defaultPdfChoice(): PdfOptionsChoice {
	return {
		range: 'all',
		from: '1',
		to: '1',
		markup: false,
		balloons: 'commentsAndFormatting',
		createBookmarks: true,
		bookmarkSource: 'headings',
		documentProperties: true,
		taggedPdf: true,
		pdfA: false,
	};
}

/** A page field's whole number from 1, or null. */
function pageNumber(text: string): number | null {
	const value = Number(text.trim());
	return text.trim() && Number.isInteger(value) && value >= 1 ? value : null;
}

/**
 * The export options the choices stand for, or the field to correct: `from` when it is not a page
 * of the document, `to` when it is before `from`. A range the context cannot offer (the current
 * page in Draft view, an empty selection) falls back to the whole document.
 */
export function pdfOptionsFromChoice(
	choice: PdfOptionsChoice,
	context: PdfOptionsContext,
): EditorPdfOptions | { invalid: 'from' | 'to' } {
	const options: EditorPdfOptions = {
		markup: choice.markup ? 'all' : 'none',
		...(choice.markup ? { markupBalloons: choice.balloons } : {}),
		bookmarks: choice.createBookmarks ? choice.bookmarkSource : 'none',
		documentProperties: choice.documentProperties,
		taggedPdf: choice.taggedPdf,
		pdfA: choice.pdfA,
	};
	if (choice.range === 'current' && context.pages) options.pages = 'current';
	else if (choice.range === 'selection' && context.hasSelection) options.pages = 'selection';
	else if (choice.range === 'pages') {
		const from = pageNumber(choice.from);
		const total = context.pages?.total;
		if (from === null || (total !== undefined && from > total)) return { invalid: 'from' };
		const to = pageNumber(choice.to);
		if (to === null || to < from) return { invalid: 'to' };
		options.pages = { from, to: total === undefined ? to : Math.min(to, total) };
	} else options.pages = 'all';
	return options;
}
