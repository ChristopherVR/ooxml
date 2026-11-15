import type { DocumentModel } from '../model';

/** How finely changed text is marked, as in Word's Compare Documents "Show changes at". */
export type CompareGranularity = 'word' | 'character';

export interface CompareDocumentsOptions {
	/** The author of every revision the comparison records (Word's "Label changes with"). */
	author: string;
	/** The revision date; an ISO 8601 string or a `Date`. Default: now. */
	date?: string | Date;
	/** Default `word`. */
	granularity?: CompareGranularity;
	/** Record formatting differences of unchanged text as `w:rPrChange`; default true. */
	compareFormatting?: boolean;
}

/** Parts of a document the comparison does not look at; listed in every result. */
export const UNCOMPARED_DOCUMENT_PARTS = [
	/** Headers and footers keep the revised document's. */
	'headersAndFooters',
	/** Footnotes and endnotes keep the revised document's; deleted note references are dropped. */
	'notes',
	/** Comments keep the revised document's; comment anchors in deleted text are dropped. */
	'comments',
	/** Text boxes, shapes and drawings are compared as whole objects, never inside. */
	'textBoxes',
	/** Moved text is shown as a deletion and an insertion, never as `w:moveFrom`/`w:moveTo`. */
	'moves',
	/** Paragraph formatting keeps the revised document's; no `w:pPrChange` is recorded. */
	'paragraphFormatting',
	/** Table, row and cell properties keep the revised document's; no `w:tblPrChange`/`w:tcPrChange`. */
	'tableFormatting',
	/** Section and page setup keeps the revised document's. */
	'sections',
] as const;
export type UncomparedDocumentPart = (typeof UNCOMPARED_DOCUMENT_PARTS)[number];

export interface CompareSummary {
	/** Spans of inserted text, paragraph marks and table rows. */
	insertions: number;
	/** Spans of deleted text, paragraph marks and table rows. */
	deletions: number;
	/** Spans of unchanged text whose run formatting changed. */
	formatChanges: number;
}

export interface CompareDocumentsResult {
	/** The new package: the revised document with the differences as tracked changes. */
	bytes: Uint8Array;
	/** `bytes` loaded again, so the model and the package agree. */
	model: DocumentModel;
	summary: CompareSummary;
	/** Content the comparison could not carry over or represent, in plain words. */
	warnings: string[];
	notCompared: readonly UncomparedDocumentPart[];
}
