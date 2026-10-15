import type { DocumentModel } from '../index';
import { layoutDocument, layoutDocumentModel } from './layout';
import { layoutDocumentIncremental, type IncrementalMemory } from './incremental-layout';
import { registerParagraphCache } from './paragraph-cache';
import type { LayoutDocumentInput } from './input';
import type { TextMeasurer } from './measure';
import type { LayoutReviewOptions } from './review-display';

export interface LayoutSessionOptions {
	/**
	 * Paginate plain documents incrementally: resume at the first edited page and stop once the
	 * layout matches the previous pass, sharing the untouched page objects between results. The
	 * output equals a full pass, but the results of one session then share pages and must be
	 * treated as read-only. Off by default, which keeps every result independently mutable.
	 */
	incremental?: boolean;
}

/** Reuses measured paragraphs across edits. Create a new session when font metrics change. */
export function createLayoutSession(measurer: TextMeasurer, options: LayoutSessionOptions = {}) {
	const sessionMeasurer: TextMeasurer = {
		widthOf: (text, font) => measurer.widthOf(text, font),
		lineHeightOf: (font) => measurer.lineHeightOf(font),
		...(measurer.ascentOf ? { ascentOf: measurer.ascentOf.bind(measurer) } : {}),
	};
	registerParagraphCache(sessionMeasurer);
	// Plain documents are paginated from the first edited page and only until the layout matches
	// the previous pass again; the output equals a full pass (see `incremental-layout.ts`).
	const memory: IncrementalMemory = {};
	const paginate = options.incremental
		? (input: LayoutDocumentInput, sessionMeasure: TextMeasurer) =>
				layoutDocumentIncremental(input, sessionMeasure, memory)
		: layoutDocument;
	return {
		layoutDocument: (input: LayoutDocumentInput) => paginate(input, sessionMeasurer),
		layoutDocumentModel: (model: DocumentModel, review: LayoutReviewOptions = {}) =>
			layoutDocumentModel(model, sessionMeasurer, review, paginate),
		/** What the last pass did: a full layout, an incremental one, or nothing changed. */
		lastPass: () => memory.stats,
	};
}
