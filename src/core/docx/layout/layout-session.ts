import type { DocumentModel } from '../index';
import { layoutDocument, layoutDocumentModel } from './layout';
import { registerParagraphCache } from './paragraph-cache';
import type { LayoutDocumentInput } from './input';
import type { TextMeasurer } from './measure';
import type { LayoutReviewOptions } from './review-display';

/** Reuses measured paragraphs across edits. Create a new session when font metrics change. */
export function createLayoutSession(measurer: TextMeasurer) {
	const sessionMeasurer: TextMeasurer = {
		widthOf: (text, font) => measurer.widthOf(text, font),
		lineHeightOf: (font) => measurer.lineHeightOf(font),
		...(measurer.ascentOf ? { ascentOf: measurer.ascentOf.bind(measurer) } : {}),
	};
	registerParagraphCache(sessionMeasurer);
	return {
		layoutDocument: (input: LayoutDocumentInput) => layoutDocument(input, sessionMeasurer),
		layoutDocumentModel: (model: DocumentModel, options: LayoutReviewOptions = {}) =>
			layoutDocumentModel(model, sessionMeasurer, options),
	};
}
