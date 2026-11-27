import { defaultStyleId } from './style-inheritance';
import type { Paragraph, ParagraphStyleCatalog } from './model';

/** What `w:contextualSpacing` compares on a paragraph and its neighbour. */
export interface ContextualSpacingParagraph {
	readonly styleId?: string | undefined;
	readonly contextualSpacing?: boolean | undefined;
}

/** The style id contextual spacing compares: the paragraph's own, else the catalog's default. */
export function contextualStyleId(
	paragraph: Pick<Paragraph, 'style'>,
	catalog: ParagraphStyleCatalog | undefined,
): string | undefined {
	return paragraph.style || (catalog ? defaultStyleId(catalog.styles) : undefined);
}

/**
 * `w:contextualSpacing`: whether `paragraph` drops its own space toward `neighbour` (its space
 * before for the previous paragraph, its space after for the next). Word measures this per
 * paragraph: only a paragraph that sets the flag loses its space, and only beside a paragraph of
 * the same style; the neighbour keeps its own (#359).
 */
export function suppressesContextualSpacing(
	paragraph: ContextualSpacingParagraph,
	neighbour: ContextualSpacingParagraph | undefined,
): boolean {
	if (!paragraph.contextualSpacing || !neighbour?.styleId) return false;
	return neighbour.styleId === paragraph.styleId;
}

/** A paragraph's resolved spacing, as `contextualSpacingTwips` reads it. */
export interface SpacedParagraph extends ContextualSpacingParagraph {
	readonly spacingBeforeTwips?: number | undefined;
	readonly spacingAfterTwips?: number | undefined;
}

/**
 * The space before and after `paragraph` once contextual spacing applies, for a view that
 * collapses adjacent spacing to the larger value (Word's HTML paragraph auto spacing, and CSS
 * margins). Word still subtracts the previous paragraph's own space after from this space before
 * when contextual spacing drops that space after, so a flagged paragraph with 10 pt after leaves
 * no gap above an unflagged one with 10 pt before (measured in Word 16.0.20625, #359).
 * Undefined sides keep the paragraph's own value.
 */
export function contextualSpacingTwips(
	previous: SpacedParagraph | undefined,
	paragraph: SpacedParagraph,
	next: SpacedParagraph | undefined,
): { beforeTwips?: number; afterTwips?: number } {
	const result: { beforeTwips?: number; afterTwips?: number } = {};
	if (suppressesContextualSpacing(paragraph, next)) result.afterTwips = 0;
	if (suppressesContextualSpacing(paragraph, previous)) result.beforeTwips = 0;
	else if (previous && suppressesContextualSpacing(previous, paragraph))
		result.beforeTwips = Math.max(
			0,
			(paragraph.spacingBeforeTwips ?? 0) - (previous.spacingAfterTwips ?? 0),
		);
	return result;
}
