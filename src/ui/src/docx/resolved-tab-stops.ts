import {
	resolveParagraphFormatting,
	type DocumentModel,
	type Paragraph,
	type TabStop,
} from 'ooxml-core/docx';

/**
 * The tab stops a paragraph shows and lays out with: the stops of its style chain and the document
 * defaults, merged with its own `w:tabs` by the core's `mergeTabStops` (a direct `clear` stop
 * removes an inherited one, so none are left in the result). The paragraph's own stops stay the
 * only ones an edit writes.
 */
export function resolvedTabStops(
	paragraph: Pick<Paragraph, 'style' | 'tabStops'>,
	model: DocumentModel | undefined,
): TabStop[] {
	const catalog = model?.paragraphStyles;
	if (!catalog) return (paragraph.tabStops ?? []).filter((stop) => stop.align !== 'clear');
	return resolveParagraphFormatting({ ...paragraph } as Paragraph, catalog).tabStops ?? [];
}

/** The same for a paragraph node's attributes. */
export function resolvedNodeTabStops(
	attrs: Record<string, unknown>,
	model: DocumentModel | undefined,
): TabStop[] {
	const style = typeof attrs.style === 'string' && attrs.style ? attrs.style : undefined;
	const tabStops = (attrs.tabStops as TabStop[] | null) ?? undefined;
	return resolvedTabStops(
		{ ...(style ? { style } : {}), ...(tabStops ? { tabStops } : {}) },
		model,
	);
}
