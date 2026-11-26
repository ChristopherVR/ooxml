import {
	fieldProperties,
	headerFooterDesignBlocks,
	type DocumentModel,
	type HeaderFooterDesign,
} from 'ooxml-core/docx';
import type {
	HeaderFooterKind,
	HeaderFooterSlot,
	enableSlot,
	newHeaderFooterId,
	nextPartName,
	withSlot,
} from './header-footer-commands';
import type { currentSectionIndex, sectionsOf } from './section-commands';

/** The helpers of the eager header and footer commands, passed in so they stay in one bundle. */
export interface HeaderFooterTools {
	enableSlot: typeof enableSlot;
	newHeaderFooterId: typeof newHeaderFooterId;
	nextPartName: typeof nextPartName;
	withSlot: typeof withSlot;
	sectionsOf: typeof sectionsOf;
	currentSectionIndex: typeof currentSectionIndex;
}

/**
 * Insert > Header or Footer > a design: puts the design's paragraphs in the selected section's
 * header or footer, creating the part when the document has none. Designs replace the story's
 * content, as picking a gallery design does in Word.
 */
export function withHeaderFooterDesign(
	tools: HeaderFooterTools,
	model: DocumentModel,
	kind: HeaderFooterKind,
	design: HeaderFooterDesign,
	newId: () => string,
	sectionIndex = 0,
	slot: HeaderFooterSlot = 'default',
	labels?: { title: string; author: string; date: string },
): DocumentModel {
	const { enableSlot, nextPartName, withSlot } = tools;
	const base = enableSlot(model, sectionIndex, slot);
	const blocks = headerFooterDesignBlocks(design, {
		newId,
		contentWidthTwips: Math.round(
			(base.page.width - base.page.marginLeft - base.page.marginRight) * 15,
		),
		properties: fieldProperties(base),
		...(labels ? { labels } : {}),
	});
	return withSlot(
		base,
		kind,
		(existing) => ({
			...existing,
			partName: existing?.partName ?? nextPartName(base, kind),
			blocks,
		}),
		sectionIndex,
		slot,
	);
}

/**
 * Insert > Header or Footer > Remove: drops every header (or footer) story of the section, so it
 * has none of its own; the references leave the section's properties when the document is saved.
 */
export function withoutHeaderFooter(
	tools: HeaderFooterTools,
	model: DocumentModel,
	kind: HeaderFooterKind,
	sectionIndex = 0,
): DocumentModel {
	const sections = tools.sectionsOf(model);
	if (!sections[sectionIndex]?.[kind]) return model;
	return {
		...model,
		sections: sections.map((section, index) => {
			if (index !== sectionIndex) return section;
			const { [kind]: _removed, ...rest } = section;
			return rest;
		}),
	};
}
