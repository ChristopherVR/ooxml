import type { Node } from 'prosemirror-model';
import { createDocument, type DocumentModel, type SectionProperties } from '../model';
import { sectionsOf } from '../section-layout';
import { acceptAllRevisions, listRevisions, rejectAllRevisions } from '../revision-commands';
import { sectionPartsJson, SectionPartsStep } from './header-footer-history';

function partModel(doc: Node): DocumentModel {
	const model = { ...createDocument(), blocks: [] };
	if (typeof doc.attrs.sectionParts !== 'string') return model;
	const parts = JSON.parse(doc.attrs.sectionParts) as Pick<
		SectionProperties,
		'endsAtBlockId' | 'headers' | 'footers'
	>[];
	const geometry = sectionsOf(model)[0]!;
	return { ...model, sections: parts.map((part) => ({ ...geometry, ...part })) };
}

export function hasSectionPartRevisions(doc: Node): boolean {
	return listRevisions(partModel(doc)).length > 0;
}

/** Reuse the product revision rules for stored header/footer paragraphs and tables. */
export function resolveSectionPartRevisions(
	doc: Node,
	mode: 'accept' | 'reject',
): SectionPartsStep | undefined {
	const model = partModel(doc);
	if (!listRevisions(model).length) return undefined;
	const next = (mode === 'accept' ? acceptAllRevisions : rejectAllRevisions)(model);
	return new SectionPartsStep(sectionPartsJson(next.sections));
}
