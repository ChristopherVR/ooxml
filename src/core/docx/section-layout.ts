import { signedTwipsFromPixels, twips, twipsFromPixels } from '../units/units';
import type { Block, DocumentModel } from './model';
import type { SectionProperties } from './section-model';

/** Section layout without header/footer content, as stored on the editor document for undo. */
export function sectionLayoutJson(sections: SectionProperties[]): string {
	return JSON.stringify(
		sections.map(({ headers: _headers, footers: _footers, ...layout }) => layout),
	);
}

/**
 * Rebuilds sections from the editor document's layout JSON, reattaching header/footer content
 * from the previous model: a section keeps the content of the prior section ending at the same
 * block, and a newly split section takes the content of the section it was split from. Sections
 * whose ending paragraph no longer exists are dropped (deleting a break merges sections).
 */
export function sectionsFromLayout(
	json: string,
	prior: SectionProperties[] | undefined,
	blocks: Block[],
): SectionProperties[] {
	const layout = JSON.parse(json) as SectionProperties[];
	const order = new Map(blocks.map((block, index) => [block.id, index]));
	const priorSections = prior ?? [];
	const endOf = (section: SectionProperties, index: number, list: SectionProperties[]) =>
		index === list.length - 1 ? Number.POSITIVE_INFINITY : (order.get(section.endsAtBlockId) ?? -1);
	const kept = layout.filter(
		(section, index) => index === layout.length - 1 || order.has(section.endsAtBlockId),
	);
	return kept.map((section, index) => {
		const end = endOf(section, index, kept);
		const source =
			priorSections.find((candidate) => candidate.endsAtBlockId === section.endsAtBlockId) ??
			priorSections.find(
				(candidate, position) => endOf(candidate, position, priorSections) >= end,
			) ??
			priorSections.at(-1);
		return {
			...section,
			...(source?.headers ? { headers: source.headers } : {}),
			...(source?.footers ? { footers: source.footers } : {}),
		};
	});
}

/** The document's sections, creating one from the page settings for documents without any. */
export function sectionsOf(model: DocumentModel): SectionProperties[] {
	if (model.sections?.length) return model.sections;
	const { page } = model;
	return [
		{
			endsAtBlockId: model.blocks.at(-1)?.id ?? '',
			type: 'nextPage',
			pageWidthTwips: twipsFromPixels(page.width),
			pageHeightTwips: twipsFromPixels(page.height),
			orientation: page.width > page.height ? 'landscape' : 'portrait',
			marginTopTwips: signedTwipsFromPixels(page.marginTop),
			marginRightTwips: twipsFromPixels(page.marginRight),
			marginBottomTwips: signedTwipsFromPixels(page.marginBottom),
			marginLeftTwips: twipsFromPixels(page.marginLeft),
			headerDistanceTwips: twips(720),
			footerDistanceTwips: twips(720),
			gutterTwips: twips(0),
			columns: { count: 1, spacingTwips: twips(720), equalWidth: true },
		},
	];
}
