import type { DocumentModel, Paragraph } from '../model';
import { sectionsOf } from '../section-layout';
import type { SectionProperties } from '../section-model';
import { twips } from '../units';

export const INDEX_COLUMN_GAP = twips(360);

function sectionIndexAt(
	sections: SectionProperties[],
	positions: Map<string, number>,
	blockIndex: number,
): number {
	return sections.findIndex(
		(section, index) =>
			index === sections.length - 1 || (positions.get(section.endsAtBlockId) ?? -1) >= blockIndex,
	);
}

export function referenceSectionAt(
	model: DocumentModel,
	blockIndex: number,
): SectionProperties | undefined {
	const sections = sectionsOf(model);
	if (sections.length === 1) return sections[0];
	const positions = new Map(model.blocks.map(({ id }, index) => [id, index]));
	return sections[sectionIndexAt(sections, positions, blockIndex)];
}

/** Columns belong to a continuous section so Word and Print Layout use the same flow. */
export function indexSection(
	model: DocumentModel,
	start: number,
	end: number,
	output: Paragraph[],
	columns: number,
	newId: () => string,
): { paragraphs: Paragraph[]; sections: SectionProperties[] } | undefined {
	if (!output.length) return undefined;
	const sections = sectionsOf(model);
	const positions = new Map(model.blocks.map(({ id }, index) => [id, index]));
	const current = sectionIndexAt(sections, positions, start);
	const section = sections[current];
	if (!section || (columns === 1 && section.columns.count === 1)) return undefined;
	const first = current ? (positions.get(sections[current - 1]?.endsAtBlockId ?? '') ?? -1) + 1 : 0;
	const last =
		current === sections.length - 1
			? model.blocks.length - 1
			: (positions.get(section.endsAtBlockId) ?? model.blocks.length - 1);
	const preceding = model.blocks[start - 1];
	const previous =
		start > first ? (preceding?.type === 'paragraph' ? preceding : blank(newId())) : undefined;
	const following = end >= last && current === sections.length - 1 ? blank(newId()) : undefined;
	const result = output.at(-1);
	if (!result) return undefined;
	const { pageNumbering, ...rest } = section;
	const { start: _restart, ...continued } = pageNumbering ?? {};
	const continuation = { ...rest, type: 'continuous' as const, pageNumbering: continued };
	const index: SectionProperties = {
		...(!previous ? section : continuation),
		endsAtBlockId: result.id,
		columns: { count: columns, spacingTwips: INDEX_COLUMN_GAP, equalWidth: true },
	};
	const next = [...sections];
	next.splice(
		current,
		1,
		...(previous ? [{ ...section, endsAtBlockId: previous.id }] : []),
		index,
		...(end < last || following
			? [{ ...continuation, endsAtBlockId: following?.id ?? section.endsAtBlockId }]
			: []),
	);
	return {
		paragraphs: [
			...(previous && previous !== preceding ? [previous] : []),
			...output,
			...(following ? [following] : []),
		],
		sections: next,
	};
}

const blank = (id: string): Paragraph => ({ type: 'paragraph', id, runs: [] });
