import type { Block, DocumentModel, Paragraph } from './model';

/** All editable block lists, including section stories and notes. */
export function documentBlockLists(model: DocumentModel): Block[][] {
	return [
		model.blocks,
		...(model.sections ?? []).flatMap((section) =>
			[section.headers, section.footers].flatMap((slots) =>
				Object.values(slots ?? {}).flatMap((content) => (content ? [content.blocks] : [])),
			),
		),
		...[...(model.footnotes ?? []), ...(model.endnotes ?? [])].map((note) => note.blocks),
	];
}

export function mapBlockParagraphs(
	blocks: Block[],
	map: (paragraph: Paragraph) => Paragraph,
): Block[] {
	return blocks.map((block) =>
		block.type === 'paragraph'
			? map(block)
			: {
					...block,
					rows: block.rows.map((row) =>
						row.map((cell) => ({ ...cell, paragraphs: cell.paragraphs.map(map) })),
					),
				},
	);
}

/** Apply the same format operation across the body, tables, headers, footers and notes. */
export function mapDocumentParagraphs(
	model: DocumentModel,
	map: (paragraph: Paragraph) => Paragraph,
): DocumentModel {
	const slots = <T extends object>(value: T): T =>
		Object.fromEntries(
			Object.entries(value as Record<string, { blocks: Block[] } | undefined>).map(
				([key, content]) => [
					key,
					content ? { ...content, blocks: mapBlockParagraphs(content.blocks, map) } : content,
				],
			),
		) as T;
	return {
		...model,
		blocks: mapBlockParagraphs(model.blocks, map),
		...(model.sections
			? {
					sections: model.sections.map((section) => ({
						...section,
						...(section.headers ? { headers: slots(section.headers) } : {}),
						...(section.footers ? { footers: slots(section.footers) } : {}),
					})),
				}
			: {}),
		...(model.footnotes
			? {
					footnotes: model.footnotes.map((note) => ({
						...note,
						blocks: mapBlockParagraphs(note.blocks, map),
					})),
				}
			: {}),
		...(model.endnotes
			? {
					endnotes: model.endnotes.map((note) => ({
						...note,
						blocks: mapBlockParagraphs(note.blocks, map),
					})),
				}
			: {}),
	};
}
