import type { Block, DocumentModel, Paragraph } from './model';
import { documentBlockLists } from './document-paragraphs';
import { isEmptyParagraph } from './paragraph-content';
import { isRunHiddenForReview } from './review-formatting-display';

const markDeleted = (paragraph: Paragraph): boolean =>
	paragraph.markRevision?.kind === 'delete' || paragraph.markRevision?.kind === 'moveFrom';

function hasFinalStructure(model: DocumentModel): boolean {
	const found = (blocks: Block[]): boolean =>
		blocks.some((block) =>
			block.type === 'paragraph'
				? markDeleted(block)
				: block.rowProperties?.some((row) => row.revision?.kind === 'delete') ||
					block.rows.some((row) => row.some((cell) => found(cell.paragraphs))),
		);
	return documentBlockLists(model).some(found);
}

/**
 * A list with its deleted paragraph marks joined to the next paragraph, as Word's final view
 * shows. The last paragraph of a list keeps its mark: a cell's last mark ends the cell.
 */
function joinDeletedMarks(blocks: Block[]): Block[] {
	const result: Block[] = [];
	let carried: Paragraph | undefined;
	for (const [index, block] of blocks.entries()) {
		if (block.type === 'table') {
			if (carried && !isEmptyParagraph(carried)) result.push(carried);
			carried = undefined;
			const kept = block.rows.flatMap((row, rowIndex) =>
				block.rowProperties?.[rowIndex]?.revision?.kind === 'delete' ? [] : [{ row, rowIndex }],
			);
			if (!kept.length) continue;
			result.push({
				...block,
				rows: kept.map(({ row }) =>
					row.map((cell) => ({
						...cell,
						paragraphs: joinDeletedMarks(cell.paragraphs) as Paragraph[],
					})),
				),
				...(block.rowProperties
					? { rowProperties: kept.flatMap(({ rowIndex }) => block.rowProperties![rowIndex] ?? []) }
					: {}),
			});
			continue;
		}
		const paragraph: Paragraph = carried
			? { ...block, runs: [...carried.runs, ...block.runs] }
			: block;
		carried = undefined;
		if (markDeleted(paragraph) && index < blocks.length - 1) {
			carried = {
				...paragraph,
				runs: paragraph.runs.filter((run) => !isRunHiddenForReview(run, 'final')),
			};
			continue;
		}
		result.push(paragraph);
	}
	return result;
}

/**
 * The structure of Word's final view (No Markup): rows deleted with tracking are gone and a
 * paragraph whose mark was deleted joins the next paragraph. Returns a copy; `model` is untouched.
 * Source offsets change, so only read-only outputs such as PDF export use it.
 */
export function finalReviewStructure(model: DocumentModel): DocumentModel {
	if (!hasFinalStructure(model)) return model;
	const swap = <T extends { blocks: Block[] }>(story: T): T => ({
		...story,
		blocks: joinDeletedMarks(story.blocks),
	});
	const slots = <T extends object>(value: T): T =>
		Object.fromEntries(
			Object.entries(value as Record<string, { blocks: Block[] } | undefined>).map(
				([key, content]) => [key, content ? swap(content) : content],
			),
		) as T;
	return {
		...model,
		blocks: joinDeletedMarks(model.blocks),
		...(model.sections
			? {
					sections: model.sections.map((section) => ({
						...section,
						...(section.headers ? { headers: slots(section.headers) } : {}),
						...(section.footers ? { footers: slots(section.footers) } : {}),
					})),
				}
			: {}),
		...(model.footnotes ? { footnotes: model.footnotes.map(swap) } : {}),
		...(model.endnotes ? { endnotes: model.endnotes.map(swap) } : {}),
		...(model.footnoteSeparators ? { footnoteSeparators: slots(model.footnoteSeparators) } : {}),
	};
}
