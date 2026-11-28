import type { Block, DocumentModel, Paragraph, Table, TableCell } from './model';
import { documentBlockLists } from './document-paragraphs';
import { isEmptyParagraph } from './paragraph-content';
import { isRunHiddenForReview } from './review-formatting-display';

const markDeleted = (paragraph: Paragraph): boolean =>
	paragraph.markRevision?.kind === 'delete' || paragraph.markRevision?.kind === 'moveFrom';
const rowDeleted = (table: Table, index: number): boolean =>
	table.rowProperties?.[index]?.revision?.kind === 'delete';

function hasFinalStructure(model: DocumentModel): boolean {
	const found = (blocks: Block[]): boolean =>
		blocks.some((block) =>
			block.type === 'paragraph'
				? markDeleted(block)
				: block.rows.some(
						(row, index) =>
							rowDeleted(block, index) ||
							row.some((cell) => found(cell.paragraphs) || found(cell.sourceBlocks ?? [])),
					),
		);
	return documentBlockLists(model).some(found);
}

/** Grid column where each cell of a row starts. */
function columns(row: TableCell[]): number[] {
	let column = 0;
	return row.map((cell) => {
		const start = column;
		column += cell.gridSpan ?? 1;
		return start;
	});
}

/**
 * The rows that survive, with a vertical merge whose first row was deleted handed on: the first
 * surviving continuation becomes the merge's start and takes the deleted start cell (its content
 * included), so the merged text is not lost.
 */
function survivingRows(table: Table): { row: TableCell[]; index: number }[] {
	const kept: { row: TableCell[]; index: number }[] = [];
	const pending = new Map<number, TableCell>();
	for (const [index, row] of table.rows.entries()) {
		const starts = columns(row);
		if (rowDeleted(table, index)) {
			for (const [i, cell] of row.entries()) {
				if (cell.verticalMerge === 'restart') pending.set(starts[i]!, cell);
				else if (cell.verticalMerge !== 'continue') pending.delete(starts[i]!);
			}
			continue;
		}
		const next = row.map((cell, i) => {
			const column = starts[i]!;
			const start = pending.get(column);
			pending.delete(column);
			return start && cell.verticalMerge === 'continue'
				? { ...start, verticalMerge: 'restart' as const }
				: cell;
		});
		kept.push({ row: next, index });
	}
	return kept;
}

/** The final view of a table, or nothing when every row was deleted. */
function finalTable(table: Table): Table | undefined {
	const kept = survivingRows(table);
	if (!kept.length) return undefined;
	return {
		...table,
		rows: kept.map(({ row }) =>
			row.map((cell) => ({
				...cell,
				paragraphs: joinDeletedMarks(cell.paragraphs) as Paragraph[],
				...(cell.sourceBlocks ? { sourceBlocks: joinDeletedMarks(cell.sourceBlocks) } : {}),
			})),
		),
		...(table.rowProperties
			? { rowProperties: kept.flatMap(({ index }) => table.rowProperties![index] ?? []) }
			: {}),
	};
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
			const table = finalTable(block);
			if (table) result.push(table);
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
 * The structure of Word's final view (No Markup): rows deleted with tracking are gone, a
 * paragraph whose mark was deleted joins the next paragraph, and a section whose last block went
 * is merged into the next one (Word removes the break). Returns a copy; `model` is untouched.
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
	const blocks = joinDeletedMarks(model.blocks);
	const alive = new Set(blocks.map((block) => block.id));
	const last = blocks.at(-1)?.id;
	const sections = model.sections
		?.flatMap((section, index, all) => {
			if (alive.has(section.endsAtBlockId)) return [section];
			// A section's break went with its last block: the text joins the following section.
			if (index < all.length - 1) return [];
			return [last === undefined ? section : { ...section, endsAtBlockId: last }];
		})
		.map((section) => ({
			...section,
			...(section.headers ? { headers: slots(section.headers) } : {}),
			...(section.footers ? { footers: slots(section.footers) } : {}),
		}));
	return {
		...model,
		blocks,
		...(sections ? { sections } : {}),
		...(model.footnotes ? { footnotes: model.footnotes.map(swap) } : {}),
		...(model.endnotes ? { endnotes: model.endnotes.map(swap) } : {}),
		...(model.footnoteSeparators ? { footnoteSeparators: slots(model.footnoteSeparators) } : {}),
	};
}
