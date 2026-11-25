import type { Block, Paragraph, TableCell } from './model';
/** Source child order, with current direct paragraphs and any new paragraphs beside their source anchor. */
export function cellSourceBlocks(cell: TableCell): Block[] {
	if (!cell.sourceBlocks) return cell.paragraphs;
	const originalIds = new Set(
		cell.sourceBlocks.filter((block) => block.type === 'paragraph').map((block) => block.id),
	);
	const positions = new Map(cell.paragraphs.map((paragraph, index) => [paragraph.id, index]));
	const blocks: Block[] = [];
	let cursor = 0;
	for (const source of cell.sourceBlocks) {
		if (source.type === 'table') {
			blocks.push(source);
			continue;
		}
		const position = positions.get(source.id);
		if (position !== undefined && position >= cursor) {
			blocks.push(...cell.paragraphs.slice(cursor, position + 1));
			cursor = position + 1;
		}
		while (cursor < cell.paragraphs.length && !originalIds.has(cell.paragraphs[cursor]!.id))
			blocks.push(cell.paragraphs[cursor++]!);
	}
	blocks.push(...cell.paragraphs.slice(cursor));
	return blocks;
}

/**
 * The paragraphs of the main document in reading order: body paragraphs, and each table's cells
 * row by row, a cell's paragraphs and the tables nested in it in the order they come. Headers,
 * footers, notes and text boxes are not searched, as in Word.
 */
export function documentParagraphs(blocks: readonly Block[]): Paragraph[] {
	const result: Paragraph[] = [];
	const visit = (items: readonly Block[]) => {
		for (const block of items) {
			if (block.type === 'paragraph') result.push(block);
			else for (const row of block.rows) for (const cell of row) visit(cellSourceBlocks(cell));
		}
	};
	visit(blocks);
	return result;
}

/** Maps paragraphs in story order, including current direct paragraphs and nested tables. */
export function mapStoryBlocks(
	blocks: Block[],
	paragraph: (value: Paragraph) => Paragraph,
	tableId: (id: string) => string = (id) => id,
): Block[] {
	let storyChanged = false;
	const updated = blocks.map((block): Block => {
		if (block.type === 'paragraph') {
			const next = paragraph(block);
			storyChanged ||= next !== block;
			return next;
		}
		let changed = false;
		const rows = block.rows.map((row) =>
			row.map((cell) => {
				const original = cellSourceBlocks(cell);
				const source = mapStoryBlocks(original, paragraph, tableId);
				if (source === original) return cell;
				changed = true;
				return {
					...cell,
					paragraphs: source.filter((item) => item.type === 'paragraph'),
					...(cell.sourceBlocks && { sourceBlocks: source }),
				};
			}),
		);
		const id = tableId(block.id);
		storyChanged ||= changed || id !== block.id;
		return changed || id !== block.id ? { ...block, id, rows } : block;
	});
	return storyChanged ? updated : blocks;
}
