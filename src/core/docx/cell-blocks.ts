import type { Block, TableCell } from './model';
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
