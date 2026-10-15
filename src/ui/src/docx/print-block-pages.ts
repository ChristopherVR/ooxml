import type { LayoutBlockBox, LayoutResult } from 'ooxml-core/docx/layout';

/** Source ranges locate the caret's page without requiring offscreen DOM. */
export function printBlockPages(result: LayoutResult) {
	const ranges = new Map<string, { page: number; from: number; to: number }[]>();
	function visit(block: LayoutBlockBox, page: number): void {
		if (block.kind === 'table') {
			for (const row of block.rows)
				for (const cell of row.cells) for (const paragraph of cell) visit(paragraph, page);
			return;
		}
		const first = block.lines[0];
		const last = block.lines.at(-1);
		if (!first || !last) return;
		const list = ranges.get(block.blockId) ?? [];
		list.push({ page, from: first.sourceStart, to: last.sourceEnd });
		ranges.set(block.blockId, list);
	}
	result.pages.forEach((page, index) => {
		for (const column of page.columns) for (const block of column.blocks) visit(block, index);
		// Footnote paragraphs live in the page's note area; their pages locate a note for editing.
		for (const note of page.footnotes ?? [])
			for (const block of note.paragraphs) visit(block, index);
	});
	return (id: string, offset: number): number | undefined => {
		const list = ranges.get(id);
		for (let index = (list?.length ?? 0) - 1; index >= 0; index--) {
			const range = list![index]!;
			if (offset >= range.from && offset <= range.to) return range.page;
		}
		return undefined;
	};
}
