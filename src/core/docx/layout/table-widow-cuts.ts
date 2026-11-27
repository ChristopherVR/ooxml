import type { LayoutTableRow } from './input';
import { adjustForWidowOrphan, widowControlEnabled } from './keep-rules';
import type { RowLayout } from './table-layout';
import { cellParagraphs } from './table-cell-layout';

/**
 * The height a row split at `available` (bottom padding included, as `splitRowAtHeight` takes it)
 * is cut at so that Word's widow and orphan control holds in its cells, or `undefined` when the
 * rule changes nothing. Word applies the rule there as in the body, and cuts every cell of the row
 * at the same height: measured with Word 16 on a three-page table, the row whose two-line
 * description no longer fit moved to the next page whole, its one-line cells too, and with
 * `w:widowControl w:val="0"` the same row split after the description's first line.
 */
export function widowControlCut(
	source: LayoutTableRow,
	row: RowLayout,
	available: number,
): number | undefined {
	const paragraphs = new Map(
		source.cells.flatMap(cellParagraphs).map((paragraph) => [paragraph.id, paragraph] as const),
	);
	let lowest: number | undefined;
	row.cells.forEach((cell, index) => {
		const padding = row.cellPadding?.[index]?.bottom ?? 0;
		const cut = available - padding;
		for (const box of cell) {
			const bottom = box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0;
			const fitting = box.lines.filter(
				(line) => box.yPx + line.yPx + line.heightPx + bottom <= cut,
			).length;
			if (fitting === 0 || fitting >= box.lines.length) continue;
			const paragraph = paragraphs.get(box.blockId);
			if (!paragraph) continue;
			// A continuation (the paragraph already started on an earlier page) has no orphan to keep.
			const continued = (box.lines[0]?.firstToken ?? 0) > 0 || !!box.lines[0]?.firstTokenOffset;
			const enabled = widowControlEnabled(paragraph.widowControl);
			let kept = adjustForWidowOrphan(box.lines.length, fitting, enabled);
			if (continued && kept === 0 && fitting > 0)
				kept = box.lines.length - fitting === 1 ? Math.max(1, fitting - 1) : fitting;
			if (kept === fitting) continue;
			const last = box.lines[kept - 1];
			// Above the first line that has to move: keep the lines before it, or none of the box.
			const at = (last ? box.yPx + last.yPx + last.heightPx + bottom : box.yPx) + padding;
			lowest = Math.min(lowest ?? Infinity, last ? at + 1e-6 : at - 1e-6);
		}
	});
	return lowest;
}
