import { cellContentBottom } from './table-content';
import { clipTableFrames } from './table-frames';
import type { LayoutParagraphBox } from './result';
import type { RowLayout } from './table-layout';
import { expectDefined } from '../expect-defined';

interface CellSplit {
	before: LayoutParagraphBox[];
	after: LayoutParagraphBox[];
	beforeHeightPx: number;
	afterOrigin?: number;
}

/** Splits one cell's already-laid-out paragraph boxes at `cutHeightPx`, never mid-line. */
function splitCellAtHeight(
	boxes: LayoutParagraphBox[],
	cutHeightPx: number,
	paragraphCuts?: ReadonlyMap<string, number>,
): CellSplit {
	const before: LayoutParagraphBox[] = [];
	const after: LayoutParagraphBox[] = [];
	let beforeHeightPx = 0;
	for (const box of boxes) {
		const cut = paragraphCuts?.get(box.blockId) ?? cutHeightPx;
		if (
			box.yPx + box.heightPx + (box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0) <=
			cut
		) {
			before.push(box);
			beforeHeightPx = Math.max(
				beforeHeightPx,
				box.yPx + box.heightPx + (box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0),
			);
			continue;
		}
		const beforeLines = box.lines.filter(
			(line) =>
				box.yPx +
					line.yPx +
					line.heightPx +
					(box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0) <=
				cut,
		);
		const afterLines = box.lines.slice(beforeLines.length);
		if (beforeLines.length) {
			const bottom = beforeLines.at(-1)!.yPx + beforeLines.at(-1)!.heightPx;
			before.push({ ...box, heightPx: bottom, lines: beforeLines });
			beforeHeightPx = Math.max(
				beforeHeightPx,
				box.yPx + bottom + (box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0),
			);
		}
		if (afterLines.length) {
			const shift = expectDefined(afterLines[0], 'first line after split').yPx;
			after.push({
				...box,
				yPx: box.yPx + shift,
				heightPx: afterLines.at(-1)!.yPx + afterLines.at(-1)!.heightPx - shift,
				lines: afterLines.map((line) => ({ ...line, yPx: line.yPx - shift })),
			});
		}
	}
	// Nested cells have paragraphs running in parallel, rather than a single sequential story.
	const afterOrigin = Math.min(...after.map((box) => box.yPx - (box.tableFrameInsets?.top ?? 0)));
	for (const box of after) box.yPx -= afterOrigin;
	return { before, after, beforeHeightPx, ...(after.length ? { afterOrigin } : {}) };
}

export interface RowSplit {
	before: RowLayout;
	after: RowLayout | null;
}

/**
 * Splits a row that does not fit in `cutHeightPx` of remaining page/column
 * space, including bottom padding. Cells normally share the same cut; optional
 * individual cuts defer a marker without truncating neighboring unreferenced cells.
 * Paragraph-specific cuts are used for independently continued inner stories; their tail
 * must be source-reflowed by the table flow before placement to restore independent origins.
 * Splits never cut a line, and continuation fragments retain their top and bottom padding.
 * Returns `after: null` when every cell fit (nothing left to continue).
 */
export function splitRowAtHeight(
	row: RowLayout,
	cutHeightPx: number,
	cellCuts?: readonly number[],
	paragraphCuts?: ReadonlyMap<string, number>,
): RowSplit {
	const splits = row.cells.map((cell, index) => {
		const padding = row.cellPadding?.[index] ?? { top: 0, bottom: 0 };
		const cuts = paragraphCuts
			? new Map(
					cell.flatMap((box) => {
						const cut = paragraphCuts.get(box.blockId);
						return cut === undefined ? [] : [[box.blockId, cut - padding.bottom] as const];
					}),
				)
			: undefined;
		const split = splitCellAtHeight(
			cell,
			(cellCuts?.[index] ?? cutHeightPx) - padding.bottom,
			cuts,
		);
		if (split.before.length) split.beforeHeightPx += padding.bottom;
		if (split.after.length)
			split.after = split.after.map((box) => ({ ...box, yPx: box.yPx + padding.top }));
		return split;
	});
	const anyAfter = splits.some((split) => split.after.length > 0);
	const beforeHeightPx = Math.max(0, ...splits.map((split) => split.beforeHeightPx));
	const before: RowLayout = {
		heightPx: beforeHeightPx,
		cells: splits.map((split) => split.before),
		...(row.cellTables
			? {
					cellTables: row.cellTables.map((tables, index) =>
						clipTableFrames(
							tables,
							0,
							splits[index]!.beforeHeightPx - (row.cellPadding?.[index]?.bottom ?? 0),
						),
					),
				}
			: {}),
		geometry: row.geometry.map(({ mergeRows: _merge, ...geometry }, index) => ({
			...geometry,
			contentHeightPx: splits[index]!.beforeHeightPx,
			...(row.borderBottomPx !== undefined ? { alignmentHeightPx: beforeHeightPx } : {}),
		})),
		...(row.cellPadding ? { cellPadding: row.cellPadding } : {}),
		...(row.borderBottomPx !== undefined ? { borderBottomPx: row.borderBottomPx } : {}),
	};
	if (!anyAfter) return { before, after: null };
	// The tail of a cell story ends with the space after its last paragraph.
	const tail = (split: CellSplit, index: number) =>
		split.after.length ? (row.trailingPx?.[index] ?? 0) : 0;
	const afterHeightPx = Math.max(
		0,
		...splits.map(
			(split, index) =>
				cellContentBottom(split.after, row.cellPadding?.[index]?.bottom ?? 0) + tail(split, index),
		),
	);
	const after: RowLayout = {
		heightPx: afterHeightPx,
		cells: splits.map((split) => split.after),
		...(row.cellTables
			? {
					cellTables: row.cellTables.map((tables, index) =>
						splits[index]!.after.length
							? clipTableFrames(
									tables,
									splits[index]!.afterOrigin ?? 0,
									row.heightPx,
									row.cellPadding?.[index]?.top ?? 0,
								)
							: [],
					),
				}
			: {}),
		geometry: row.geometry.map(({ mergeRows: _merge, ...geometry }, index) => ({
			...geometry,
			contentHeightPx:
				cellContentBottom(splits[index]!.after, row.cellPadding?.[index]?.bottom ?? 0) +
				tail(splits[index]!, index),
			...(row.borderBottomPx !== undefined ? { alignmentHeightPx: afterHeightPx } : {}),
		})),
		...(row.cellPadding ? { cellPadding: row.cellPadding } : {}),
		...(row.trailingPx ? { trailingPx: row.trailingPx } : {}),
		...(row.borderBottomPx !== undefined ? { borderBottomPx: row.borderBottomPx } : {}),
	};
	return { before, after };
}
