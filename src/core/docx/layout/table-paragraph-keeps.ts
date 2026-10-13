import type { LayoutParagraph, LayoutTable } from './input';
import type { RowLayout } from './table-layout';
import { cellParagraphs } from './table-cell-layout';
import { nestedFrameProfile } from './nested-frame-profile';

/** Source-backed admission for the measured ordinary, note-free keep-lines profile. */
export function tableParagraphKeeps(table: LayoutTable): TableParagraphKeeps | undefined {
	const row = table.rows[0];
	if (
		table.rows.length !== 1 ||
		!row ||
		row.cells.length !== 2 ||
		row.heightPx !== undefined ||
		row.heightSpecified ||
		row.cantSplit ||
		row.isHeader ||
		table.widthPx === undefined ||
		!Number.isFinite(table.widthPx) ||
		table.widthPx <= 0 ||
		!nestedFrameProfile(table, table.widthPx) ||
		row.cells.some(
			(cell) =>
				!cell.borders?.top ||
				(cell.gridSpan ?? 1) !== 1 ||
				cell.verticalMerge !== undefined ||
				(cell.verticalAlign && cell.verticalAlign !== 'top') ||
				cell.blocks?.some((block) => block.kind !== 'paragraph'),
		)
	)
		return undefined;
	const paragraphs = row.cells.flatMap(cellParagraphs);
	if (
		paragraphs.some(
			(p) =>
				p.keepNext ||
				p.pageBreakBefore ||
				p.footnotes?.length ||
				p.floats?.length ||
				p.borders ||
				p.shading ||
				p.widowControl !== false ||
				p.lineSpacingRule !== 'exact' ||
				p.lineSpacingTwips === undefined ||
				!Number.isFinite(p.lineSpacingTwips) ||
				p.lineSpacingTwips <= 0 ||
				(p.spacingBeforeTwips ?? 0) !== 0 ||
				(p.spacingAfterTwips ?? 0) !== 0 ||
				p.runs.some((run) => run.footnoteReferenceId !== undefined || run.breakAfter || run.object),
		)
	)
		return undefined;
	const sources = new Map(paragraphs.map((p) => [p.id, p]));
	if (sources.size !== paragraphs.length || !paragraphs.some((p) => p.keepLines)) return undefined;
	return new TableParagraphKeeps(sources);
}

/** Uses existing paragraph boxes and source tokens; it does not lay out or paginate another story. */
export class TableParagraphKeeps {
	constructor(private readonly sources: ReadonlyMap<string, LayoutParagraph>) {}
	/** Before an ordinary row starts, a first kept paragraph can defer the entire row. */
	startupMinimum(row: RowLayout): number {
		return Math.max(
			0,
			...row.cells.map((cell, index) => {
				const first = cell[0];
				return first && this.sources.get(first.blockId)?.keepLines
					? first.yPx + first.heightPx + (row.cellPadding?.[index]?.bottom ?? 0)
					: 0;
			}),
		);
	}
	/** Defer an unstarted paragraph and its following cell content, leaving other cells independent. */
	cuts(row: RowLayout, available: number, fresh: boolean): Map<string, number> | undefined {
		const cuts = new Map<string, number>();
		row.cells.forEach((cell, index) => {
			const padding = row.cellPadding?.[index] ?? { top: 0, bottom: 0 };
			let deferred = false;
			for (const [position, box] of cell.entries()) {
				const source = this.sources.get(box.blockId);
				const atFreshOrigin = fresh && position === 0 && Math.abs(box.yPx - padding.top) < 1e-6;
				if (
					!deferred &&
					source?.keepLines &&
					(box.lines[0]?.firstToken ?? 0) === 0 &&
					box.yPx + box.heightPx + padding.bottom > available + 1e-6 &&
					!atFreshOrigin
				)
					deferred = true;
				if (deferred) cuts.set(box.blockId, 0);
			}
		});
		return cuts.size ? cuts : undefined;
	}
}
