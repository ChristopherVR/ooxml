import type { LayoutTableCell } from './input';
import { nestedFrameProfile } from './nested-frame-profile';
/** Native-measured bounded profile: one automatic, top-aligned nested row, then direct text. */
export function independentNoteStories(cell: LayoutTableCell): Map<string, string> | undefined {
	const first = cell.blocks?.[0];
	if (
		first?.kind !== 'table' ||
		first.rows.length !== 1 ||
		(cell.verticalAlign && cell.verticalAlign !== 'top')
	)
		return undefined;
	if (cell.blocks!.slice(1).some((block) => block.kind !== 'paragraph' || block.footnotes?.length))
		return undefined;
	const row = first.rows[0]!;
	if (
		row.heightPx !== undefined ||
		row.cantSplit ||
		row.isHeader ||
		!nestedFrameProfile(first, first.widthPx)
	)
		return undefined;
	if (
		row.cells.some(
			(inner) =>
				inner.blocks?.some((block) => block.kind === 'table') ||
				inner.paragraphs.some(
					(paragraph) => paragraph.keepLines || paragraph.keepNext || paragraph.pageBreakBefore,
				) ||
				(inner.verticalAlign && inner.verticalAlign !== 'top'),
		)
	)
		return undefined;
	const groups = new Map<string, string>();
	row.cells.forEach((inner, index) => {
		for (const paragraph of inner.paragraphs) groups.set(paragraph.id, first.id + ':' + index);
	});
	return groups;
}
