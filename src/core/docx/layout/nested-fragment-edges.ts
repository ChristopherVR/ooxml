import type { LayoutParagraph, LayoutTable, LayoutTableRow } from './input';
import { nestedFrameProfile } from './nested-frame-profile';
const plain = (paragraph: LayoutParagraph) =>
	!paragraph.keepLines &&
	!paragraph.keepNext &&
	!paragraph.pageBreakBefore &&
	!paragraph.footnotes?.length &&
	!paragraph.runs.some((run) => run.footnoteReferenceId !== undefined || run.breakAfter);
/** Two or more automatic top-aligned rows with direct paragraphs and complete shared edges. */
export function nestedFragmentEdges(table: LayoutTable) {
	if (
		table.widthPx === undefined ||
		!Number.isFinite(table.widthPx) ||
		table.widthPx <= 0 ||
		table.rows.length < 2 ||
		!nestedFrameProfile(table, table.widthPx) ||
		table.rows.some(
			(row) =>
				row.heightPx !== undefined ||
				row.cantSplit ||
				row.isHeader ||
				row.cells.length !== 2 ||
				row.cells.some(
					(cell) =>
						(cell.gridSpan ?? 1) !== 1 ||
						cell.verticalMerge !== undefined ||
						(cell.verticalAlign && cell.verticalAlign !== 'top') ||
						cell.blocks?.some((block) => block.kind !== 'paragraph') ||
						!(cell.blocks ?? cell.paragraphs).every(
							(block) => block.kind === 'paragraph' && plain(block),
						),
				),
		)
	)
		return undefined;
	const top = table.rows[0]!.cells.map((cell) => cell.borders?.top);
	const bottom = table.rows.at(-1)!.cells.map((cell) => cell.borders?.bottom);
	return top.every(Boolean) && bottom.every(Boolean) ? { top, bottom } : undefined;
}
/** The enclosing story must also remain within the measured automatic, note-free profile. */
export function reflowNestedFragments(row: LayoutTableRow): boolean {
	if (row.heightPx !== undefined || row.cantSplit || row.isHeader || row.cells.length !== 2)
		return false;
	let tables = 0;
	for (const cell of row.cells) {
		if ((cell.gridSpan ?? 1) !== 1 || cell.verticalMerge !== undefined) return false;
		if (cell.verticalAlign && cell.verticalAlign !== 'top') return false;
		for (const [index, block] of (cell.blocks ?? cell.paragraphs).entries()) {
			if (block.kind === 'paragraph') {
				if (!plain(block)) return false;
			} else {
				if (index !== 0 || !nestedFragmentEdges(block) || ++tables > 1) return false;
			}
		}
	}
	return tables === 1;
}
