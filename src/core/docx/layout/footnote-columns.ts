import type { LayoutColumns } from './input';
import type { LayoutFootnoteBox, LayoutPageBox } from './result';
import { splitNote } from './footnote-split';
export interface FootnoteColumn {
	xPx: number;
	widthPx: number;
}
export type ReflowFootnote = (note: LayoutFootnoteBox, width: number) => LayoutFootnoteBox;
export function footnoteColumns(page: LayoutPageBox, columns: LayoutColumns): FootnoteColumn[] {
	const width = page.widthPx - page.marginLeftPx - page.marginRightPx;
	const equal = (width - columns.gapPx * (columns.count - 1)) / columns.count;
	let x = 0;
	return Array.from({ length: columns.count }, (_, index) => {
		const own = columns.widths?.[index];
		const column = { xPx: x, widthPx: Math.max(1, own?.widthPx ?? equal) };
		x += column.widthPx + (own?.gapPx ?? columns.gapPx);
		return column;
	});
}
/** Columns flow in reference order; a split within this physical page is not a continued-page mark. */
export function packFootnotes(
	notes: readonly LayoutFootnoteBox[],
	columns: readonly FootnoteColumn[],
	height: number,
	reflow: ReflowFootnote,
) {
	const remaining = [...notes],
		placed: LayoutFootnoteBox[] = [];
	let used = 0;
	for (const [columnIndex, column] of columns.entries()) {
		let y = 0;
		while (remaining.length) {
			const current = reflow(remaining[0]!, column.widthPx);
			const split = splitNote(current, Math.max(0, height - y));
			if (!split.before) break;
			placed.push({
				...split.before,
				yPx: y,
				xPx: column.xPx,
				widthPx: column.widthPx,
				columnIndex,
			});
			y += split.before.heightPx;
			used = Math.max(used, y);
			if (split.after) {
				remaining[0] = { ...split.after, continued: !!current.continued };
				break;
			}
			remaining.shift();
		}
	}
	return { placed, remaining, heightPx: used };
}
/** Minimal balanced height, bounded to 0.01px; the result retains complete source lines. */
export function balancedFootnotes(
	notes: readonly LayoutFootnoteBox[],
	columns: readonly FootnoteColumn[],
	limit: number,
	reflow: ReflowFootnote,
) {
	if (!notes.length) return { placed: [], remaining: [], heightPx: 0 };
	const full = packFootnotes(notes, columns, limit, reflow);
	if (full.remaining.length) return full;
	let lower = 0,
		upper = limit;
	for (let step = 0; step < 24 && upper - lower > 0.005; step++) {
		const middle = (lower + upper) / 2;
		if (packFootnotes(notes, columns, middle, reflow).remaining.length) lower = middle;
		else upper = middle;
	}
	return packFootnotes(notes, columns, Math.ceil(upper * 100) / 100, reflow);
}
export function minimumFootnoteHeight(
	notes: readonly LayoutFootnoteBox[],
	columns: readonly FootnoteColumn[],
	reflow: ReflowFootnote,
): number {
	const width = columns[0]!.widthPx;
	const maximum = notes.reduce((sum, note) => sum + reflow(note, width).heightPx, 0);
	return balancedFootnotes(notes, columns, maximum, reflow).heightPx;
}
