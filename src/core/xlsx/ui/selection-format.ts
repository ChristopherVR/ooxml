import { normalizeRange, styleAt, type CellRange, type Workbook } from '../index';
import { effectiveStyleId } from '../layout/cell-view';

/** Flat formatting for host controls. null means mixed, or unexamined in a large selection. */
export function selectionFormatting(
	workbook: Workbook | undefined,
	sheetIndex: number,
	ranges: readonly CellRange[],
) {
	const values = {
		bold: new Set<boolean>(),
		italic: new Set<boolean>(),
		underline: new Set<boolean>(),
		fontFamily: new Set<string>(),
		fontSize: new Set<number>(),
		alignment: new Set<string>(),
		numberFormat: new Set<string>(),
	};
	const sheet = workbook?.sheets[sheetIndex];
	let remaining = 4096;
	let complete = true;
	if (workbook && sheet)
		outer: for (const raw of ranges) {
			const range = normalizeRange(raw);
			for (let row = range.start.row; row <= range.end.row; row++) {
				for (let col = range.start.col; col <= range.end.col; col++) {
					if (--remaining < 0) {
						complete = false;
						break outer;
					}
					const style = styleAt(workbook, effectiveStyleId(sheet, row, col));
					values.bold.add(Boolean(style.font.bold));
					values.italic.add(Boolean(style.font.italic));
					values.underline.add(Boolean(style.font.underline));
					values.fontFamily.add(style.font.name ?? 'Aptos');
					values.fontSize.add(style.font.size ?? 11);
					values.alignment.add(style.alignment?.horizontal ?? 'general');
					values.numberFormat.add(style.numFmt ?? 'General');
				}
			}
		}
	const single = <T>(set: Set<T>): T | null => (complete && set.size === 1 ? [...set][0]! : null);
	return {
		bold: single(values.bold),
		italic: single(values.italic),
		underline: single(values.underline),
		fontFamily: single(values.fontFamily),
		fontSize: single(values.fontSize),
		alignment: single(values.alignment),
		numberFormat: single(values.numberFormat),
		formattingComplete: complete,
	};
}
