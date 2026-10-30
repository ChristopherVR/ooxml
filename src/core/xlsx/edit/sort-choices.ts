// The colours and icons a sort level can pick from: those present in one column (or row) of the
// range, conditional formatting included. Backs the colour and icon pickers of Data > Sort.
import { type CellRange, normalizeRange } from '../address';
import { getCell } from '../cells';
import { cellFillColor, cellFontColor } from '../filter-criteria';
import { createConditionalFormatEvaluator } from '../layout/cf-evaluator';
import type { IconView } from '../layout/types';
import type { CellValue, Workbook } from '../model';

export interface SortChoices {
	/** `#RRGGBB` fills in use, sorted. */
	fills: string[];
	/** Some cell has no fill ("No Cell Color"). */
	noFill: boolean;
	fonts: string[];
	icons: IconView[];
	/** Some cell shows no icon ("No Cell Icon"). */
	noIcon: boolean;
}

const MAX_CELLS = 20000;

/**
 * Colours and icons in the `key` column of `range` (or the `key` row when `leftToRight`),
 * skipping the header. `evaluate` computes conditional-format formulas.
 */
export function sortChoices(
	workbook: Workbook,
	sheetIndex: number,
	evaluate: (formula: string, at: { sheet: number; row: number; col: number }) => CellValue,
	range: CellRange,
	key: number,
	hasHeader: boolean,
	leftToRight = false,
): SortChoices {
	const sheet = workbook.sheets[sheetIndex];
	const out: SortChoices = { fills: [], noFill: false, fonts: [], icons: [], noIcon: false };
	if (!sheet) return out;
	const r = normalizeRange(range);
	let cf: ReturnType<typeof createConditionalFormatEvaluator> | undefined;
	if (sheet.conditionalFormats.length) {
		try {
			cf = createConditionalFormatEvaluator(workbook, sheetIndex, evaluate);
		} catch {
			cf = undefined;
		}
	}
	const fills = new Set<string>();
	const fonts = new Set<string>();
	const icons = new Map<string, IconView>();
	const first = leftToRight ? r.start.col + (hasHeader ? 1 : 0) : r.start.row + (hasHeader ? 1 : 0);
	const last = Math.min(leftToRight ? r.end.col : r.end.row, first + MAX_CELLS);
	for (let pos = first; pos <= last; pos++) {
		const [row, col] = leftToRight ? [key, pos] : [pos, key];
		const cell = getCell(sheet, row, col);
		const format = cf?.at(row, col);
		const fill = cellFillColor(workbook, cell, format);
		if (fill) fills.add(fill);
		else out.noFill = true;
		fonts.add(cellFontColor(workbook, cell, format));
		const icon = format?.icon;
		if (icon) icons.set(`${icon.set}:${icon.index}`, icon);
		else out.noIcon = true;
	}
	out.fills = [...fills].sort();
	out.fonts = [...fonts].sort();
	out.icons = [...icons.values()].sort((a, b) =>
		a.set === b.set ? a.index - b.index : a.set.localeCompare(b.set),
	);
	return out;
}
