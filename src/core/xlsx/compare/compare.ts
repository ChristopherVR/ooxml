import type { Workbook } from '../model';
import { Budget, compareSheetCells } from './cells';
import { FormatIndex } from './format-key';
import { groupCellDifferences } from './group';
import { matchSheets, sheetDifferences } from './sheets';
import { nameDifferences, sheetStructure } from './structure';
import {
	COMPARE_CATEGORIES,
	NOT_COMPARED,
	type CompareOptions,
	type SheetComparison,
	type WorkbookComparison,
} from './types';

/**
 * What changed from `before` to `after`: sheets, defined names, the structure of each sheet pair
 * and its cells in Spreadsheet Compare's categories (entered values, formulas, calculated values
 * and effective formats). Cells are matched by address, so an inserted row reports every cell
 * below it. Neither workbook is modified. Each stored cell of the two workbooks is visited once
 * (row and column numbers are sorted per sheet and row; no dense grid is built); the result lists
 * `notCompared` parts so a report can say what it did not check.
 */
export function compareWorkbooks(
	before: Workbook,
	after: Workbook,
	options: CompareOptions = {},
): WorkbookComparison {
	const wanted = options.categories;
	const categories = COMPARE_CATEGORIES.filter((c) => !wanted || wanted.includes(c));
	const on = new Set(categories);
	const budget = new Budget(options.maxDifferences ?? Infinity);
	const pairs = matchSheets(before, after, options.sheetMatch ?? 'auto');
	const workbook = [
		...(on.has('sheet') ? budget.admit(sheetDifferences(before, after, pairs)) : []),
		...(on.has('name') ? budget.admit(nameDifferences(before, after, pairs)) : []),
	];
	const ctx = { categories: on, formats: new FormatIndex(), budget };
	const sheets: SheetComparison[] = [];
	for (const pair of pairs) {
		const b = before.sheets[pair.before];
		const a = after.sheets[pair.after];
		if (!b || !a) continue;
		// Once the cap has cut a difference, later sheets are only listed.
		const structure = budget.truncated ? [] : budget.admit(sheetStructure(b, a, on));
		const cells = budget.truncated ? [] : compareSheetCells(ctx, before, b, after, a);
		sheets.push({
			name: a.name,
			beforeName: b.name,
			beforeIndex: pair.before,
			afterIndex: pair.after,
			matchedBy: pair.matchedBy,
			structure,
			cells,
			ranges: groupCellDifferences(cells),
		});
	}
	return {
		workbook,
		sheets,
		count: budget.count,
		truncated: budget.truncated,
		categories,
		notCompared: [...COMPARE_CATEGORIES.filter((c) => !on.has(c)), ...NOT_COMPARED],
	};
}
