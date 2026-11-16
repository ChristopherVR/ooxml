// The numbers of a workbook range a chart edit names, such as a custom error-bar source typed in
// the Format pane (`Sheet1!$C$2:$C$5`). Only a single row or column of stored numbers resolves;
// blanks, text, errors, larger blocks and unknown sheets leave the edit refused.
import type { ChartRangeNumbers } from '../../chart/edit/error-bar-source';
import { rangeSize } from '../address';
import { getValue } from '../cells';
import { resolveSparklineRange } from '../layout/sparkline-data';
import type { Workbook, Worksheet } from '../model';

const MAX_CELLS = 10_000;

/** Reads ranges of `workbook`, an unqualified one on `sheet`. */
export function workbookChartRangeNumbers(workbook: Workbook, sheet: Worksheet): ChartRangeNumbers {
	return (formula) => {
		const target = resolveSparklineRange(workbook, sheet, formula);
		if (!target) return undefined;
		const { start, end } = target.range;
		const count = rangeSize(target.range);
		if (count < 1 || count > MAX_CELLS || (start.row !== end.row && start.col !== end.col))
			return undefined;
		const values: number[] = [];
		for (let row = start.row; row <= end.row; row++)
			for (let col = start.col; col <= end.col; col++) {
				const value = getValue(target.sheet, row, col);
				if (typeof value !== 'number' || !Number.isFinite(value)) return undefined;
				values.push(value);
			}
		return values;
	};
}
