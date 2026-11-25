import { formatRange, quoteSheetName } from '../address';
import { sparklineDataRange } from '../edit/sparkline-ranges';
import type { Workbook } from '../model';

/** An explicit local range or a workbook table's header and body (excluding its totals row). */
export function pivotSourceRange(workbook: Workbook, sheet: number, source: string) {
	const name = source.trim().toLowerCase();
	if (!/[:!$]/.test(name))
		for (const [index, ws] of workbook.sheets.entries()) {
			const table = ws.tables.find(
				(entry) => entry.name.toLowerCase() === name || entry.displayName.toLowerCase() === name,
			);
			if (!table) continue;
			if (!table.headerRow)
				throw new Error('Every PivotTable source column needs a unique, non-empty header.');
			const { start, end } = table.range;
			return {
				sheet: index,
				range: { start, end: { ...end, row: end.row - Number(table.totalsRow) } },
				tableName: table.name,
				source: table.name,
			};
		}
	const range = sparklineDataRange(workbook, sheet, source);
	return {
		...range,
		source: `${quoteSheetName(workbook.sheets[range.sheet]!.name)}!${formatRange(range.range)}`,
	};
}
