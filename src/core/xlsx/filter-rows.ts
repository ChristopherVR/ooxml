import { getCell } from './cells';
import { displayText } from './display-text';
import type { AutoFilter, Workbook, Worksheet } from './model';

/** Matches the value/blank filters supported by the editor. */
export function rowMatchesFilter(
	workbook: Workbook,
	sheet: Worksheet,
	row: number,
	filter: AutoFilter,
): boolean {
	return (filter.columns ?? []).every((column) => {
		if (!column.values && !column.blank) return true;
		const text = displayText(workbook, getCell(sheet, row, filter.range.start.col + column.offset));
		return text === '' ? !!column.blank : (column.values ?? []).includes(text);
	});
}

/** Loaded rows have a combined hidden bit; infer the filter cause from supported criteria. */
export function isFilteredRow(workbook: Workbook, sheet: Worksheet, row: number): boolean {
	const info = sheet.rowInfo.get(row);
	if (info?.filteredOut !== undefined) return info.filteredOut;
	const filter = sheet.autoFilter;
	return (
		!!info?.hidden &&
		!!filter &&
		row > filter.range.start.row &&
		row <= filter.range.end.row &&
		!rowMatchesFilter(workbook, sheet, row, filter)
	);
}

/** Capture loaded filter state so editing a value does not implicitly reapply the filter. */
export function readFilteredRows(workbook: Workbook): void {
	for (const sheet of workbook.sheets)
		for (const [row, info] of sheet.rowInfo)
			if (
				sheet.autoFilter &&
				row > sheet.autoFilter.range.start.row &&
				row <= sheet.autoFilter.range.end.row &&
				info.hidden
			)
				info.filteredOut = isFilteredRow(workbook, sheet, row);
}
