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

/** Records a filter result while preserving manual hiding and other row metadata. */
export function setRowFiltered(
	sheet: Worksheet,
	row: number,
	filteredOut: boolean,
	hasCriteria = false,
): void {
	const info = { ...sheet.rowInfo.get(row) };
	const manual = info.manuallyHidden || (info.hidden && !info.filteredOut);
	if (filteredOut) {
		info.hidden = true;
		info.filteredOut = true;
		if (manual) info.manuallyHidden = true;
	} else {
		delete info.filteredOut;
		delete info.manuallyHidden;
		if (manual) {
			info.hidden = true;
			if (hasCriteria) info.filteredOut = false;
		} else delete info.hidden;
	}
	if (Object.keys(info).length) sheet.rowInfo.set(row, info);
	else sheet.rowInfo.delete(row);
}

/** Removing the filter restores rows hidden only by it, retaining manual hiding. */
export function clearFilteredRows(sheet: Worksheet): void {
	for (const [row, info] of sheet.rowInfo)
		if (info.filteredOut !== undefined || info.manuallyHidden !== undefined)
			setRowFiltered(sheet, row, false);
}

/** Loaded rows have a combined hidden bit; infer the filter cause from supported criteria. */
export function isFilteredRow(workbook: Workbook, sheet: Worksheet, row: number): boolean {
	const filter = sheet.autoFilter;
	if (!filter) return false;
	const info = sheet.rowInfo.get(row);
	if (info?.filteredOut !== undefined) return info.filteredOut;
	return (
		!!info?.hidden &&
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
