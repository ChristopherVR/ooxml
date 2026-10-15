import { displayText } from '../display-text';
import { getCell } from '../cells';
import { rowMatchesFilter, setRowFiltered } from '../filter-rows';
import type { AutoFilter, Table, TableColumnFilter, Workbook, Worksheet } from '../model';
import { escapeAttr } from '../write/xml-out';

/** The table's data rows' range (header and totals rows excluded from filtering). */
export function tableFilterBounds(table: Table) {
	return {
		start: table.range.start,
		end: { ...table.range.end, row: table.range.end.row - (table.totalsRow ? 1 : 0) },
	};
}

/** The table's value-list filters in the auto-filter shape the row matcher understands. */
export function tableAutoFilter(table: Table): AutoFilter | undefined {
	const columns = table.columns.flatMap((column, offset) =>
		column.filter
			? [
					{
						offset,
						values: [...column.filter.values],
						...(column.filter.blank ? { blank: true } : {}),
					},
				]
			: [],
	);
	if (!columns.length) return undefined;
	return { range: tableFilterBounds(table), columns };
}

/** Whether a data row passes every value-list filter of the table. */
export function tableRowVisible(
	workbook: Workbook,
	sheet: Worksheet,
	table: Table,
	row: number,
): boolean {
	const filter = tableAutoFilter(table);
	return !filter || rowMatchesFilter(workbook, sheet, row, filter);
}

/** Shows or hides the table's data rows to match its column filters. */
export function applyTableFilter(workbook: Workbook, sheet: Worksheet, table: Table): void {
	if (!table.headerRow) return;
	const bounds = tableFilterBounds(table);
	const filter = tableAutoFilter(table);
	for (let row = bounds.start.row + 1; row <= bounds.end.row; row++) {
		const visible = !filter || rowMatchesFilter(workbook, sheet, row, filter);
		setRowFiltered(sheet, row, !visible, !!filter);
	}
}

/**
 * Loaded rows carry one combined hidden bit: mark the rows a table filter hides so a later
 * slicer click restores them without un-hiding rows the user hid by hand.
 */
export function readTableFilteredRows(workbook: Workbook): void {
	for (const sheet of workbook.sheets)
		for (const table of sheet.tables) {
			const filter = tableAutoFilter(table);
			if (!filter) continue;
			for (let row = filter.range.start.row + 1; row <= filter.range.end.row; row++) {
				const info = sheet.rowInfo.get(row);
				if (info?.hidden && info.filteredOut === undefined)
					info.filteredOut = !rowMatchesFilter(workbook, sheet, row, filter);
			}
		}
}

/** The display texts of a table column's data rows, in row order (`''` for blanks). */
export function tableColumnTexts(
	workbook: Workbook,
	sheet: Worksheet,
	table: Table,
	offset: number,
) {
	const bounds = tableFilterBounds(table);
	const texts: { row: number; text: string }[] = [];
	for (let row = bounds.start.row + 1; row <= bounds.end.row; row++)
		texts.push({
			row,
			text: displayText(workbook, getCell(sheet, row, bounds.start.col + offset)),
		});
	return texts;
}

/** `<filterColumn>` for a value-list filter, as Excel writes it. */
export function filterColumnXml(offset: number, filter: TableColumnFilter): string {
	const values = filter.values.map((value) => `<filter val="${escapeAttr(value)}"/>`).join('');
	return `<filterColumn colId="${offset}"><filters${filter.blank ? ' blank="1"' : ''}>${values}</filters></filterColumn>`;
}
