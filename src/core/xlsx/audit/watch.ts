// Watch Window rows: the cells a user pins to see their value and formula while working elsewhere.
import { formatAddress } from '../address';
import { type CellValue, type Workbook } from '../model';
import { isCellError } from '../model';
import type { CellRef } from './trace';

export interface WatchRow {
	sheet: string;
	/** A defined name that refers to exactly this cell, if any. */
	name: string;
	cell: string;
	value: CellValue;
	formula: string;
}

/** Whether a watch points at a sheet that still exists. */
export const watchExists = (workbook: Workbook, w: CellRef): boolean => !!workbook.sheets[w.sheet];

function nameOf(workbook: Workbook, w: CellRef): string {
	const sheetName = workbook.sheets[w.sheet]?.name ?? '';
	const a1 = formatAddress({ row: w.row, col: w.col });
	for (const n of workbook.definedNames) {
		if (n.hidden || n.name.startsWith('_xlnm.')) continue;
		if (n.localSheet !== undefined && n.localSheet !== w.sheet) continue;
		const m = /^(?:'((?:[^']|'')+)'|([^!']+))!\$?([A-Z]{1,3})\$?(\d+)$/i.exec(n.formula.trim());
		if (!m) continue;
		const sheet = (m[1] ?? m[2] ?? '').replace(/''/g, "'");
		if (sheet.toLowerCase() === sheetName.toLowerCase() && `${m[3]}${m[4]}`.toUpperCase() === a1)
			return n.name;
	}
	return '';
}

/** The current state of one watched cell. */
export function describeWatch(workbook: Workbook, w: CellRef): WatchRow {
	const cell = workbook.sheets[w.sheet]?.rows.get(w.row)?.get(w.col);
	return {
		sheet: workbook.sheets[w.sheet]?.name ?? '',
		name: nameOf(workbook, w),
		cell: formatAddress({ row: w.row, col: w.col }),
		value: cell?.value ?? null,
		formula: cell?.formula ? `=${cell.formula}` : '',
	};
}

/** Value text for the grid of the window (errors by code, empty for blank). */
export function watchValueText(value: CellValue): string {
	if (value === null) return '';
	if (isCellError(value)) return value.error;
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	return String(value);
}
