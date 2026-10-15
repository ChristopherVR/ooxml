import { getCell } from '../cells';
import { sparklineDataRange } from '../edit/sparkline-ranges';
import { serialToYmd, ymdToSerial } from '../formula/date-serial';
import type { PivotTable, Workbook } from '../model';
import { isDateFormat } from '../numfmt/index';
import { styleAt } from '../styles';

const pad = (value: number, width = 2) => String(value).padStart(width, '0');

/** `2025-02-01T00:00:00` (the form pivot caches and timelines store) from a date serial. */
export function serialToIso(serial: number, date1904: boolean): string {
	const { year, month, day } = serialToYmd(serial, date1904);
	const seconds = Math.round((serial - Math.floor(serial)) * 86_400);
	const [h, m, s] = [Math.floor(seconds / 3600), Math.floor((seconds % 3600) / 60), seconds % 60];
	return `${pad(year, 4)}-${pad(month)}-${pad(day)}T${pad(h)}:${pad(m)}:${pad(s)}`;
}

/** The serial of an ISO date or date-time, or undefined when it is not one. */
export function isoToSerial(iso: string, date1904: boolean): number | undefined {
	const match = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?Z?)?$/.exec(
		iso.trim(),
	);
	if (!match) return undefined;
	const [year, month, day, h, m, s] = match.slice(1).map((part) => Number(part ?? 0));
	if (month! < 1 || month! > 12 || day! < 1 || day! > 31) return undefined;
	const base = ymdToSerial(year!, month!, day!, date1904);
	return base === undefined
		? undefined
		: base + ((h ?? 0) * 3600 + (m ?? 0) * 60 + (s ?? 0)) / 86_400;
}

/**
 * Source columns that hold only dates (numbers shown with a date format, no blanks), which a
 * timeline can filter. `skip` names fields that stay ordinary (those on an axis or measuring).
 */
export function pivotDateColumns(
	workbook: Workbook,
	sheet: number,
	source: string,
	fields: readonly string[],
	skip: ReadonlySet<string>,
): string[] {
	const data = sparklineDataRange(workbook, sheet, source);
	const ws = workbook.sheets[data.sheet]!;
	const { start, end } = data.range;
	const dates: string[] = [];
	fields.forEach((name, index) => {
		if (skip.has(name)) return;
		const col = start.col + index;
		let rows = 0;
		for (let row = start.row + 1; row <= end.row; row++) {
			const cell = getCell(ws, row, col);
			if (
				typeof cell?.value !== 'number' ||
				!Number.isInteger(cell.value) ||
				cell.value < 1 ||
				!isDateFormat(styleAt(workbook, cell.styleId).numFmt ?? 'General')
			)
				return;
			rows++;
		}
		if (rows) dates.push(name);
	});
	return dates;
}

/** The date fields a PivotTable keeps as dates: date-only source columns on no axis or measure. */
export function pivotDateFieldsOf(
	workbook: Workbook,
	sheet: number,
	pivot: Pick<
		PivotTable,
		| 'source'
		| 'fields'
		| 'rowField'
		| 'secondRowField'
		| 'columnField'
		| 'pageFilter'
		| 'valueField'
		| 'additionalValues'
	>,
): string[] {
	const used = new Set<string>(
		[
			pivot.rowField,
			pivot.secondRowField,
			pivot.columnField,
			pivot.pageFilter?.field,
			pivot.valueField,
			...(pivot.additionalValues ?? []).map((measure) => measure.field),
		].filter((name): name is string => !!name),
	);
	return pivotDateColumns(workbook, sheet, pivot.source, pivot.fields, used);
}
