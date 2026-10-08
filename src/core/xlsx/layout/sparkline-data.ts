// Finds the sparkline drawn in a cell and reads its data range from the workbook, then hands both
// to the pure layout in `sparkline-view.ts`. Cell values are read as stored (formula cells by their
// calculated value), so the view is current whenever the grid rebuilds its cell views.
import { cellKey, normalizeRange, parseRange, type CellRange } from '../address';
import { getCell } from '../cells';
import type { SparklineGroup, Workbook, Worksheet } from '../model';
import { sheetSparklineGroups } from '../read/sparklines';
import { sheetByName } from '../workbook';
import {
	sparklineScale,
	sparklineView,
	type SparklineData,
	type SparklineView,
} from './sparkline-view';

/** The most points one sparkline reads; a longer range (a whole column) stops at the last stored row. */
const MAX_POINTS = 10_000;

interface HostEntry {
	group: SparklineGroup;
	index: number;
}

const hostIndex = new WeakMap<readonly SparklineGroup[], Map<number, HostEntry>>();

function hosts(sheet: Worksheet): Map<number, HostEntry> {
	const groups = sheetSparklineGroups(sheet);
	let index = hostIndex.get(groups);
	if (!index) {
		index = new Map();
		for (const group of groups)
			group.sparklines.forEach((line, i) => {
				const key = cellKey(line.host.start.row, line.host.start.col);
				if (!index!.has(key)) index!.set(key, { group, index: i });
			});
		hostIndex.set(groups, index);
	}
	return index;
}

/** True when a sparkline is drawn in the cell. */
export function hasSparklineAt(sheet: Worksheet, row: number, col: number): boolean {
	return hosts(sheet).has(cellKey(row, col));
}

/** Resolves `Sheet1!$A$1:$E$1`, `'My sheet'!A1:E1` or an unqualified range on `sheet`. */
export function resolveSparklineRange(
	workbook: Workbook,
	sheet: Worksheet,
	formula: string,
): { sheet: Worksheet; range: CellRange } | undefined {
	const ref = formula.trim().replace(/^=/, '');
	const bang = ref.lastIndexOf('!');
	let target: Worksheet | undefined = sheet;
	if (bang >= 0) {
		const raw = ref.slice(0, bang);
		const name = /^'.*'$/.test(raw) ? raw.slice(1, -1).replace(/''/g, "'") : raw;
		target = sheetByName(workbook, name);
	}
	const range = parseRange(ref.slice(bang + 1).replace(/\$/g, ''));
	return target && range ? { sheet: target, range: normalizeRange(range) } : undefined;
}

const rowHidden = (sheet: Worksheet, row: number): boolean =>
	sheet.rowInfo.get(row)?.hidden === true;
const colHidden = (sheet: Worksheet, col: number): boolean =>
	sheet.columns.some((c) => c.hidden === true && col >= c.min && col <= c.max);

/**
 * The numbers of a one-dimensional range, along its row or column (a two-dimensional range is read
 * row by row). Empty and non-numeric cells are null; hidden rows and columns are skipped unless
 * `displayHidden`.
 */
export function readSparklineValues(
	sheet: Worksheet,
	range: CellRange,
	displayHidden: boolean,
): (number | null)[] {
	let endRow = range.end.row;
	if (endRow - range.start.row >= MAX_POINTS) {
		let lastRow = range.start.row;
		for (const row of sheet.rows.keys()) if (row > lastRow) lastRow = row;
		endRow = Math.min(endRow, lastRow);
	}
	const out: (number | null)[] = [];
	for (let row = range.start.row; row <= endRow && out.length < MAX_POINTS; row++) {
		if (!displayHidden && rowHidden(sheet, row)) continue;
		for (let col = range.start.col; col <= range.end.col && out.length < MAX_POINTS; col++) {
			if (!displayHidden && colHidden(sheet, col)) continue;
			const value = getCell(sheet, row, col)?.value;
			out.push(typeof value === 'number' && Number.isFinite(value) ? value : null);
		}
	}
	return out;
}

function groupData(workbook: Workbook, sheet: Worksheet, group: SparklineGroup): SparklineData[] {
	const dateRange = group.dateAxis && group.dateFormula;
	const dates = dateRange ? resolveSparklineRange(workbook, sheet, dateRange) : undefined;
	const dateValues = dates
		? readSparklineValues(dates.sheet, dates.range, group.displayHidden)
		: undefined;
	return group.sparklines.map((line) => {
		const source = line.formula ? resolveSparklineRange(workbook, sheet, line.formula) : undefined;
		const values = source
			? readSparklineValues(source.sheet, source.range, group.displayHidden)
			: [];
		return dateValues ? { values, dates: dateValues } : { values };
	});
}

/** The paint descriptor of the sparkline hosted in a cell, if any. */
export function sparklineViewAt(
	workbook: Workbook,
	sheetIndex: number,
	row: number,
	col: number,
): SparklineView | undefined {
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) return undefined;
	const entry = hosts(sheet).get(cellKey(row, col));
	if (!entry) return undefined;
	const { group, index } = entry;
	const needsGroup = group.minAxisType === 'group' || group.maxAxisType === 'group';
	const data = needsGroup
		? groupData(workbook, sheet, group)
		: groupData(workbook, sheet, { ...group, sparklines: [group.sparklines[index]!] });
	const at = needsGroup ? index : 0;
	return sparklineView(
		group,
		data[at] ?? { values: [] },
		sparklineScale(group, data, at),
		workbook.theme,
	);
}
