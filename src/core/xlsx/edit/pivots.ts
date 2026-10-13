import {
	MAX_COL,
	MAX_ROW,
	rangesIntersect,
	rangeContains,
	parseAddress,
	type CellAddress,
	type CellRange,
} from '../address';
import { forEachCellInRange, putCell, deleteCell } from '../cells';
import type { PivotTable, PivotTableOptions, PivotTablePatch, Workbook } from '../model';
import { pivotGrid } from '../pivot/aggregate';
import { pivotCaption, pivotSourceData, validatePivotFields } from '../pivot/data';
import { sparklineDataRange } from './sparkline-ranges';
import { type EditContext, sheetAt } from './context';
import { stylePatcher } from './format';
import { styleAt } from '../styles';
import { normalizedPivotMeasures } from '../pivot/measures';

function pivotAt(ctx: EditContext, sheet: number, name: string): PivotTable {
	const pivot = sheetAt(ctx.workbook, sheet).pivotTables?.find((entry) => entry.name === name);
	if (!pivot) throw new Error('The selected PivotTable is not available.');
	return pivot;
}

function uniqueName(ctx: EditContext): string {
	const used = reservedPivotNames(ctx.workbook);
	for (let i = 1; ; i++) if (!used.has(`pivottable${i}`)) return `PivotTable${i}`;
}

export function reservedPivotNames(workbook: Workbook): Set<string> {
	return new Set(
		workbook.sheets
			.flatMap((sheet) => [
				...(sheet.pivotTables ?? []).map((pivot) => pivot.name),
				...(sheet.preserved.get('source:unsupportedPivotNames') ?? []),
			])
			.map((name) => name.toLowerCase()),
	);
}

function checkName(ctx: EditContext, name: string): void {
	if (
		!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ||
		name.length > 255 ||
		parseAddress(name) ||
		reservedPivotNames(ctx.workbook).has(name.toLowerCase())
	)
		throw new Error('A PivotTable name must be unique and cannot look like a cell reference.');
}

function plan(ctx: EditContext, s: number, pivot: PivotTable, old?: PivotTable) {
	const sheet = sheetAt(ctx.workbook, s);
	const grid = pivotGrid(pivot);
	const start = pivot.location.start;
	const location = {
		start,
		end: {
			row: start.row + grid.values.length - 1,
			col: start.col + (grid.values[0]?.length ?? 0) - 1,
		},
	};
	if (
		![start.row, start.col].every(Number.isInteger) ||
		start.row < 0 ||
		start.col < 0 ||
		location.end.row > MAX_ROW ||
		location.end.col > MAX_COL
	)
		throw new Error('The PivotTable result does not fit on the worksheet.');
	const source = sparklineDataRange(ctx.workbook, s, pivot.source);
	if (
		(source.sheet === s && rangesIntersect(source.range, location)) ||
		sheet.merges.some((range) => rangesIntersect(range, location)) ||
		sheet.tables.some((table) => rangesIntersect(table.range, location)) ||
		sheet.pivotTables?.some((other) => other !== old && rangesIntersect(other.location, location))
	)
		throw new Error('The PivotTable would overlap source data, another table or merged cells.');
	let occupied = false;
	forEachCellInRange(sheet, location, (cell, row, col) => {
		if (
			(!old || !rangeContains(old.location, { row, col })) &&
			(cell.value !== null || cell.formula !== undefined)
		)
			occupied = true;
	});
	if (occupied) throw new Error('The PivotTable would overwrite existing cell content.');
	return { grid, location };
}

function write(ctx: EditContext, s: number, pivot: PivotTable, old?: PivotTable): void {
	const sheet = sheetAt(ctx.workbook, s);
	const { grid, location } = plan(ctx, s, pivot, old);
	const ranges = [...(old ? [old.location] : []), location];
	ctx.run(
		old ? 'Refresh PivotTable' : 'Create PivotTable',
		'annotations',
		[
			{ kind: 'parts', sheet: s, parts: ['pivotTables'] },
			{ kind: 'cells', sheet: s, ranges },
		],
		() => {
			const compact = pivot.layout === 'compact';
			const indents = new Map(
				[...new Set([0, ...(grid.indents ?? [])])].map(
					(indent) =>
						[
							indent,
							stylePatcher(ctx.workbook, {
								alignment: { indent, ...(compact || indent > 0 ? { horizontal: 'left' } : {}) },
							}),
						] as const,
				),
			);
			const styles = new Map<string, number>();
			forEachCellInRange(sheet, location, (cell, row, col) => {
				if (cell.styleId !== undefined) styles.set(`${row}:${col}`, cell.styleId);
			});
			if (old)
				forEachCellInRange(sheet, old.location, (cell, row, col) => {
					if (cell.styleId !== undefined)
						putCell(sheet, row, col, { value: null, styleId: cell.styleId });
					else deleteCell(sheet, row, col);
				});
			const next = { ...pivot, location, dirty: true };
			const list = sheet.pivotTables ?? [];
			sheet.pivotTables = old
				? list.map((entry) => (entry === old ? next : entry))
				: [...list, next];
			grid.values.forEach((line, row) =>
				line.forEach((value, col) => {
					const r = startRow(location, row);
					const c = location.start.col + col;
					let styleId = styles.get(`${r}:${c}`);
					if (col < (grid.rowLabelColumns ?? 1)) {
						const indent = col === 0 ? (grid.indents?.[row] ?? 0) : 0;
						const alignment = styleAt(ctx.workbook, styleId).alignment;
						if (
							(alignment?.indent ?? 0) !== indent ||
							(compact && alignment?.horizontal !== 'left')
						)
							styleId = indents.get(indent)!(styleId);
					}
					putCell(sheet, r, c, { value, ...(styleId !== undefined ? { styleId } : {}) });
				}),
			);
		},
		{ sheet: s, ranges },
	);
}
const startRow = (range: CellRange, offset: number) => range.start.row + offset;

export function createPivotTable(
	ctx: EditContext,
	s: number,
	destination: CellAddress,
	options: PivotTableOptions,
): string {
	const data = pivotSourceData(ctx.workbook, s, options.source);
	const name = options.name ?? uniqueName(ctx);
	checkName(ctx, name);
	const aggregate = options.aggregate ?? 'sum';
	const pivot: PivotTable = {
		...data,
		name,
		rowField: options.rowField,
		layout: options.layout ?? 'tabular',
		subtotalPosition:
			options.subtotalPosition ??
			(options.layout && options.layout !== 'tabular' ? 'top' : 'bottom'),
		...(options.secondRowField ? { secondRowField: options.secondRowField } : {}),
		...(options.subtotals !== undefined ? { subtotals: options.subtotals } : {}),
		...(options.repeatLabels !== undefined ? { repeatLabels: options.repeatLabels } : {}),
		valueField: options.valueField,
		aggregate,
		valueCaption: pivotCaption(aggregate, options.valueField),
		...(options.pageFilter !== undefined
			? { pageFilter: structuredClone(options.pageFilter) }
			: {}),
		...(options.valuesAxis !== undefined ? { valuesAxis: options.valuesAxis } : {}),
		...(options.valuesPosition !== undefined ? { valuesPosition: options.valuesPosition } : {}),
		...(options.additionalValues
			? { additionalValues: normalizedPivotMeasures(options.additionalValues) }
			: {}),
		location: { start: { ...destination }, end: { ...destination } },
		...(options.columnField ? { columnField: options.columnField } : {}),
	};
	validatePivotFields(pivot);
	write(ctx, s, pivot);
	return name;
}

export function refreshPivotTable(
	ctx: EditContext,
	s: number,
	name: string,
	patch: PivotTablePatch = {},
): void {
	const old = pivotAt(ctx, s, name);
	const aggregate = patch.aggregate ?? old.aggregate;
	const valueField = patch.valueField ?? old.valueField;
	const next: PivotTable = {
		...old,
		...pivotSourceData(ctx.workbook, s, old.source),
		rowField: patch.rowField ?? old.rowField,
		valueField,
		aggregate,
		valueCaption:
			aggregate === old.aggregate && valueField === old.valueField
				? old.valueCaption
				: pivotCaption(aggregate, valueField),
	};
	if (patch.columnField !== undefined) {
		if (patch.columnField) next.columnField = patch.columnField;
		else delete next.columnField;
	}
	if (patch.secondRowField !== undefined) {
		if (patch.secondRowField) next.secondRowField = patch.secondRowField;
		else delete next.secondRowField;
	}
	if (patch.subtotals !== undefined) next.subtotals = patch.subtotals;
	if (patch.repeatLabels !== undefined) next.repeatLabels = patch.repeatLabels;
	if (patch.layout !== undefined) next.layout = patch.layout;
	if (patch.subtotalPosition !== undefined) next.subtotalPosition = patch.subtotalPosition;
	if (patch.valuesAxis !== undefined) next.valuesAxis = patch.valuesAxis;
	if (patch.pageFilter !== undefined) {
		if (patch.pageFilter === null) delete next.pageFilter;
		else next.pageFilter = structuredClone(patch.pageFilter);
	}
	if (patch.valuesPosition !== undefined) next.valuesPosition = patch.valuesPosition;
	else if (patch.valuesAxis === 'columns' || patch.secondRowField === '')
		delete next.valuesPosition;
	if (patch.additionalValues !== undefined) {
		const measures = normalizedPivotMeasures(patch.additionalValues);
		if (measures.length) next.additionalValues = measures;
		else delete next.additionalValues;
	}
	validatePivotFields(next);
	write(ctx, s, next, old);
}

export function deletePivotTable(ctx: EditContext, s: number, name: string): void {
	const sheet = sheetAt(ctx.workbook, s);
	const pivot = pivotAt(ctx, s, name);
	ctx.run(
		'Delete PivotTable',
		'annotations',
		[
			{ kind: 'parts', sheet: s, parts: ['pivotTables'] },
			{ kind: 'cells', sheet: s, ranges: [pivot.location] },
		],
		() => {
			sheet.pivotTables = sheet.pivotTables!.filter((entry) => entry !== pivot);
			forEachCellInRange(sheet, pivot.location, (_cell, row, col) => deleteCell(sheet, row, col));
		},
		{ sheet: s, ranges: [pivot.location] },
	);
}
